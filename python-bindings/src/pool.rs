//! File-level scheduling lives in Rust. Python callbacks reuse the optional,
//! duck-typed dataframe adapter; their native writer calls release the GIL.
use pyo3::exceptions::PyBaseException;
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::{PyTraverseError, PyVisit};
use std::{
    cell::Cell,
    collections::HashSet,
    panic::{AssertUnwindSafe, catch_unwind},
    path::PathBuf,
    sync::{
        Arc, Condvar, Mutex,
        atomic::{AtomicU8, AtomicUsize, Ordering},
        mpsc::{SyncSender, sync_channel},
    },
    thread::{self, JoinHandle},
};

static NEXT_POOL: AtomicUsize = AtomicUsize::new(1);
thread_local! { static CURRENT_POOL: Cell<usize> = const { Cell::new(0) }; }

type Outcome = Arc<Option<Py<PyBaseException>>>;

struct Completion {
    pool: usize,
    // 0: queued, 1: running, 2: cancelled before starting.
    phase: AtomicU8,
    outcome: Mutex<Option<Outcome>>,
    ready: Condvar,
}
impl Completion {
    fn finish(&self, result: PyResult<()>) {
        // Normalize before locking: exception normalization may run Python.
        let error = result
            .err()
            .map(|error| Python::attach(|py| error.into_value(py)));
        *self.outcome.lock().unwrap() = Some(Arc::new(error));
        self.ready.notify_all();
    }
    fn wait(&self) -> Outcome {
        let mut outcome = self.outcome.lock().unwrap();
        while outcome.is_none() {
            outcome = self.ready.wait(outcome).unwrap();
        }
        Arc::clone(outcome.as_ref().unwrap())
    }
    fn cancel(&self) -> bool {
        self.phase
            .compare_exchange(0, 2, Ordering::AcqRel, Ordering::Acquire)
            .is_ok()
    }
    fn check_wait(&self) -> PyResult<()> {
        if CURRENT_POOL.get() == self.pool {
            return Err(PyRuntimeError::new_err(
                "Cannot wait for this pool from its own worker",
            ));
        }
        Ok(())
    }
}

#[pyclass(name = "_ExportTask", frozen)]
pub(crate) struct ExportTask {
    state: Arc<Completion>,
}
#[pymethods]
impl ExportTask {
    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        if let Some(result) = self.state.outcome.lock().unwrap().as_ref()
            && let Some(error) = result.as_ref()
        {
            visit.call(error)?;
        }
        Ok(())
    }
    fn __clear__(&self) {
        let result = self.state.outcome.lock().unwrap().take();
        drop(result);
    }
    fn done(&self) -> bool {
        self.state.outcome.lock().unwrap().is_some()
    }
    fn cancel(&self) -> bool {
        self.state.cancel()
    }
    fn wait(&self, py: Python<'_>) -> PyResult<()> {
        self.state.check_wait()?;
        // Never acquire the GIL while holding the completion mutex.
        let result = py.detach(|| self.state.wait());
        match result.as_ref() {
            Some(error) => Err(PyErr::from_value(error.bind(py).clone().into_any())),
            None => Ok(()),
        }
    }
}

struct Job {
    callback: Py<PyAny>,
    state: Arc<Completion>,
}
struct State {
    sender: Option<SyncSender<Job>>,
    workers: Vec<JoinHandle<()>>,
    tasks: Vec<Py<ExportTask>>,
    destinations: HashSet<PathBuf>,
}
#[pyclass(name = "_WorkbookPool", frozen)]
pub(crate) struct WorkbookPool {
    id: usize,
    state: Mutex<State>,
}
impl WorkbookPool {
    fn check_caller(&self) -> PyResult<()> {
        if CURRENT_POOL.get() == self.id {
            return Err(PyRuntimeError::new_err(
                "Cannot submit or wait on a pool from its own worker",
            ));
        }
        Ok(())
    }
}
#[pymethods]
impl WorkbookPool {
    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        for task in &self.state.lock().unwrap().tasks {
            visit.call(task)?;
        }
        Ok(())
    }
    fn __clear__(&self) {
        let tasks = std::mem::take(&mut self.state.lock().unwrap().tasks);
        drop(tasks);
    }
    #[new]
    fn new(workers: usize) -> PyResult<Self> {
        if workers == 0 {
            return Err(PyValueError::new_err("workers must be positive"));
        }
        let id = NEXT_POOL.fetch_add(1, Ordering::Relaxed);
        // Bound pending file jobs as well as active writers. Sending must detach
        // from Python: workers can need the GIL to make room in this queue.
        let (sender, receiver) = sync_channel::<Job>(workers);
        let receiver = Arc::new(Mutex::new(receiver));
        let mut state = State {
            sender: Some(sender),
            workers: Vec::new(),
            tasks: Vec::new(),
            destinations: HashSet::new(),
        };
        for index in 0..workers {
            let receiver = Arc::clone(&receiver);
            let handle = thread::Builder::new()
                .name(format!("excelsior-{id}-{index}"))
                .spawn(move || {
                    CURRENT_POOL.set(id);
                    loop {
                        let job = { receiver.lock().unwrap().recv() };
                        let Ok(job) = job else { break };
                        let result = if job
                            .state
                            .phase
                            .compare_exchange(0, 1, Ordering::AcqRel, Ordering::Acquire)
                            .is_ok()
                        {
                            catch_unwind(AssertUnwindSafe(|| {
                                Python::attach(|py| job.callback.bind(py).call0().map(|_| ()))
                            }))
                            .unwrap_or_else(|_| {
                                Err(PyRuntimeError::new_err("Workbook export worker panicked"))
                            })
                        } else {
                            Err(PyRuntimeError::new_err(
                                "Workbook export cancelled before starting",
                            ))
                        };
                        job.state.finish(result);
                    }
                    CURRENT_POOL.set(0);
                });
            match handle {
                Ok(handle) => state.workers.push(handle),
                Err(error) => {
                    state.sender.take();
                    for worker in state.workers {
                        let _ = worker.join();
                    }
                    return Err(PyRuntimeError::new_err(error.to_string()));
                }
            }
        }
        Ok(Self {
            id,
            state: Mutex::new(state),
        })
    }

    fn submit(
        &self,
        py: Python<'_>,
        path: PathBuf,
        callback: Py<PyAny>,
    ) -> PyResult<Py<ExportTask>> {
        self.check_caller()?;
        let completion = Arc::new(Completion {
            pool: self.id,
            phase: AtomicU8::new(0),
            outcome: Mutex::new(None),
            ready: Condvar::new(),
        });
        // Allocation may trigger GC/traversal; do it outside the state mutex.
        let task = Py::new(
            py,
            ExportTask {
                state: Arc::clone(&completion),
            },
        )?;
        let sender = {
            let mut state = self.state.lock().unwrap();
            let sender = state
                .sender
                .as_ref()
                .ok_or_else(|| PyRuntimeError::new_err("WorkbookPool is closed"))?
                .clone();
            if !state.destinations.insert(path) {
                return Err(PyValueError::new_err(
                    "Each destination can only be submitted once per WorkbookPool",
                ));
            }
            state.tasks.push(task.clone_ref(py));
            sender
        };
        let job = Job {
            callback,
            state: Arc::clone(&completion),
        };
        if py.detach(move || sender.send(job)).is_err() {
            completion.finish(Err(PyRuntimeError::new_err(
                "WorkbookPool workers have stopped",
            )));
        }
        Ok(task)
    }

    fn is_closed(&self) -> bool {
        self.state.lock().unwrap().sender.is_none()
    }

    #[pyo3(signature = (*, cancel_pending = false))]
    fn close(&self, py: Python<'_>, cancel_pending: bool) -> PyResult<()> {
        self.check_caller()?;
        let (tasks, workers) = {
            let mut state = self.state.lock().unwrap();
            state.sender.take();
            (
                state
                    .tasks
                    .iter()
                    .map(|task| Arc::clone(&task.get().state))
                    .collect::<Vec<_>>(),
                std::mem::take(&mut state.workers),
            )
        };
        if cancel_pending {
            for task in &tasks {
                task.cancel();
            }
        }
        let results = py.detach(move || {
            let results = tasks.iter().map(|task| task.wait()).collect::<Vec<_>>();
            for worker in workers {
                let _ = worker.join();
            }
            results
        });
        for result in results {
            if let Some(error) = result.as_ref() {
                return Err(PyErr::from_value(error.bind(py).clone().into_any()));
            }
        }
        Ok(())
    }
}
impl Drop for WorkbookPool {
    fn drop(&mut self) {
        let state = self.state.get_mut().unwrap();
        state.sender.take();
        for task in &state.tasks {
            task.get().state.cancel();
        }
        let workers = std::mem::take(&mut state.workers);
        // Explicit context exit is the normal path. Abandoned pools must not
        // join GIL-dependent workers while their Python wrapper is deallocated.
        Python::try_attach(move |py| {
            py.detach(move || {
                for worker in workers {
                    if worker.thread().id() != thread::current().id() {
                        let _ = worker.join();
                    }
                }
            })
        });
    }
}
