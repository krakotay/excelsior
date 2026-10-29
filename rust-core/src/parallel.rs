//! Bounded parallel compression of ZIP members into disk-backed mini archives.
//! The final archive receives already-compressed entries in submission order.
use anyhow::{Context, Result, ensure};
use std::{
    collections::VecDeque,
    fs::File,
    io::{Seek, Write},
    path::{Path, PathBuf},
    thread::{self, JoinHandle},
};
use tempfile::NamedTempFile;
use zip::{ZipArchive, ZipWriter};

type Dimensions = (String, (usize, u32));
struct Prepared {
    file: NamedTempFile,
    dimensions: Option<Dimensions>,
}
pub(crate) struct Compressor {
    directory: PathBuf,
    workers: usize,
    pending: VecDeque<JoinHandle<Result<Prepared>>>,
    dimensions: Vec<Dimensions>,
}
impl Compressor {
    pub fn new(directory: &Path, workers: usize) -> Result<Self> {
        ensure!(workers > 0, "compression_threads must be positive");
        Ok(Self {
            directory: directory.to_owned(),
            workers,
            pending: VecDeque::new(),
            dimensions: Vec::new(),
        })
    }
    pub fn submit<W, F>(&mut self, output: &mut ZipWriter<W>, job: F) -> Result<()>
    where
        W: Write + Seek,
        F: FnOnce(&mut ZipWriter<File>) -> Result<Option<Dimensions>> + Send + 'static,
    {
        if self.pending.len() >= self.workers {
            self.drain_one(output)?;
        }
        let file = tempfile::Builder::new()
            .prefix(".excelsior-part-")
            .suffix(".zip")
            .tempfile_in(&self.directory)?;
        let handle = thread::Builder::new()
            .name("excelsior-zip".into())
            .spawn(move || {
                let mut zip = ZipWriter::new(file.reopen()?);
                let dimensions = job(&mut zip)?;
                zip.finish()?;
                Ok(Prepared { file, dimensions })
            })?;
        self.pending.push_back(handle);
        Ok(())
    }
    fn drain_one<W: Write + Seek>(&mut self, output: &mut ZipWriter<W>) -> Result<()> {
        let prepared = self
            .pending
            .pop_front()
            .context("No pending ZIP member")?
            .join()
            .map_err(|_| anyhow::anyhow!("ZIP compression worker panicked"))??;
        let mut input = ZipArchive::new(prepared.file.reopen()?)?;
        output.raw_copy_file(input.by_index(0)?)?;
        if let Some(dimensions) = prepared.dimensions {
            self.dimensions.push(dimensions);
        }
        Ok(())
    }
    pub fn finish<W: Write + Seek>(mut self, output: &mut ZipWriter<W>) -> Result<Vec<Dimensions>> {
        while !self.pending.is_empty() {
            self.drain_one(output)?;
        }
        Ok(std::mem::take(&mut self.dimensions))
    }
}
impl Drop for Compressor {
    fn drop(&mut self) {
        // Join even on an error/abort so worker files are gone before returning.
        for handle in self.pending.drain(..) {
            let _ = handle.join();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::{
        io::{Cursor, Read},
        sync::{Arc, Barrier},
    };
    #[test]
    fn members_compress_concurrently_and_keep_submission_order() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let mut output = ZipWriter::new(Cursor::new(Vec::new()));
        let mut compressor = Compressor::new(directory.path(), 2)?;
        let barrier = Arc::new(Barrier::new(2));
        for name in ["one", "two"] {
            let barrier = barrier.clone();
            compressor.submit(&mut output, move |zip| {
                barrier.wait();
                zip.start_file(name, crate::files_part::compression_options(3))?;
                zip.write_all(name.as_bytes())?;
                Ok(None)
            })?;
        }
        compressor.finish(&mut output)?;
        let mut archive = ZipArchive::new(output.finish()?)?;
        for (i, name) in ["one", "two"].into_iter().enumerate() {
            let mut file = archive.by_index(i)?;
            assert_eq!(file.name()?, name);
            let mut bytes = String::new();
            file.read_to_string(&mut bytes)?;
            assert_eq!(bytes, name);
        }
        assert_eq!(std::fs::read_dir(directory.path())?.count(), 0);
        Ok(())
    }
    #[test]
    fn failed_member_joins_other_workers_and_removes_temporary_files() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let mut output = ZipWriter::new(Cursor::new(Vec::new()));
        let mut compressor = Compressor::new(directory.path(), 2)?;
        let barrier = Arc::new(Barrier::new(2));
        let first = barrier.clone();
        compressor.submit(&mut output, move |_| {
            first.wait();
            anyhow::bail!("deliberate worker failure")
        })?;
        compressor.submit(&mut output, move |zip| {
            barrier.wait();
            zip.start_file("valid", crate::files_part::compression_options(1))?;
            zip.write_all(b"valid")?;
            Ok(None)
        })?;
        assert!(compressor.finish(&mut output).is_err());
        assert_eq!(std::fs::read_dir(directory.path())?.count(), 0);
        Ok(())
    }
}
