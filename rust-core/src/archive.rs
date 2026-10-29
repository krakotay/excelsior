//! Independent ZIP cursors sharing the same open source file. Atomic saves keep
//! that original file alive for all Editor branches without copying it into RAM.
use std::{
    fs::File,
    io::{self, Read, Seek, SeekFrom},
    path::Path,
    sync::{Arc, Mutex},
};
#[derive(Clone)]
pub(crate) struct Source {
    file: Arc<Mutex<File>>,
    length: u64,
}
impl Source {
    pub fn open(path: &Path) -> io::Result<Self> {
        Self::from_file(File::open(path)?)
    }
    pub fn from_file(file: File) -> io::Result<Self> {
        let length = file.metadata()?.len();
        Ok(Self {
            file: Arc::new(Mutex::new(file)),
            length,
        })
    }
    pub fn reader(&self) -> Cursor {
        Cursor {
            source: self.clone(),
            position: 0,
        }
    }
}
pub(crate) struct Cursor {
    source: Source,
    position: u64,
}
impl Read for Cursor {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let mut file = self
            .source
            .file
            .lock()
            .map_err(|_| io::Error::other("Source file lock poisoned"))?;
        file.seek(SeekFrom::Start(self.position))?;
        let read = file.read(output)?;
        self.position += read as u64;
        Ok(read)
    }
}
impl Seek for Cursor {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let position = match position {
            SeekFrom::Start(n) => n as i128,
            SeekFrom::Current(n) => self.position as i128 + n as i128,
            SeekFrom::End(n) => self.source.length as i128 + n as i128,
        };
        self.position = u64::try_from(position)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "Invalid source seek"))?;
        Ok(self.position)
    }
}
