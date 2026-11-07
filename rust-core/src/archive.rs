//! Independent ZIP cursors sharing a disk-backed source snapshot. Unix keeps the
//! original file open; Windows uses a temporary copy so MoveFileEx can replace
//! the original path while Editor branches retain their original archive.
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
        #[cfg(windows)]
        let file = {
            // MoveFileEx cannot replace an open destination, even when its
            // handles allow delete sharing. Do not hold a published path open.
            // A disk snapshot also retains lazy ZIP parts across branch saves
            // without buffering the workbook in memory or copying each branch.
            let mut original = file;
            original.rewind()?;
            let mut snapshot = tempfile::tempfile()?;
            io::copy(&mut original, &mut snapshot)?;
            snapshot
        };
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn readers_keep_snapshot_after_atomic_replacement() -> io::Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("source.xlsx");
        std::fs::write(&path, b"original archive")?;
        let source = Source::open(&path)?;
        let branch = source.clone();
        let mut cursor = branch.reader();
        cursor.seek(SeekFrom::Start(9))?;
        let mut replacement = tempfile::NamedTempFile::new_in(directory.path())?;
        std::io::Write::write_all(&mut replacement, b"replacement archive")?;
        let published = replacement.persist(&path).map_err(|e| e.error)?;
        let published = Source::from_file(published)?;
        let mut original = String::new();
        source.reader().read_to_string(&mut original)?;
        assert_eq!(original, "original archive");
        let mut remainder = String::new();
        cursor.read_to_string(&mut remainder)?;
        assert_eq!(remainder, "archive");
        let mut current = String::new();
        published.reader().read_to_string(&mut current)?;
        assert_eq!(current, "replacement archive");
        std::fs::remove_file(&path)?;
        Ok(())
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
