use crate::{FileDriver, IoError};
use asat_core::Workbook;
use std::path::Path;

const ASAT_MAGIC: &[u8; 4] = b"ASAT";
const ASAT_VERSION: u32 = 1;

pub struct AsatDriver;

impl FileDriver for AsatDriver {
    fn read(&self, path: &Path) -> Result<Workbook, IoError> {
        let compressed = std::fs::read(path)?;

        // Verify magic bytes + version
        if compressed.len() < 8 {
            return Err(IoError::Codec("file too short".into()));
        }
        if &compressed[..4] != ASAT_MAGIC {
            return Err(IoError::Codec("not an ASAT file".into()));
        }
        let version =
            u32::from_le_bytes([compressed[4], compressed[5], compressed[6], compressed[7]]);
        if version != ASAT_VERSION {
            return Err(IoError::Codec(format!(
                "unsupported ASAT version {}",
                version
            )));
        }

        let decompressed =
            zstd::decode_all(&compressed[8..]).map_err(|e| IoError::Codec(e.to_string()))?;

        let mut wb: Workbook =
            bincode::deserialize(&decompressed).map_err(|e| IoError::Codec(e.to_string()))?;

        wb.file_path = Some(path.to_path_buf());
        wb.dirty = false;
        Ok(wb)
    }

    fn write(&self, workbook: &Workbook, path: &Path) -> Result<(), IoError> {
        let serialized = bincode::serialize(workbook).map_err(|e| IoError::Codec(e.to_string()))?;

        let compressed = zstd::encode_all(serialized.as_slice(), 3)
            .map_err(|e| IoError::Codec(e.to_string()))?;

        let mut output = Vec::with_capacity(8 + compressed.len());
        output.extend_from_slice(ASAT_MAGIC);
        output.extend_from_slice(&ASAT_VERSION.to_le_bytes());
        output.extend_from_slice(&compressed);

        std::fs::write(path, &output)?;
        Ok(())
    }

    fn extensions(&self) -> &[&str] {
        &["asat"]
    }
}
