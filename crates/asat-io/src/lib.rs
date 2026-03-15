pub mod csv_driver;
pub mod xlsx_driver;
pub mod ods_driver;
pub mod asat_driver;

use std::path::Path;
use asat_core::Workbook;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum IoError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("CSV error: {0}")]
    Csv(String),
    #[error("XLSX error: {0}")]
    Xlsx(String),
    #[error("ODS error: {0}")]
    Ods(String),
    #[error("unsupported format: {0}")]
    UnsupportedFormat(String),
    #[error("encode/decode error: {0}")]
    Codec(String),
}

pub trait FileDriver: Send + Sync {
    fn read(&self, path: &Path) -> Result<Workbook, IoError>;
    fn write(&self, workbook: &Workbook, path: &Path) -> Result<(), IoError>;
    fn extensions(&self) -> &[&str];
}

/// Load a workbook from any supported file path
pub fn load(path: &Path) -> Result<Workbook, IoError> {
    let ext = path.extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    match ext.as_str() {
        "csv" | "tsv" => csv_driver::CsvDriver.read(path),
        "xlsx" | "xls" | "xlsm" => xlsx_driver::XlsxDriver.read(path),
        "ods" => ods_driver::OdsDriver.read(path),
        "asat" => asat_driver::AsatDriver.read(path),
        _ => Err(IoError::UnsupportedFormat(ext)),
    }
}

/// Save a workbook to any supported file path
pub fn save(workbook: &Workbook, path: &Path) -> Result<(), IoError> {
    let ext = path.extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    match ext.as_str() {
        "csv" | "tsv" => csv_driver::CsvDriver.write(workbook, path),
        "xlsx" => xlsx_driver::XlsxDriver.write(workbook, path),
        "ods" => ods_driver::OdsDriver.write(workbook, path),
        "asat" => asat_driver::AsatDriver.write(workbook, path),
        _ => Err(IoError::UnsupportedFormat(ext)),
    }
}
