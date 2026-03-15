use crate::{FileDriver, IoError};
use asat_core::{Cell, CellValue, Sheet, Workbook};
use std::path::Path;

pub struct CsvDriver;

impl FileDriver for CsvDriver {
    fn read(&self, path: &Path) -> Result<Workbook, IoError> {
        let is_tsv = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.eq_ignore_ascii_case("tsv"))
            .unwrap_or(false);

        let delimiter = if is_tsv { b'\t' } else { b',' };

        let mut rdr = csv::ReaderBuilder::new()
            .has_headers(false)
            .delimiter(delimiter)
            .flexible(true)
            .from_path(path)
            .map_err(|e| IoError::Csv(e.to_string()))?;

        let sheet_name = path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("Sheet1")
            .to_string();

        let mut sheet = Sheet::new(sheet_name);

        for (row_idx, result) in rdr.records().enumerate() {
            let record = result.map_err(|e| IoError::Csv(e.to_string()))?;
            for (col_idx, field) in record.iter().enumerate() {
                if !field.is_empty() {
                    let value = parse_csv_value(field);
                    sheet.set_cell(row_idx as u32, col_idx as u32, Cell::new(value));
                }
            }
        }

        let mut wb = Workbook::new();
        wb.sheets[0] = sheet;
        wb.file_path = Some(path.to_path_buf());
        wb.dirty = false;
        Ok(wb)
    }

    fn write(&self, workbook: &Workbook, path: &Path) -> Result<(), IoError> {
        let is_tsv = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.eq_ignore_ascii_case("tsv"))
            .unwrap_or(false);

        let delimiter = if is_tsv { b'\t' } else { b',' };

        let sheet = workbook.active();

        if sheet.cells.is_empty() {
            // Write empty file
            std::fs::write(path, b"")?;
            return Ok(());
        }

        let max_row = sheet.max_row();
        let max_col = sheet.max_col();

        let mut wtr = csv::WriterBuilder::new()
            .delimiter(delimiter)
            .from_path(path)
            .map_err(|e| IoError::Csv(e.to_string()))?;

        for row in 0..=max_row {
            let record: Vec<String> = (0..=max_col)
                .map(|col| sheet.display_value(row, col))
                .collect();
            wtr.write_record(&record)
                .map_err(|e| IoError::Csv(e.to_string()))?;
        }

        wtr.flush().map_err(|e| IoError::Csv(e.to_string()))?;
        Ok(())
    }

    fn extensions(&self) -> &[&str] {
        &["csv", "tsv"]
    }
}

fn parse_csv_value(s: &str) -> CellValue {
    if s.is_empty() {
        return CellValue::Empty;
    }
    // Formula
    if s.starts_with('=') {
        return CellValue::Formula(s[1..].to_string());
    }
    // Number
    if let Ok(n) = s.parse::<f64>() {
        return CellValue::Number(n);
    }
    // Boolean
    match s.to_uppercase().as_str() {
        "TRUE" => return CellValue::Boolean(true),
        "FALSE" => return CellValue::Boolean(false),
        _ => {}
    }
    CellValue::Text(s.to_string())
}
