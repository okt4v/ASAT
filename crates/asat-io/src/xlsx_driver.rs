use std::path::Path;
use asat_core::{Cell, CellValue, Sheet, Workbook};
use calamine::{open_workbook_auto, Data, Reader};
use crate::{FileDriver, IoError};

pub struct XlsxDriver;

impl FileDriver for XlsxDriver {
    fn read(&self, path: &Path) -> Result<Workbook, IoError> {
        let mut calamine_wb: calamine::Sheets<_> = open_workbook_auto(path)
            .map_err(|e| IoError::Xlsx(e.to_string()))?;

        let mut wb = Workbook {
            sheets: Vec::new(),
            active_sheet: 0,
            file_path: Some(path.to_path_buf()),
            dirty: false,
            named_ranges: Default::default(),
        };

        let sheet_names: Vec<String> = calamine_wb.sheet_names().to_vec();

        for name in &sheet_names {
            let range = calamine_wb
                .worksheet_range(name)
                .map_err(|e| IoError::Xlsx(e.to_string()))?;

            let mut sheet = Sheet::new(name.clone());

            for (row_idx, row) in range.rows().enumerate() {
                for (col_idx, cell) in row.iter().enumerate() {
                    let value = calamine_data_to_cell_value(cell);
                    if !value.is_empty() {
                        sheet.set_cell(row_idx as u32, col_idx as u32, Cell::new(value));
                    }
                }
            }

            wb.sheets.push(sheet);
        }

        if wb.sheets.is_empty() {
            wb.sheets.push(Sheet::new("Sheet1"));
        }

        Ok(wb)
    }

    fn write(&self, workbook: &Workbook, path: &Path) -> Result<(), IoError> {
        use rust_xlsxwriter::{Formula, Workbook as XlWorkbook};

        let mut xl_wb = XlWorkbook::new();

        for sheet in workbook.sheets.iter() {
            let worksheet = xl_wb.add_worksheet();
            worksheet.set_name(&sheet.name)
                .map_err(|e| IoError::Xlsx(e.to_string()))?;

            let max_row = sheet.max_row();
            let max_col = sheet.max_col();

            for row in 0..=max_row {
                for col in 0..=max_col {
                    let value = sheet.get_value(row, col);
                    match value {
                        CellValue::Empty => {}
                        CellValue::Text(s) => {
                            worksheet.write_string(row as u32, col as u16, s)
                                .map_err(|e| IoError::Xlsx(e.to_string()))?;
                        }
                        CellValue::Number(n) => {
                            worksheet.write_number(row as u32, col as u16, *n)
                                .map_err(|e| IoError::Xlsx(e.to_string()))?;
                        }
                        CellValue::Boolean(b) => {
                            worksheet.write_boolean(row as u32, col as u16, *b)
                                .map_err(|e| IoError::Xlsx(e.to_string()))?;
                        }
                        CellValue::Formula(f) => {
                            let formula = Formula::new(format!("={}", f));
                            worksheet.write_formula(row as u32, col as u16, formula)
                                .map_err(|e| IoError::Xlsx(e.to_string()))?;
                        }
                        CellValue::Error(e) => {
                            worksheet.write_string(row as u32, col as u16, &e.to_string())
                                .map_err(|e2| IoError::Xlsx(e2.to_string()))?;
                        }
                    }
                }
            }
        }

        xl_wb.save(path).map_err(|e| IoError::Xlsx(e.to_string()))?;
        Ok(())
    }

    fn extensions(&self) -> &[&str] {
        &["xlsx", "xls", "xlsm"]
    }
}

fn calamine_data_to_cell_value(dt: &Data) -> CellValue {
    match dt {
        Data::Empty => CellValue::Empty,
        Data::String(s) => {
            if s.starts_with('=') {
                CellValue::Formula(s[1..].to_string())
            } else {
                CellValue::Text(s.clone())
            }
        }
        Data::Float(f) => CellValue::Number(*f),
        Data::Int(i)   => CellValue::Number(*i as f64),
        Data::Bool(b)  => CellValue::Boolean(*b),
        Data::Error(_) => CellValue::Error(asat_core::CellError::Value),
        Data::DateTime(dt) => CellValue::Number(dt.as_f64()),
        Data::DateTimeIso(s) => CellValue::Text(s.clone()),
        Data::DurationIso(s) => CellValue::Text(s.clone()),
    }
}
