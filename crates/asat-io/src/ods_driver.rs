use std::io::Write;
use std::path::Path;

use asat_core::{Cell, CellValue, Workbook};
use calamine::{open_workbook_auto, Data, Reader};
use zip::{write::SimpleFileOptions, CompressionMethod, ZipWriter};

use crate::{FileDriver, IoError};

pub struct OdsDriver;

impl FileDriver for OdsDriver {
    fn read(&self, path: &Path) -> Result<Workbook, IoError> {
        // calamine's open_workbook_auto handles .ods natively
        let mut cal: calamine::Sheets<_> = open_workbook_auto(path)
            .map_err(|e| IoError::Ods(e.to_string()))?;

        let mut wb = Workbook {
            sheets: Vec::new(),
            active_sheet: 0,
            file_path: Some(path.to_path_buf()),
            dirty: false,
            named_ranges: Default::default(),
        };

        let sheet_names: Vec<String> = cal.sheet_names().to_vec();

        for name in &sheet_names {
            let range = cal
                .worksheet_range(name)
                .map_err(|e| IoError::Ods(e.to_string()))?;

            let mut sheet = asat_core::Sheet::new(name.clone());

            for (row_idx, row) in range.rows().enumerate() {
                for (col_idx, cell) in row.iter().enumerate() {
                    let value = data_to_value(cell);
                    if !value.is_empty() {
                        sheet.set_cell(row_idx as u32, col_idx as u32, Cell::new(value));
                    }
                }
            }

            wb.sheets.push(sheet);
        }

        if wb.sheets.is_empty() {
            wb.sheets.push(asat_core::Sheet::new("Sheet1"));
        }

        Ok(wb)
    }

    fn write(&self, workbook: &Workbook, path: &Path) -> Result<(), IoError> {
        let file = std::fs::File::create(path).map_err(IoError::Io)?;
        let mut zip = ZipWriter::new(file);

        // ── mimetype — must be FIRST, stored (not compressed) ──────────
        zip.start_file(
            "mimetype",
            SimpleFileOptions::default().compression_method(CompressionMethod::Stored),
        )
        .map_err(|e| IoError::Ods(e.to_string()))?;
        zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .map_err(|e| IoError::Ods(e.to_string()))?;

        let deflate = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

        // ── META-INF/manifest.xml ───────────────────────────────────────
        zip.start_file("META-INF/manifest.xml", deflate)
            .map_err(|e| IoError::Ods(e.to_string()))?;
        zip.write_all(MANIFEST_XML.as_bytes())
            .map_err(|e| IoError::Ods(e.to_string()))?;

        // ── content.xml ────────────────────────────────────────────────
        let content = build_content_xml(workbook);
        zip.start_file("content.xml", deflate)
            .map_err(|e| IoError::Ods(e.to_string()))?;
        zip.write_all(content.as_bytes())
            .map_err(|e| IoError::Ods(e.to_string()))?;

        zip.finish().map_err(|e| IoError::Ods(e.to_string()))?;
        Ok(())
    }

    fn extensions(&self) -> &[&str] {
        &["ods"]
    }
}

// ── XML helpers ───────────────────────────────────────────────────────────────

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

// ── Content builder ───────────────────────────────────────────────────────────

fn build_content_xml(workbook: &Workbook) -> String {
    let mut xml = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\
         <office:document-content\
           xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"\
           xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"\
           xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"\
           xmlns:of=\"urn:oasis:names:tc:opendocument:xmlns:of:1.2\"\
           office:version=\"1.2\">\
         <office:body>\
         <office:spreadsheet>\n",
    );

    for sheet in &workbook.sheets {
        xml.push_str(&format!(
            "<table:table table:name=\"{}\">\n",
            xml_escape(&sheet.name)
        ));

        if sheet.cells.is_empty() {
            xml.push_str("<table:table-row><table:table-cell/></table:table-row>\n");
        } else {
            let max_row = sheet.max_row();
            let max_col = sheet.max_col();

            for row in 0..=max_row {
                xml.push_str("<table:table-row>\n");
                for col in 0..=max_col {
                    let cell_xml = value_to_cell_xml(sheet.get_value(row, col));
                    xml.push_str(&cell_xml);
                    xml.push('\n');
                }
                xml.push_str("</table:table-row>\n");
            }
        }

        xml.push_str("</table:table>\n");
    }

    xml.push_str("</office:spreadsheet>\n</office:body>\n</office:document-content>");
    xml
}

fn value_to_cell_xml(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "<table:table-cell/>".to_string(),

        CellValue::Text(s) => format!(
            "<table:table-cell office:value-type=\"string\"><text:p>{}</text:p></table:table-cell>",
            xml_escape(s)
        ),

        CellValue::Number(n) => format!(
            "<table:table-cell office:value-type=\"float\" office:value=\"{n}\"><text:p>{n}</text:p></table:table-cell>",
        ),

        CellValue::Boolean(b) => format!(
            "<table:table-cell office:value-type=\"boolean\" office:boolean-value=\"{b}\"><text:p>{}</text:p></table:table-cell>",
            if *b { "TRUE" } else { "FALSE" }
        ),

        // Formulas: stored with of:= prefix; LibreOffice will recalculate on open.
        // Simple A1-style references are compatible between Excel and ODS formula syntax.
        CellValue::Formula(f) => format!(
            "<table:table-cell table:formula=\"of:={f}\" office:value-type=\"string\"><text:p>={f}</text:p></table:table-cell>",
            f = xml_escape(f)
        ),

        CellValue::Error(e) => format!(
            "<table:table-cell office:value-type=\"string\"><text:p>{}</text:p></table:table-cell>",
            xml_escape(&e.to_string())
        ),
    }
}

// ── Manifest ──────────────────────────────────────────────────────────────────

const MANIFEST_XML: &str = "\
<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\
<manifest:manifest\
  xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\"\
  manifest:version=\"1.2\">\n\
  <manifest:file-entry manifest:full-path=\"/\"\
    manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>\n\
  <manifest:file-entry manifest:full-path=\"content.xml\"\
    manifest:media-type=\"text/xml\"/>\n\
</manifest:manifest>";

// ── calamine Data → CellValue ─────────────────────────────────────────────────

fn data_to_value(dt: &Data) -> CellValue {
    match dt {
        Data::Empty           => CellValue::Empty,
        Data::String(s)       => {
            if s.starts_with('=') {
                CellValue::Formula(s[1..].to_string())
            } else {
                CellValue::Text(s.clone())
            }
        }
        Data::Float(f)        => CellValue::Number(*f),
        Data::Int(i)          => CellValue::Number(*i as f64),
        Data::Bool(b)         => CellValue::Boolean(*b),
        Data::Error(_)        => CellValue::Error(asat_core::CellError::Value),
        Data::DateTime(dt)    => CellValue::Number(dt.as_f64()),
        Data::DateTimeIso(s)  => CellValue::Text(s.clone()),
        Data::DurationIso(s)  => CellValue::Text(s.clone()),
    }
}
