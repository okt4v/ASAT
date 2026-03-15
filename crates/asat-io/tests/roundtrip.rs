use asat_core::{Cell, CellValue, Workbook};
use asat_io::{load, save};
use std::path::PathBuf;

fn sample_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.active_mut();
    sheet.set_cell(0, 0, Cell::new(CellValue::Text("Name".into())));
    sheet.set_cell(0, 1, Cell::new(CellValue::Text("Score".into())));
    sheet.set_cell(1, 0, Cell::new(CellValue::Text("Alice".into())));
    sheet.set_cell(1, 1, Cell::new(CellValue::Number(95.5)));
    sheet.set_cell(2, 0, Cell::new(CellValue::Text("Bob".into())));
    sheet.set_cell(2, 1, Cell::new(CellValue::Number(87.0)));
    wb
}

#[test]
fn csv_roundtrip() {
    let tmp = std::env::temp_dir().join("asat_test_roundtrip.csv");
    let wb = sample_workbook();

    // Write
    save(&wb, &tmp).expect("save failed");
    assert!(tmp.exists());

    // Read back
    let wb2 = load(&tmp).expect("load failed");
    let sheet = wb2.active();

    assert_eq!(sheet.display_value(0, 0), "Name");
    assert_eq!(sheet.display_value(0, 1), "Score");
    assert_eq!(sheet.display_value(1, 0), "Alice");
    assert_eq!(sheet.display_value(1, 1), "95.5");
    assert_eq!(sheet.display_value(2, 0), "Bob");
    assert_eq!(sheet.display_value(2, 1), "87");

    std::fs::remove_file(&tmp).ok();
}

#[test]
fn asat_format_roundtrip() {
    let tmp = std::env::temp_dir().join("asat_test_roundtrip.asat");
    let wb = sample_workbook();

    save(&wb, &tmp).expect("save .asat failed");
    let wb2 = load(&tmp).expect("load .asat failed");
    let sheet = wb2.active();

    assert_eq!(sheet.display_value(0, 0), "Name");
    assert_eq!(sheet.display_value(1, 1), "95.5");

    std::fs::remove_file(&tmp).ok();
}
