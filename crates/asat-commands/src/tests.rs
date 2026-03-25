use asat_core::{Cell, CellValue, ColMeta, RowMeta, Workbook};

use crate::{Command, DeleteCol, DeleteRow, InsertCol, InsertRow, RemoveSheet, SetCell};

fn wb() -> Workbook {
    Workbook::new()
}

fn num(n: f64) -> Cell {
    Cell::new(CellValue::Number(n))
}

fn txt(s: &str) -> Cell {
    Cell::new(CellValue::Text(s.into()))
}

// ── InsertRow ─────────────────────────────────────────────────────────────────

#[test]
fn insert_row_shifts_cells_down() {
    let mut wb = wb();
    wb.active_mut().set_cell(0, 0, num(1.0)); // A1
    wb.active_mut().set_cell(1, 0, num(2.0)); // A2

    let cmd = InsertRow { sheet: 0, row: 0 };
    cmd.execute(&mut wb).unwrap();

    // Original A1 → A2; original A2 → A3; new A1 is empty
    assert_eq!(wb.active().get_value(0, 0), &CellValue::Empty);
    assert_eq!(wb.active().get_value(1, 0), &CellValue::Number(1.0));
    assert_eq!(wb.active().get_value(2, 0), &CellValue::Number(2.0));
}

#[test]
fn insert_row_shifts_row_meta() {
    let mut wb = wb();
    wb.active_mut().row_meta.insert(0, RowMeta { height: Some(5), ..Default::default() });

    let cmd = InsertRow { sheet: 0, row: 0 };
    cmd.execute(&mut wb).unwrap();

    // Row 0 meta should have shifted to row 1
    assert!(wb.active().row_meta.get(&0).is_none());
    assert_eq!(wb.active().row_meta.get(&1).unwrap().height, Some(5));
}

#[test]
fn insert_row_undo_restores_state() {
    let mut wb = wb();
    wb.active_mut().set_cell(0, 0, num(42.0));
    wb.active_mut().row_meta.insert(0, RowMeta { height: Some(3), ..Default::default() });

    let cmd = InsertRow { sheet: 0, row: 0 };
    cmd.execute(&mut wb).unwrap();
    cmd.undo(&mut wb).unwrap();

    assert_eq!(wb.active().get_value(0, 0), &CellValue::Number(42.0));
    assert_eq!(wb.active().get_value(1, 0), &CellValue::Empty);
    assert_eq!(wb.active().row_meta.get(&0).unwrap().height, Some(3));
}

// ── DeleteRow ─────────────────────────────────────────────────────────────────

#[test]
fn delete_row_shifts_cells_up() {
    let mut wb = wb();
    wb.active_mut().set_cell(0, 0, num(1.0));
    wb.active_mut().set_cell(1, 0, num(2.0));
    wb.active_mut().set_cell(2, 0, num(3.0));

    let cmd = DeleteRow::new(&wb, 0, 0);
    cmd.execute(&mut wb).unwrap();

    assert_eq!(wb.active().get_value(0, 0), &CellValue::Number(2.0));
    assert_eq!(wb.active().get_value(1, 0), &CellValue::Number(3.0));
    assert_eq!(wb.active().get_value(2, 0), &CellValue::Empty);
}

#[test]
fn delete_row_shifts_row_meta() {
    let mut wb = wb();
    wb.active_mut().row_meta.insert(1, RowMeta { height: Some(7), ..Default::default() });

    let cmd = DeleteRow::new(&wb, 0, 0);
    cmd.execute(&mut wb).unwrap();

    assert_eq!(wb.active().row_meta.get(&0).unwrap().height, Some(7));
    assert!(wb.active().row_meta.get(&1).is_none());
}

#[test]
fn delete_row_undo_restores_cells_and_meta() {
    let mut wb = wb();
    wb.active_mut().set_cell(0, 0, num(99.0));
    wb.active_mut().set_cell(1, 0, num(100.0));
    wb.active_mut().row_meta.insert(0, RowMeta { height: Some(4), ..Default::default() });

    let cmd = DeleteRow::new(&wb, 0, 0);
    cmd.execute(&mut wb).unwrap();

    assert_eq!(wb.active().get_value(0, 0), &CellValue::Number(100.0));

    cmd.undo(&mut wb).unwrap();

    assert_eq!(wb.active().get_value(0, 0), &CellValue::Number(99.0));
    assert_eq!(wb.active().get_value(1, 0), &CellValue::Number(100.0));
    assert_eq!(wb.active().row_meta.get(&0).unwrap().height, Some(4));
}

// ── InsertCol / DeleteCol ─────────────────────────────────────────────────────

#[test]
fn insert_col_shifts_cells_right() {
    let mut wb = wb();
    wb.active_mut().set_cell(0, 0, num(1.0)); // A1
    wb.active_mut().set_cell(0, 1, num(2.0)); // B1

    let cmd = InsertCol { sheet: 0, col: 0 };
    cmd.execute(&mut wb).unwrap();

    assert_eq!(wb.active().get_value(0, 0), &CellValue::Empty);
    assert_eq!(wb.active().get_value(0, 1), &CellValue::Number(1.0));
    assert_eq!(wb.active().get_value(0, 2), &CellValue::Number(2.0));
}

#[test]
fn insert_col_shifts_col_meta() {
    let mut wb = wb();
    wb.active_mut().col_meta.insert(0, ColMeta { width: Some(20), ..Default::default() });

    let cmd = InsertCol { sheet: 0, col: 0 };
    cmd.execute(&mut wb).unwrap();

    assert!(wb.active().col_meta.get(&0).is_none());
    assert_eq!(wb.active().col_meta.get(&1).unwrap().width, Some(20));
}

#[test]
fn insert_col_undo_restores_col_meta() {
    let mut wb = wb();
    wb.active_mut().col_meta.insert(0, ColMeta { width: Some(15), ..Default::default() });

    let cmd = InsertCol { sheet: 0, col: 0 };
    cmd.execute(&mut wb).unwrap();
    cmd.undo(&mut wb).unwrap();

    assert_eq!(wb.active().col_meta.get(&0).unwrap().width, Some(15));
    assert!(wb.active().col_meta.get(&1).is_none());
}

#[test]
fn delete_col_undo_restores_col_meta() {
    let mut wb = wb();
    wb.active_mut().set_cell(0, 0, txt("x"));
    wb.active_mut().col_meta.insert(0, ColMeta { width: Some(12), ..Default::default() });
    wb.active_mut().set_cell(0, 1, txt("y"));

    let cmd = DeleteCol::new(&wb, 0, 0);
    cmd.execute(&mut wb).unwrap();

    assert_eq!(wb.active().get_raw_value(0, 0), &CellValue::Text("y".into()));
    assert!(wb.active().col_meta.get(&0).is_none());

    cmd.undo(&mut wb).unwrap();

    assert_eq!(wb.active().get_raw_value(0, 0), &CellValue::Text("x".into()));
    assert_eq!(wb.active().col_meta.get(&0).unwrap().width, Some(12));
}

// ── SetCell merge (undo coalescing) ───────────────────────────────────────────

#[test]
fn setcell_merge_same_cell_coalesces() {
    let mut wb = wb();
    let cmd1 = SetCell::new(&wb, 0, 0, 0, CellValue::Text("a".into()));
    cmd1.execute(&mut wb).unwrap();

    let cmd2 = SetCell::new(&wb, 0, 0, 0, CellValue::Text("ab".into()));

    let merged = cmd1.merge(cmd2.as_any().downcast_ref::<SetCell>().unwrap());
    assert!(merged.is_some(), "consecutive edits to same cell should merge");

    let m = merged.unwrap();
    // The merged cmd's new_value should be cmd2's new_value
    assert_eq!(m.description(), "set cell");
}

#[test]
fn setcell_merge_different_cells_does_not_coalesce() {
    let wb = wb();
    let cmd1 = SetCell::new(&wb, 0, 0, 0, CellValue::Number(1.0));
    let cmd2 = SetCell::new(&wb, 0, 1, 0, CellValue::Number(2.0));

    let merged = cmd1.merge(&cmd2);
    assert!(merged.is_none(), "edits to different cells must not merge");
}

// ── RemoveSheet ───────────────────────────────────────────────────────────────

#[test]
fn remove_sheet_prevents_closing_last_sheet() {
    let mut wb = wb();
    let cmd = RemoveSheet::new(&wb, 0).unwrap();
    let result = cmd.execute(&mut wb);
    assert!(result.is_err(), "cannot close the only sheet");
}

#[test]
fn remove_sheet_execute_and_undo() {
    let mut wb = wb();
    wb.add_sheet("Sheet2");
    assert_eq!(wb.sheets.len(), 2);

    // Write something to Sheet2
    wb.sheets[1].set_cell(0, 0, num(77.0));

    let cmd = RemoveSheet::new(&wb, 1).unwrap();
    cmd.execute(&mut wb).unwrap();
    assert_eq!(wb.sheets.len(), 1, "sheet2 should be removed");

    cmd.undo(&mut wb).unwrap();
    assert_eq!(wb.sheets.len(), 2, "sheet2 should be restored");
    assert_eq!(
        wb.sheets[1].get_value(0, 0),
        &CellValue::Number(77.0),
        "sheet2 data should be restored"
    );
}
