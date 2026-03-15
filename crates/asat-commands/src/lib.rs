use asat_core::{Cell, CellValue, CellStyle, Workbook};
use thiserror::Error;

// ── Error ────────────────────────────────────────────────────────────────────

#[derive(Debug, Error)]
pub enum CommandError {
    #[error("sheet index {0} out of range")]
    SheetOutOfRange(usize),
    #[error("invalid operation: {0}")]
    Invalid(String),
}

// ── Command Trait ────────────────────────────────────────────────────────────

pub trait Command: Send + Sync {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError>;
    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError>;
    fn description(&self) -> &str;
    /// Try to merge two adjacent commands of the same kind into one.
    /// Returns the merged command if merge succeeded; caller discards `other`.
    fn merge(&self, _other: &dyn Command) -> Option<Box<dyn Command>> {
        None
    }
}

// ── Undo Stack ───────────────────────────────────────────────────────────────

pub struct UndoStack {
    past: Vec<Box<dyn Command>>,
    future: Vec<Box<dyn Command>>,
    max_depth: usize,
}

impl UndoStack {
    pub fn new() -> Self {
        UndoStack { past: Vec::new(), future: Vec::new(), max_depth: 1000 }
    }

    pub fn with_limit(max_depth: usize) -> Self {
        UndoStack { past: Vec::new(), future: Vec::new(), max_depth: max_depth.max(1) }
    }

    pub fn push(&mut self, cmd: Box<dyn Command>) {
        self.future.clear();
        // Try to merge with the most recent command
        if let Some(last) = self.past.last() {
            if let Some(merged) = last.merge(cmd.as_ref()) {
                self.past.pop();
                self.past.push(merged);
                if self.past.len() > self.max_depth {
                    self.past.remove(0);
                }
                return;
            }
        }
        self.past.push(cmd);
        if self.past.len() > self.max_depth {
            self.past.remove(0);
        }
    }

    pub fn undo(&mut self, workbook: &mut Workbook) -> Result<bool, CommandError> {
        if let Some(cmd) = self.past.pop() {
            cmd.undo(workbook)?;
            self.future.push(cmd);
            Ok(true)
        } else {
            Ok(false)
        }
    }

    pub fn redo(&mut self, workbook: &mut Workbook) -> Result<bool, CommandError> {
        if let Some(cmd) = self.future.pop() {
            cmd.execute(workbook)?;
            self.past.push(cmd);
            Ok(true)
        } else {
            Ok(false)
        }
    }

    pub fn can_undo(&self) -> bool { !self.past.is_empty() }
    pub fn can_redo(&self) -> bool { !self.future.is_empty() }
}

impl Default for UndoStack {
    fn default() -> Self {
        UndoStack::new()
    }
}

// ── SetCell Command ───────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct SetCell {
    pub sheet: usize,
    pub row: u32,
    pub col: u32,
    pub old_value: CellValue,
    pub new_value: CellValue,
    pub old_style: Option<CellStyle>,
    pub new_style: Option<CellStyle>,
}

impl SetCell {
    pub fn new(
        workbook: &Workbook,
        sheet: usize,
        row: u32,
        col: u32,
        new_value: CellValue,
    ) -> Self {
        let old_cell = workbook.sheet(sheet)
            .and_then(|s| s.get_cell(row, col));
        let old_value = old_cell.map(|c| c.value.clone()).unwrap_or(CellValue::Empty);
        let old_style = old_cell.and_then(|c| c.style.clone());
        SetCell {
            sheet, row, col,
            old_value, old_style,
            new_value, new_style: None,
        }
    }
}

impl Command for SetCell {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let cell = Cell { value: self.new_value.clone(), style: self.new_style.clone() };
        sheet.set_cell(self.row, self.col, cell);
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let cell = Cell { value: self.old_value.clone(), style: self.old_style.clone() };
        sheet.set_cell(self.row, self.col, cell);
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str { "set cell" }

    fn merge(&self, other: &dyn Command) -> Option<Box<dyn Command>> {
        // Downcast attempt: merge consecutive SetCell on same cell
        // We use a simple trick: check description and coords via Any
        let _ = other; // No std::any::Any bound — skip merge for now
        None
    }
}

// ── InsertRow / DeleteRow ────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct InsertRow {
    pub sheet: usize,
    pub row: u32,
}

impl Command for InsertRow {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Shift all cells at or below `row` down by 1
        let affected: Vec<_> = sheet.cells.keys()
            .filter(|(r, _)| *r >= self.row)
            .cloned()
            .collect();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r + 1, c), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Remove row `row`, shift everything above it back up
        // First remove the inserted row's cells
        let row_cells: Vec<_> = sheet.cells.keys()
            .filter(|(r, _)| *r == self.row)
            .cloned()
            .collect();
        for coord in row_cells {
            sheet.cells.remove(&coord);
        }
        // Shift everything above back down
        let affected: Vec<_> = sheet.cells.keys()
            .filter(|(r, _)| *r > self.row)
            .cloned()
            .collect();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r - 1, c), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str { "insert row" }
}

#[derive(Debug, Clone)]
pub struct DeleteRow {
    pub sheet: usize,
    pub row: u32,
    pub saved_cells: Vec<(u32, Cell)>,  // (col, cell) for undo
}

impl DeleteRow {
    pub fn new(workbook: &Workbook, sheet: usize, row: u32) -> Self {
        let saved_cells = workbook.sheet(sheet)
            .map(|s| {
                s.cells.iter()
                    .filter(|((r, _), _)| *r == row)
                    .map(|((_, c), cell)| (*c, cell.clone()))
                    .collect()
            })
            .unwrap_or_default();
        DeleteRow { sheet, row, saved_cells }
    }
}

impl Command for DeleteRow {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Remove this row's cells
        let row_cells: Vec<_> = sheet.cells.keys()
            .filter(|(r, _)| *r == self.row)
            .cloned()
            .collect();
        for coord in row_cells {
            sheet.cells.remove(&coord);
        }
        // Shift rows above up
        let affected: Vec<_> = sheet.cells.keys()
            .filter(|(r, _)| *r > self.row)
            .cloned()
            .collect();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r - 1, c), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Shift rows back down
        let affected: Vec<_> = sheet.cells.keys()
            .filter(|(r, _)| *r >= self.row)
            .cloned()
            .collect();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r + 1, c), cell);
            }
        }
        // Restore saved cells
        for (col, cell) in &self.saved_cells {
            sheet.cells.insert((self.row, *col), cell.clone());
        }
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str { "delete row" }
}

// ── InsertCol / DeleteCol ────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct InsertCol {
    pub sheet: usize,
    pub col: u32,
}

impl Command for InsertCol {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let affected: Vec<_> = sheet.cells.keys()
            .filter(|(_, c)| *c >= self.col)
            .cloned()
            .collect();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r, c + 1), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let col_cells: Vec<_> = sheet.cells.keys()
            .filter(|(_, c)| *c == self.col)
            .cloned()
            .collect();
        for coord in col_cells {
            sheet.cells.remove(&coord);
        }
        let affected: Vec<_> = sheet.cells.keys()
            .filter(|(_, c)| *c > self.col)
            .cloned()
            .collect();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r, c - 1), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str { "insert column" }
}

#[derive(Debug, Clone)]
pub struct DeleteCol {
    pub sheet: usize,
    pub col: u32,
    pub saved_cells: Vec<(u32, Cell)>,  // (row, cell) for undo
}

impl DeleteCol {
    pub fn new(workbook: &Workbook, sheet: usize, col: u32) -> Self {
        let saved_cells = workbook.sheet(sheet)
            .map(|s| {
                s.cells.iter()
                    .filter(|((_, c), _)| *c == col)
                    .map(|((r, _), cell)| (*r, cell.clone()))
                    .collect()
            })
            .unwrap_or_default();
        DeleteCol { sheet, col, saved_cells }
    }
}

impl Command for DeleteCol {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let col_cells: Vec<_> = sheet.cells.keys()
            .filter(|(_, c)| *c == self.col)
            .cloned()
            .collect();
        for coord in col_cells {
            sheet.cells.remove(&coord);
        }
        let affected: Vec<_> = sheet.cells.keys()
            .filter(|(_, c)| *c > self.col)
            .cloned()
            .collect();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r, c - 1), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook.sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let affected: Vec<_> = sheet.cells.keys()
            .filter(|(_, c)| *c >= self.col)
            .cloned()
            .collect();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r, c + 1), cell);
            }
        }
        for (row, cell) in &self.saved_cells {
            sheet.cells.insert((*row, self.col), cell.clone());
        }
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str { "delete column" }
}

// ── GroupedCommand ────────────────────────────────────────────────────────────

pub struct GroupedCommand {
    pub description: String,
    pub commands: Vec<Box<dyn Command>>,
}

impl Command for GroupedCommand {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        for cmd in &self.commands {
            cmd.execute(workbook)?;
        }
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        for cmd in self.commands.iter().rev() {
            cmd.undo(workbook)?;
        }
        Ok(())
    }

    fn description(&self) -> &str { &self.description }
}

// ── Clipboard / Registers ─────────────────────────────────────────────────────

#[derive(Debug, Clone, Default)]
pub struct Register {
    pub cells: Vec<Vec<CellValue>>,  // rows of cols
    pub is_line: bool,               // true if yanked whole row(s)
}

#[derive(Debug, Default)]
pub struct RegisterMap {
    pub named: std::collections::HashMap<char, Register>,
    pub unnamed: Register,  // "0 register
}

impl RegisterMap {
    pub fn yank(&mut self, name: Option<char>, cells: Vec<Vec<CellValue>>, is_line: bool) {
        let reg = Register { cells, is_line };
        self.unnamed = reg.clone();
        if let Some(ch) = name {
            self.named.insert(ch, reg);
        }
    }

    pub fn get(&self, name: Option<char>) -> &Register {
        if let Some(ch) = name {
            self.named.get(&ch).unwrap_or(&self.unnamed)
        } else {
            &self.unnamed
        }
    }
}
