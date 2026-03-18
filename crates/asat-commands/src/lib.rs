use asat_core::{Cell, CellStyle, CellValue, MergeRegion, Workbook};
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
    /// Returns (sheet, row, col) of the primary cell affected by this command.
    fn affected_cell(&self) -> Option<(usize, u32, u32)> {
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
        UndoStack {
            past: Vec::new(),
            future: Vec::new(),
            max_depth: 1000,
        }
    }

    pub fn with_limit(max_depth: usize) -> Self {
        UndoStack {
            past: Vec::new(),
            future: Vec::new(),
            max_depth: max_depth.max(1),
        }
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

    pub fn undo(
        &mut self,
        workbook: &mut Workbook,
    ) -> Result<Option<(usize, u32, u32)>, CommandError> {
        if let Some(cmd) = self.past.pop() {
            cmd.undo(workbook)?;
            let cell = cmd.affected_cell();
            self.future.push(cmd);
            Ok(cell)
        } else {
            Ok(None)
        }
    }

    pub fn redo(
        &mut self,
        workbook: &mut Workbook,
    ) -> Result<Option<(usize, u32, u32)>, CommandError> {
        if let Some(cmd) = self.future.pop() {
            cmd.execute(workbook)?;
            let cell = cmd.affected_cell();
            self.past.push(cmd);
            Ok(cell)
        } else {
            Ok(None)
        }
    }

    pub fn can_undo(&self) -> bool {
        !self.past.is_empty()
    }
    pub fn can_redo(&self) -> bool {
        !self.future.is_empty()
    }
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
        let old_cell = workbook.sheet(sheet).and_then(|s| s.get_cell(row, col));
        let old_value = old_cell
            .map(|c| c.value.clone())
            .unwrap_or(CellValue::Empty);
        let old_style = old_cell.and_then(|c| c.style.clone());
        SetCell {
            sheet,
            row,
            col,
            old_value,
            new_value,
            // Preserve the existing cell style (wrap, bold, colours, etc.) so that
            // typing into a styled empty cell doesn't silently erase the style.
            new_style: old_style.clone(),
            old_style,
        }
    }
}

impl Command for SetCell {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let cell = Cell {
            value: self.new_value.clone(),
            style: self.new_style.clone(),
        };
        sheet.set_cell(self.row, self.col, cell);
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let cell = Cell {
            value: self.old_value.clone(),
            style: self.old_style.clone(),
        };
        sheet.set_cell(self.row, self.col, cell);
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str {
        "set cell"
    }
    fn affected_cell(&self) -> Option<(usize, u32, u32)> {
        Some((self.sheet, self.row, self.col))
    }

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
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Shift all cells at or below `row` down by 1 — descending order to
        // avoid overwriting cells that haven't been moved yet.
        let mut affected: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(r, _)| *r >= self.row)
            .cloned()
            .collect();
        affected.sort_by(|a, b| b.cmp(a));
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r + 1, c), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Remove row `row`, shift everything above it back up
        // First remove the inserted row's cells
        let row_cells: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(r, _)| *r == self.row)
            .cloned()
            .collect();
        for coord in row_cells {
            sheet.cells.remove(&coord);
        }
        // Shift everything above back down — ascending order to avoid overwrites.
        let mut affected: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(r, _)| *r > self.row)
            .cloned()
            .collect();
        affected.sort();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r - 1, c), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str {
        "insert row"
    }
    fn affected_cell(&self) -> Option<(usize, u32, u32)> {
        Some((self.sheet, self.row, 0))
    }
}

#[derive(Debug, Clone)]
pub struct DeleteRow {
    pub sheet: usize,
    pub row: u32,
    pub saved_cells: Vec<(u32, Cell)>, // (col, cell) for undo
}

impl DeleteRow {
    pub fn new(workbook: &Workbook, sheet: usize, row: u32) -> Self {
        let saved_cells = workbook
            .sheet(sheet)
            .map(|s| {
                s.cells
                    .iter()
                    .filter(|((r, _), _)| *r == row)
                    .map(|((_, c), cell)| (*c, cell.clone()))
                    .collect()
            })
            .unwrap_or_default();
        DeleteRow {
            sheet,
            row,
            saved_cells,
        }
    }
}

impl Command for DeleteRow {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Remove this row's cells
        let row_cells: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(r, _)| *r == self.row)
            .cloned()
            .collect();
        for coord in row_cells {
            sheet.cells.remove(&coord);
        }
        // Shift rows above up — must process in ascending row order so that
        // moving row N to N-1 doesn't overwrite row N-1 before it is moved.
        let mut affected: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(r, _)| *r > self.row)
            .cloned()
            .collect();
        affected.sort();
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r - 1, c), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Shift rows back down — must process in descending row order so that
        // moving row N to N+1 doesn't overwrite row N+1 before it is moved.
        let mut affected: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(r, _)| *r >= self.row)
            .cloned()
            .collect();
        affected.sort_by(|a, b| b.cmp(a));
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

    fn description(&self) -> &str {
        "delete row"
    }
    fn affected_cell(&self) -> Option<(usize, u32, u32)> {
        Some((self.sheet, self.row, 0))
    }
}

// ── InsertCol / DeleteCol ────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct InsertCol {
    pub sheet: usize,
    pub col: u32,
}

impl Command for InsertCol {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Shift right — descending col order to avoid overwrites
        let mut affected: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(_, c)| *c >= self.col)
            .cloned()
            .collect();
        affected.sort_by(|a, b| b.1.cmp(&a.1).then(b.0.cmp(&a.0)));
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r, c + 1), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let col_cells: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(_, c)| *c == self.col)
            .cloned()
            .collect();
        for coord in col_cells {
            sheet.cells.remove(&coord);
        }
        // Shift left — ascending col order to avoid overwrites
        let mut affected: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(_, c)| *c > self.col)
            .cloned()
            .collect();
        affected.sort_by(|a, b| a.1.cmp(&b.1).then(a.0.cmp(&b.0)));
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r, c - 1), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str {
        "insert column"
    }
    fn affected_cell(&self) -> Option<(usize, u32, u32)> {
        Some((self.sheet, 0, self.col))
    }
}

#[derive(Debug, Clone)]
pub struct DeleteCol {
    pub sheet: usize,
    pub col: u32,
    pub saved_cells: Vec<(u32, Cell)>, // (row, cell) for undo
}

impl DeleteCol {
    pub fn new(workbook: &Workbook, sheet: usize, col: u32) -> Self {
        let saved_cells = workbook
            .sheet(sheet)
            .map(|s| {
                s.cells
                    .iter()
                    .filter(|((_, c), _)| *c == col)
                    .map(|((r, _), cell)| (*r, cell.clone()))
                    .collect()
            })
            .unwrap_or_default();
        DeleteCol {
            sheet,
            col,
            saved_cells,
        }
    }
}

impl Command for DeleteCol {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        let col_cells: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(_, c)| *c == self.col)
            .cloned()
            .collect();
        for coord in col_cells {
            sheet.cells.remove(&coord);
        }
        // Shift left — ascending col order to avoid overwrites
        let mut affected: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(_, c)| *c > self.col)
            .cloned()
            .collect();
        affected.sort_by(|a, b| a.1.cmp(&b.1).then(a.0.cmp(&b.0)));
        for (r, c) in affected {
            if let Some(cell) = sheet.cells.remove(&(r, c)) {
                sheet.cells.insert((r, c - 1), cell);
            }
        }
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheet_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        // Shift right — descending col order to avoid overwrites
        let mut affected: Vec<_> = sheet
            .cells
            .keys()
            .filter(|(_, c)| *c >= self.col)
            .cloned()
            .collect();
        affected.sort_by(|a, b| b.1.cmp(&a.1).then(b.0.cmp(&a.0)));
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

    fn description(&self) -> &str {
        "delete column"
    }
    fn affected_cell(&self) -> Option<(usize, u32, u32)> {
        Some((self.sheet, 0, self.col))
    }
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

    fn description(&self) -> &str {
        &self.description
    }
    fn affected_cell(&self) -> Option<(usize, u32, u32)> {
        self.commands.first().and_then(|c| c.affected_cell())
    }
}

// ── Clipboard / Registers ─────────────────────────────────────────────────────

#[derive(Debug, Clone, Default)]
pub struct Register {
    pub cells: Vec<Vec<CellValue>>,          // rows of cols
    pub styles: Vec<Vec<Option<CellStyle>>>, // matching styles grid
    pub is_line: bool,                       // true if yanked whole row(s)
    pub source_row: u32,                     // top-left row of yanked region
    pub source_col: u32,                     // top-left col of yanked region
}

#[derive(Debug, Default)]
pub struct RegisterMap {
    pub named: std::collections::HashMap<char, Register>,
    pub unnamed: Register, // "0 register
}

impl RegisterMap {
    pub fn yank(&mut self, name: Option<char>, cells: Vec<Vec<CellValue>>, is_line: bool) {
        let rows = cells.len();
        let cols = cells.first().map(|r| r.len()).unwrap_or(0);
        let styles = vec![vec![None; cols]; rows];
        self.yank_at(name, cells, styles, is_line, 0, 0);
    }

    pub fn yank_at(
        &mut self,
        name: Option<char>,
        cells: Vec<Vec<CellValue>>,
        styles: Vec<Vec<Option<CellStyle>>>,
        is_line: bool,
        source_row: u32,
        source_col: u32,
    ) {
        let reg = Register {
            cells,
            styles,
            is_line,
            source_row,
            source_col,
        };
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

// ── MergeCells ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct MergeCells {
    pub sheet: usize,
    pub row_start: u32,
    pub col_start: u32,
    pub row_end: u32,
    pub col_end: u32,
    /// Overlapping merges that were removed — restored on undo.
    pub removed_merges: Vec<MergeRegion>,
}

impl MergeCells {
    pub fn new(
        workbook: &Workbook,
        sheet: usize,
        row_start: u32,
        col_start: u32,
        row_end: u32,
        col_end: u32,
    ) -> Self {
        let removed_merges = workbook
            .sheets
            .get(sheet)
            .map(|s| {
                s.merges
                    .iter()
                    .filter(|m| {
                        !(m.row_end < row_start
                            || m.row_start > row_end
                            || m.col_end < col_start
                            || m.col_start > col_end)
                    })
                    .cloned()
                    .collect()
            })
            .unwrap_or_default();
        MergeCells {
            sheet,
            row_start,
            col_start,
            row_end,
            col_end,
            removed_merges,
        }
    }
}

impl Command for MergeCells {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheets
            .get_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        sheet.add_merge(self.row_start, self.col_start, self.row_end, self.col_end);
        workbook.dirty = true;
        Ok(())
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheets
            .get_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        sheet.remove_merge(self.row_start, self.col_start);
        for m in &self.removed_merges {
            sheet.merges.push(m.clone());
        }
        workbook.dirty = true;
        Ok(())
    }

    fn description(&self) -> &str {
        "merge cells"
    }
}

// ── UnmergeCells ─────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct UnmergeCells {
    pub sheet: usize,
    pub anchor_row: u32,
    pub anchor_col: u32,
    pub saved: Option<MergeRegion>,
}

impl UnmergeCells {
    pub fn new(workbook: &Workbook, sheet: usize, row: u32, col: u32) -> Self {
        // Find merge whose anchor is exactly (row, col) OR that contains (row, col)
        let saved = workbook
            .sheets
            .get(sheet)
            .and_then(|s| s.merge_at(row, col).cloned());
        UnmergeCells {
            sheet,
            anchor_row: row,
            anchor_col: col,
            saved,
        }
    }
}

impl Command for UnmergeCells {
    fn execute(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        let sheet = workbook
            .sheets
            .get_mut(self.sheet)
            .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
        if let Some(ref m) = self.saved {
            sheet.remove_merge(m.row_start, m.col_start);
            workbook.dirty = true;
            Ok(())
        } else {
            Err(CommandError::Invalid(
                "no merged cell at cursor".to_string(),
            ))
        }
    }

    fn undo(&self, workbook: &mut Workbook) -> Result<(), CommandError> {
        if let Some(ref m) = self.saved {
            let sheet = workbook
                .sheets
                .get_mut(self.sheet)
                .ok_or(CommandError::SheetOutOfRange(self.sheet))?;
            sheet.merges.push(m.clone());
            workbook.dirty = true;
        }
        Ok(())
    }

    fn description(&self) -> &str {
        "unmerge cells"
    }
}
