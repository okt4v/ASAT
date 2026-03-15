use std::collections::{HashMap, HashSet};
use std::path::PathBuf;
use serde::{Deserialize, Serialize};

// ── Cell Value ─────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    Boolean(bool),
    Formula(String),
    Error(CellError),
}

impl CellValue {
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Display string for rendering in the grid
    pub fn display(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
            CellValue::Formula(f) => format!("={}", f),
            CellValue::Error(e) => e.to_string(),
        }
    }

    /// The raw formula string (for formula bar display)
    pub fn formula_bar_display(&self) -> String {
        match self {
            CellValue::Formula(f) => format!("={}", f),
            other => other.display(),
        }
    }
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

// ── Cell Error ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum CellError {
    Div0,
    Name,
    Value,
    Ref,
    Num,
    NA,
    Null,
}

impl std::fmt::Display for CellError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CellError::Div0  => write!(f, "#DIV/0!"),
            CellError::Name  => write!(f, "#NAME?"),
            CellError::Value => write!(f, "#VALUE!"),
            CellError::Ref   => write!(f, "#REF!"),
            CellError::Num   => write!(f, "#NUM!"),
            CellError::NA    => write!(f, "#N/A"),
            CellError::Null  => write!(f, "#NULL!"),
        }
    }
}

// ── Style Types ─────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Color {
    pub const fn rgb(r: u8, g: u8, b: u8) -> Self {
        Color { r, g, b }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize, Default)]
pub enum Alignment {
    #[default]
    Default,
    Left,
    Center,
    Right,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum NumberFormat {
    General,
    Integer,
    Decimal(u8),       // decimal places
    Percentage(u8),
    Currency(String),  // symbol
    Date(String),      // strftime-style pattern
    Custom(String),
    Thousands,             // #,##0
    ThousandsDecimals(u8), // #,##0.00
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub fg: Option<Color>,
    pub bg: Option<Color>,
    pub align: Alignment,
    pub format: Option<NumberFormat>,
}

impl Default for CellStyle {
    fn default() -> Self {
        CellStyle {
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            fg: None,
            bg: None,
            align: Alignment::Default,
            format: None,
        }
    }
}

// ── Cell ────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Cell {
    pub value: CellValue,
    pub style: Option<CellStyle>,
}

impl Cell {
    pub fn new(value: CellValue) -> Self {
        Cell { value, style: None }
    }

    pub fn with_style(value: CellValue, style: CellStyle) -> Self {
        Cell { value, style: Some(style) }
    }
}

impl Default for Cell {
    fn default() -> Self {
        Cell { value: CellValue::Empty, style: None }
    }
}

// ── Row / Col Metadata ──────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RowMeta {
    pub height: Option<u16>,  // None = auto
    pub hidden: bool,
    pub style: Option<CellStyle>,
}

impl Default for RowMeta {
    fn default() -> Self {
        RowMeta { height: None, hidden: false, style: None }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColMeta {
    pub width: Option<u16>,   // None = auto (chars)
    pub hidden: bool,
    pub style: Option<CellStyle>,
}

impl Default for ColMeta {
    fn default() -> Self {
        ColMeta { width: None, hidden: false, style: None }
    }
}

// ── Cell Range ──────────────────────────────────────────────────────────────

/// A rectangular range of cells on a single sheet
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellRange {
    pub sheet: usize,
    pub row_start: u32,
    pub col_start: u32,
    pub row_end: u32,
    pub col_end: u32,
}

impl CellRange {
    pub fn single(sheet: usize, row: u32, col: u32) -> Self {
        CellRange { sheet, row_start: row, col_start: col, row_end: row, col_end: col }
    }

    pub fn new(sheet: usize, row_start: u32, col_start: u32, row_end: u32, col_end: u32) -> Self {
        CellRange {
            sheet,
            row_start: row_start.min(row_end),
            col_start: col_start.min(col_end),
            row_end: row_start.max(row_end),
            col_end: col_start.max(col_end),
        }
    }

    pub fn iter_coords(&self) -> impl Iterator<Item = (u32, u32)> + '_ {
        (self.row_start..=self.row_end)
            .flat_map(move |r| (self.col_start..=self.col_end).map(move |c| (r, c)))
    }
}

// ── Sheet ───────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    pub cells: HashMap<(u32, u32), Cell>,
    pub row_meta: HashMap<u32, RowMeta>,
    pub col_meta: HashMap<u32, ColMeta>,
    pub freeze_rows: u32,
    pub freeze_cols: u32,
    pub computed: HashMap<(u32, u32), CellValue>,
    pub dirty: HashSet<(u32, u32)>,
    pub notes: HashMap<(u32, u32), String>,
}

impl Sheet {
    pub fn new(name: impl Into<String>) -> Self {
        Sheet {
            name: name.into(),
            cells: HashMap::new(),
            row_meta: HashMap::new(),
            col_meta: HashMap::new(),
            freeze_rows: 0,
            freeze_cols: 0,
            computed: HashMap::new(),
            dirty: HashSet::new(),
            notes: HashMap::new(),
        }
    }

    pub fn get_cell(&self, row: u32, col: u32) -> Option<&Cell> {
        self.cells.get(&(row, col))
    }

    pub fn get_cell_mut(&mut self, row: u32, col: u32) -> Option<&mut Cell> {
        self.cells.get_mut(&(row, col))
    }

    pub fn set_cell(&mut self, row: u32, col: u32, cell: Cell) {
        if cell.value.is_empty() && cell.style.is_none() {
            self.cells.remove(&(row, col));
        } else {
            self.cells.insert((row, col), cell);
            self.dirty.insert((row, col));
        }
    }

    pub fn get_value(&self, row: u32, col: u32) -> &CellValue {
        // Return computed value if available, else raw value
        if let Some(computed) = self.computed.get(&(row, col)) {
            return computed;
        }
        self.cells
            .get(&(row, col))
            .map(|c| &c.value)
            .unwrap_or(&CellValue::Empty)
    }

    pub fn display_value(&self, row: u32, col: u32) -> String {
        let val = self.get_value(row, col);
        // Apply NumberFormat from cell style if present
        if let Some(fmt) = self.get_cell(row, col)
            .and_then(|c| c.style.as_ref())
            .and_then(|s| s.format.as_ref())
        {
            return apply_number_format(val, fmt);
        }
        val.display()
    }

    /// Returns the 0-indexed max row that has any data
    pub fn max_row(&self) -> u32 {
        self.cells.keys().map(|(r, _)| *r).max().unwrap_or(0)
    }

    /// Returns the 0-indexed max col that has any data
    pub fn max_col(&self) -> u32 {
        self.cells.keys().map(|(_, c)| *c).max().unwrap_or(0)
    }

    /// Returns the raw stored value (formula string for formula cells), bypassing the computed cache.
    /// Use this for the formula bar display.
    pub fn get_raw_value(&self, row: u32, col: u32) -> &CellValue {
        self.cells
            .get(&(row, col))
            .map(|c| &c.value)
            .unwrap_or(&CellValue::Empty)
    }

    pub fn col_width(&self, col: u32) -> u16 {
        self.col_meta
            .get(&col)
            .and_then(|m| m.width)
            .unwrap_or(10)
    }

    pub fn row_height(&self, row: u32) -> u16 {
        self.row_meta
            .get(&row)
            .and_then(|m| m.height)
            .unwrap_or(1)
    }
}

// ── Workbook ─────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub active_sheet: usize,
    pub file_path: Option<PathBuf>,
    pub dirty: bool,
    pub named_ranges: HashMap<String, CellRange>,
}

impl Workbook {
    pub fn new() -> Self {
        Workbook {
            sheets: vec![Sheet::new("Sheet1")],
            active_sheet: 0,
            file_path: None,
            dirty: false,
            named_ranges: HashMap::new(),
        }
    }

    pub fn active(&self) -> &Sheet {
        &self.sheets[self.active_sheet]
    }

    pub fn active_mut(&mut self) -> &mut Sheet {
        &mut self.sheets[self.active_sheet]
    }

    pub fn sheet(&self, idx: usize) -> Option<&Sheet> {
        self.sheets.get(idx)
    }

    pub fn sheet_mut(&mut self, idx: usize) -> Option<&mut Sheet> {
        self.sheets.get_mut(idx)
    }

    pub fn add_sheet(&mut self, name: impl Into<String>) -> usize {
        let idx = self.sheets.len();
        self.sheets.push(Sheet::new(name));
        idx
    }

    pub fn file_name(&self) -> Option<&str> {
        self.file_path
            .as_deref()
            .and_then(|p| p.file_name())
            .and_then(|n| n.to_str())
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Workbook::new()
    }
}

// ── Cell address helpers ─────────────────────────────────────────────────────

/// Convert 0-indexed column number to Excel-style letter (0 → "A", 25 → "Z", 26 → "AA")
pub fn col_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

/// Convert Excel-style letter to 0-indexed column number ("A" → 0, "Z" → 25, "AA" → 26)
pub fn letter_to_col(s: &str) -> Option<u32> {
    let mut result: u32 = 0;
    for ch in s.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        result = result * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
    }
    if result == 0 { None } else { Some(result - 1) }
}

/// Format a CellValue according to a NumberFormat.
pub fn apply_number_format(val: &CellValue, fmt: &NumberFormat) -> String {
    let n = match val {
        CellValue::Number(n) => *n,
        _ => return val.display(), // non-numeric cells ignore format
    };
    match fmt {
        NumberFormat::General       => val.display(),
        NumberFormat::Integer       => format!("{}", n.round() as i64),
        NumberFormat::Decimal(d)    => format!("{:.prec$}", n, prec = *d as usize),
        NumberFormat::Percentage(d) => format!("{:.prec$}%", n * 100.0, prec = *d as usize),
        NumberFormat::Currency(sym) => {
            if n < 0.0 { format!("-{}{:.2}", sym, n.abs()) }
            else       { format!("{}{:.2}", sym, n) }
        }
        NumberFormat::Date(_pat) => {
            // Treat n as days since 1900-01-01 (Excel epoch)
            let days = n as i64;
            let y = 1900 + days / 365;
            let d_in_year = days % 365;
            format!("{}-{:02}-{:02}", y, d_in_year / 30 + 1, d_in_year % 30 + 1)
        }
        NumberFormat::Custom(_) => val.display(),
        NumberFormat::Thousands => {
            let n_int = n.round() as i64;
            // Format with thousands separator
            let abs = n_int.abs();
            let s = abs.to_string();
            let with_sep: String = s.chars().rev().enumerate()
                .flat_map(|(i, c)| {
                    if i > 0 && i % 3 == 0 { vec![',', c] } else { vec![c] }
                })
                .collect::<String>()
                .chars()
                .rev()
                .collect();
            if n_int < 0 { format!("-{}", with_sep) } else { with_sep }
        }
        NumberFormat::ThousandsDecimals(d) => {
            let factor = 10f64.powi(*d as i32);
            let rounded = (n * factor).round() / factor;
            let int_part = rounded.abs() as i64;
            let s = int_part.to_string();
            let with_sep: String = s.chars().rev().enumerate()
                .flat_map(|(i, c)| {
                    if i > 0 && i % 3 == 0 { vec![',', c] } else { vec![c] }
                })
                .collect::<String>()
                .chars()
                .rev()
                .collect();
            let frac_str = if *d > 0 {
                format!(".{:0>prec$}", ((rounded.abs() - int_part as f64) * factor).round() as u64, prec = *d as usize)
            } else {
                String::new()
            };
            if n < 0.0 { format!("-{}{}", with_sep, frac_str) } else { format!("{}{}", with_sep, frac_str) }
        }
    }
}

/// Format a cell address as "A1" style (0-indexed row, col)
pub fn cell_address(row: u32, col: u32) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

// ── Plugin custom-function registry ──────────────────────────────────────────
//
// Stored here (in the dependency-free core) so both asat-formula and asat-plugins
// can access it without creating a circular dependency.

use std::sync::{Arc, Mutex, OnceLock};

/// A callable registered by the plugin system as a custom formula function.
pub type CustomFn = Arc<dyn Fn(&[CellValue]) -> CellValue + Send + Sync + 'static>;

static CUSTOM_FN_REGISTRY: OnceLock<Mutex<HashMap<String, CustomFn>>> = OnceLock::new();

fn custom_fn_registry() -> &'static Mutex<HashMap<String, CustomFn>> {
    CUSTOM_FN_REGISTRY.get_or_init(|| Mutex::new(HashMap::new()))
}

/// Register a plugin-defined formula function (e.g. `=DOUBLE(A1)`).
/// Names are stored upper-case; existing registrations are replaced.
pub fn register_custom_fn(name: &str, f: CustomFn) {
    if let Ok(mut reg) = custom_fn_registry().lock() {
        reg.insert(name.to_ascii_uppercase(), f);
    }
}

/// Remove a previously registered custom function.
pub fn unregister_custom_fn(name: &str) {
    if let Ok(mut reg) = custom_fn_registry().lock() {
        reg.remove(&name.to_ascii_uppercase());
    }
}

/// Returns `true` if a custom function with this name is registered.
pub fn has_custom_fn(name: &str) -> bool {
    custom_fn_registry()
        .lock()
        .map(|r| r.contains_key(&name.to_ascii_uppercase()))
        .unwrap_or(false)
}

/// Call a registered custom function. Returns `None` if the name is not registered.
/// The registry lock is released *before* calling the function to prevent deadlocks
/// when the function itself tries to register/call other functions.
pub fn call_custom_fn(name: &str, args: &[CellValue]) -> Option<CellValue> {
    let f: CustomFn = {
        let reg = custom_fn_registry().lock().ok()?;
        reg.get(&name.to_ascii_uppercase())?.clone() // clone the Arc; lock drops here
    };
    Some(f(args))
}

/// Returns a snapshot of all registered custom function names (upper-case).
pub fn list_custom_fns() -> Vec<String> {
    custom_fn_registry()
        .lock()
        .map(|r| r.keys().cloned().collect())
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(51), "AZ");
        assert_eq!(col_to_letter(52), "BA");
    }

    #[test]
    fn test_letter_to_col() {
        assert_eq!(letter_to_col("A"), Some(0));
        assert_eq!(letter_to_col("Z"), Some(25));
        assert_eq!(letter_to_col("AA"), Some(26));
        assert_eq!(letter_to_col("AB"), Some(27));
    }

    #[test]
    fn test_cell_address() {
        assert_eq!(cell_address(0, 0), "A1");
        assert_eq!(cell_address(9, 1), "B10");
    }

    #[test]
    fn test_sheet_operations() {
        let mut sheet = Sheet::new("Test");
        sheet.set_cell(0, 0, Cell::new(CellValue::Text("Hello".into())));
        assert_eq!(sheet.display_value(0, 0), "Hello");
        assert_eq!(sheet.display_value(0, 1), "");
        assert_eq!(sheet.max_row(), 0);
        assert_eq!(sheet.max_col(), 0);
    }
}
