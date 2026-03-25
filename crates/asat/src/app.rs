use std::collections::HashMap;
use std::path::PathBuf;

use arboard::Clipboard;

use asat_commands::{Command, UndoStack};
use asat_core::{CellStyle, CellValue, Workbook};
use asat_input::InputState;

pub(crate) enum ActionResult {
    Continue,
    Quit,
    ForceQuit,
    /// External process took over the terminal; force a full redraw on return.
    ClearTerminal,
}

// ── Free helper functions (pub(crate) so other modules can use them) ──

pub(crate) fn set_status(status: &mut Option<(String, std::time::Instant)>, msg: String) {
    *status = Some((msg, std::time::Instant::now()));
}

pub(crate) fn swap_path(file_path: &std::path::Path) -> PathBuf {
    let name = file_path
        .file_name()
        .unwrap_or_default()
        .to_string_lossy();
    file_path
        .parent()
        .unwrap_or_else(|| std::path::Path::new("."))
        .join(format!(".{}.swp", name))
}

pub(crate) fn write_swap(workbook: &Workbook) {
    if let Some(ref p) = workbook.file_path {
        let _ = asat_io::save_swap(workbook, &swap_path(p));
    }
}

pub(crate) fn delete_swap(workbook: &Workbook) {
    if let Some(ref p) = workbook.file_path {
        let _ = std::fs::remove_file(swap_path(p));
    }
}

/// Evaluate all formula cells in every sheet and store results in `sheet.computed`.
pub(crate) fn recalculate_all(
    workbook: &mut Workbook,
    ast_cache: &mut HashMap<String, asat_formula::Expr>,
) {
    for idx in 0..workbook.sheets.len() {
        recalculate_sheet(workbook, idx, ast_cache);
    }
}

pub(crate) fn recalculate_sheet(
    workbook: &mut Workbook,
    sheet_idx: usize,
    ast_cache: &mut HashMap<String, asat_formula::Expr>,
) {
    // Collect formula cells (immutable borrow ends at the semicolon)
    let formula_cells: Vec<(u32, u32, String)> = {
        let sheet = &workbook.sheets[sheet_idx];
        sheet
            .cells
            .iter()
            .filter_map(|((r, c), cell)| {
                if let CellValue::Formula(f) = &cell.value {
                    Some((*r, *c, f.clone()))
                } else {
                    None
                }
            })
            .collect()
    };

    if formula_cells.is_empty() {
        workbook.sheets[sheet_idx].computed.clear();
        return;
    }

    // Populate the AST cache for any formula strings we haven't seen yet.
    for (_, _, f) in &formula_cells {
        if !ast_cache.contains_key(f.as_str()) {
            if let Some(expr) = asat_formula::parse_formula(f) {
                ast_cache.insert(f.clone(), expr);
            }
        }
    }

    workbook.sheets[sheet_idx].computed.clear();

    // ── Circular reference detection ──
    let formula_set: std::collections::HashSet<(u32, u32)> =
        formula_cells.iter().map(|(r, c, _)| (*r, *c)).collect();

    let deps: HashMap<(u32, u32), Vec<(u32, u32)>> = formula_cells
        .iter()
        .map(|(r, c, f)| {
            let refs: Vec<(u32, u32)> = if let Some(expr) = ast_cache.get(f.as_str()) {
                asat_formula::collect_same_sheet_refs_expr(expr)
            } else {
                asat_formula::collect_same_sheet_refs(f)
            }
            .into_iter()
            .filter(|rc| formula_set.contains(rc))
            .collect();
            ((*r, *c), refs)
        })
        .collect();

    // DFS cycle detection
    let mut cycle_cells = std::collections::HashSet::new();
    {
        #[derive(PartialEq)]
        enum Color {
            White,
            Gray,
            Black,
        }
        let mut color: HashMap<(u32, u32), Color> = formula_cells
            .iter()
            .map(|(r, c, _)| ((*r, *c), Color::White))
            .collect();

        fn dfs(
            node: (u32, u32),
            deps: &HashMap<(u32, u32), Vec<(u32, u32)>>,
            color: &mut HashMap<(u32, u32), Color>,
            cycle_cells: &mut std::collections::HashSet<(u32, u32)>,
        ) {
            if color.get(&node) == Some(&Color::Black) {
                return;
            }
            if color.get(&node) == Some(&Color::Gray) {
                cycle_cells.insert(node);
                return;
            }
            color.insert(node, Color::Gray);
            if let Some(neighbors) = deps.get(&node) {
                for &neighbor in neighbors {
                    if color.get(&neighbor) == Some(&Color::Gray) {
                        cycle_cells.insert(node);
                        cycle_cells.insert(neighbor);
                    } else {
                        dfs(neighbor, deps, color, cycle_cells);
                        if cycle_cells.contains(&neighbor) {
                            cycle_cells.insert(node);
                        }
                    }
                }
            }
            color.insert(node, Color::Black);
        }

        let keys: Vec<(u32, u32)> = deps.keys().copied().collect();
        for node in keys {
            dfs(node, &deps, &mut color, &mut cycle_cells);
        }
    }

    // Mark cycle cells
    let sheet = &mut workbook.sheets[sheet_idx];
    for &(r, c) in &cycle_cells {
        sheet.computed.insert(
            (r, c),
            asat_core::CellValue::Error(asat_core::CellError::CircularRef),
        );
    }

    // Up to 3 passes so formula->formula chains resolve correctly.
    for _pass in 0..3 {
        let results: Vec<(u32, u32, CellValue)> = formula_cells
            .iter()
            .filter(|(r, c, _)| !cycle_cells.contains(&(*r, *c)))
            .map(|(r, c, f)| {
                let val = if let Some(expr) = ast_cache.get(f.as_str()) {
                    asat_formula::evaluate_expr(expr, workbook, sheet_idx, *r, *c)
                } else {
                    asat_formula::evaluate(f, workbook, sheet_idx, *r, *c)
                };
                (*r, *c, val)
            })
            .collect();
        let sheet = &mut workbook.sheets[sheet_idx];
        for (r, c, val) in results {
            sheet.computed.insert((r, c), val);
        }
    }
}

/// Compute how many spreadsheet rows are fully visible starting from `top_row`.
pub(crate) fn visible_rows_in_height(
    grid_height: u16,
    sheet: &asat_core::Sheet,
    top_row: u32,
) -> u32 {
    let mut remaining = grid_height as u32;
    let mut count = 0u32;
    let mut r = top_row;
    loop {
        let h = sheet.row_height(r) as u32;
        if remaining < h {
            break;
        }
        remaining -= h;
        count += 1;
        r += 1;
        if r > 1_000_000 {
            break;
        }
    }
    count.max(1)
}

/// Compute how many columns are visible starting from `left_col` given `width` pixels.
pub(crate) fn visible_cols_in_width(
    grid_width: u16,
    sheet: &asat_core::Sheet,
    left_col: u32,
) -> u32 {
    const GUTTER: u16 = 5;
    const MIN_COL: u16 = 3;
    let mut avail = grid_width.saturating_sub(GUTTER);
    let mut count = 0u32;
    let mut c = left_col;
    loop {
        let w = sheet.col_width(c).max(MIN_COL);
        if avail < w {
            break;
        }
        avail -= w;
        count += 1;
        c += 1;
        if c > 10_000 {
            break;
        }
    }
    count.max(1)
}

/// Write `text` to the system clipboard using a persistent Clipboard instance.
pub(crate) fn copy_to_clipboard(cb: &mut Option<Clipboard>, text: &str) {
    if let Some(cb) = cb {
        let _ = cb.set_text(text.to_owned());
    }
}

/// Serialise a 2-D grid of CellValues as tab-separated values.
pub(crate) fn cells_to_tsv(cells: &[Vec<CellValue>]) -> String {
    cells
        .iter()
        .map(|row| {
            row.iter()
                .map(|v| v.display())
                .collect::<Vec<_>>()
                .join("\t")
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// Compare two CellValues for sorting purposes.
pub(crate) fn compare_cell_values(a: &CellValue, b: &CellValue) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    match (a, b) {
        (CellValue::Empty, CellValue::Empty) => Ordering::Equal,
        (CellValue::Empty, _) => Ordering::Greater,
        (_, CellValue::Empty) => Ordering::Less,
        (CellValue::Number(x), CellValue::Number(y)) => {
            x.partial_cmp(y).unwrap_or(Ordering::Equal)
        }
        (CellValue::Boolean(x), CellValue::Boolean(y)) => x.cmp(y),
        (CellValue::Text(x), CellValue::Text(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
        (CellValue::Number(_), _) => Ordering::Less,
        (_, CellValue::Number(_)) => Ordering::Greater,
        (CellValue::Text(_), _) => Ordering::Less,
        (_, CellValue::Text(_)) => Ordering::Greater,
        _ => Ordering::Equal,
    }
}

/// Parse a color argument: either `#rrggbb` hex or a named color.
pub(crate) fn parse_color_arg(arg: &str) -> Option<asat_core::Color> {
    let s = arg.trim_start_matches('#');
    if s.len() == 6 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&s[0..2], 16),
            u8::from_str_radix(&s[2..4], 16),
            u8::from_str_radix(&s[4..6], 16),
        ) {
            return Some(asat_core::Color::rgb(r, g, b));
        }
    }
    let (r, g, b): (u8, u8, u8) = match arg.to_ascii_lowercase().as_str() {
        "red" => (220, 50, 50),
        "darkred" | "maroon" => (140, 20, 20),
        "lightred" | "salmon" => (255, 120, 120),
        "orange" => (230, 130, 30),
        "gold" => (220, 180, 30),
        "yellow" => (220, 200, 50),
        "lightyellow" => (255, 255, 150),
        "lime" => (100, 220, 50),
        "green" => (50, 180, 80),
        "darkgreen" | "olive" => (50, 120, 30),
        "lightgreen" => (120, 220, 120),
        "teal" | "cyan" => (40, 180, 180),
        "lightcyan" => (100, 230, 230),
        "blue" => (50, 100, 220),
        "navy" | "darkblue" => (20, 50, 140),
        "lightblue" | "sky" => (100, 160, 255),
        "purple" | "violet" => (140, 80, 220),
        "indigo" => (80, 50, 160),
        "lightpurple" | "lavender" => (180, 140, 255),
        "pink" => (255, 150, 190),
        "hotpink" | "magenta" => (220, 60, 160),
        "brown" | "sienna" => (160, 90, 40),
        "white" => (245, 245, 245),
        "lightgray" | "lightgrey" => (200, 200, 200),
        "gray" | "grey" => (130, 130, 130),
        "darkgray" | "darkgrey" => (70, 70, 70),
        "black" => (15, 15, 15),
        _ => return None,
    };
    Some(asat_core::Color::rgb(r, g, b))
}

/// Return the style range: the visual selection if active, otherwise the cursor cell.
pub(crate) fn style_range(workbook: &Workbook, input: &InputState) -> (u32, u32, u32, u32) {
    use asat_input::Mode;
    if matches!(input.mode, Mode::Visual { .. } | Mode::VisualLine) {
        let (rs, cs, re, ce) = input.visual_selection_bounds();
        let sh = workbook.active();
        (rs, cs, re.min(sh.max_row()), ce.min(sh.max_col()))
    } else if let Some((rs, cs, re, ce)) = input.visual_command_range {
        let sh = workbook.active();
        (rs, cs, re.min(sh.max_row()), ce.min(sh.max_col()))
    } else {
        (
            input.cursor.row,
            input.cursor.col,
            input.cursor.row,
            input.cursor.col,
        )
    }
}

/// Apply a style mutation to the current cell or visual selection (with undo).
pub(crate) fn apply_style_sel(
    workbook: &mut Workbook,
    input: &InputState,
    undo: &mut UndoStack,
    status: &mut Option<(String, std::time::Instant)>,
    f: &dyn Fn(&mut CellStyle),
    msg: &str,
) {
    let sheet_idx = workbook.active_sheet;
    let (row_start, col_start, row_end, col_end) = style_range(workbook, input);

    if row_start == row_end && col_start == col_end {
        apply_style(
            workbook, undo, status, sheet_idx, row_start, col_start, f, msg,
        );
        return;
    }

    let count = (row_end - row_start + 1) as u64 * (col_end - col_start + 1) as u64;
    if count > 50_000 {
        set_status(status, "Selection too large (max 50 000 cells)".to_string());
        return;
    }

    let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
    for r in row_start..=row_end {
        for c in col_start..=col_end {
            let old_cell = workbook.sheet(sheet_idx).and_then(|s| s.get_cell(r, c));
            let old_value = old_cell
                .map(|c| c.value.clone())
                .unwrap_or(CellValue::Empty);
            let old_style = old_cell.and_then(|c| c.style.clone());
            let mut new_style = old_style.clone().unwrap_or_default();
            f(&mut new_style);
            cmds.push(Box::new(asat_commands::SetCell {
                sheet: sheet_idx,
                row: r,
                col: c,
                old_value: old_value.clone(),
                new_value: old_value,
                old_style,
                new_style: Some(new_style),
            }));
        }
    }
    let n_rows = row_end - row_start + 1;
    let n_cols = col_end - col_start + 1;
    let grouped = Box::new(asat_commands::GroupedCommand {
        description: msg.to_string(),
        commands: cmds,
    });
    match grouped.execute(workbook) {
        Ok(_) => {
            undo.push(grouped);
            set_status(status, format!("{} ({}x{})", msg, n_rows, n_cols));
        }
        Err(e) => set_status(status, format!("Error: {}", e)),
    }
}

/// Fetch the current CellStyle for a cell (or default), apply `f`, then push as a SetCell command.
#[allow(clippy::too_many_arguments)]
pub(crate) fn apply_style(
    workbook: &mut Workbook,
    undo: &mut UndoStack,
    status: &mut Option<(String, std::time::Instant)>,
    sheet_idx: usize,
    row: u32,
    col: u32,
    f: impl FnOnce(&mut CellStyle),
    msg: &str,
) {
    let old_cell = workbook.sheet(sheet_idx).and_then(|s| s.get_cell(row, col));
    let old_value = old_cell
        .map(|c| c.value.clone())
        .unwrap_or(CellValue::Empty);
    let old_style = old_cell.and_then(|c| c.style.clone());
    let mut new_style = old_style.clone().unwrap_or_default();
    f(&mut new_style);
    let cmd = Box::new(asat_commands::SetCell {
        sheet: sheet_idx,
        row,
        col,
        old_value: old_value.clone(),
        new_value: old_value,
        old_style,
        new_style: Some(new_style),
    });
    match cmd.execute(workbook) {
        Ok(_) => {
            undo.push(cmd);
            set_status(status, msg.to_string());
        }
        Err(e) => set_status(status, format!("Error: {}", e)),
    }
}

// ── Recent files I/O ──

pub(crate) fn recent_files_path() -> PathBuf {
    dirs::data_local_dir()
        .unwrap_or_else(|| {
            PathBuf::from(std::env::var("HOME").unwrap_or_default()).join(".local/share")
        })
        .join("asat")
        .join("recent")
}

pub(crate) fn load_recent_files(limit: usize) -> Vec<String> {
    let path = recent_files_path();
    std::fs::read_to_string(path)
        .unwrap_or_default()
        .lines()
        .filter(|l| !l.trim().is_empty())
        .map(|l| {
            let p = PathBuf::from(l);
            if p.is_absolute() {
                l.to_string()
            } else {
                std::fs::canonicalize(&p)
                    .map(|c| c.to_string_lossy().into_owned())
                    .unwrap_or_else(|_| {
                        std::env::current_dir()
                            .map(|cwd| cwd.join(l).to_string_lossy().into_owned())
                            .unwrap_or_else(|_| l.to_string())
                    })
            }
        })
        .take(limit.max(1))
        .collect()
}

pub(crate) fn save_recent_files(files: &[String]) {
    let path = recent_files_path();
    if let Some(dir) = path.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    let content = files.join("\n");
    let _ = std::fs::write(path, content);
}

/// Prepend `path` to the recent list (dedup + cap at 20).
pub(crate) fn push_recent(path: &str, recents: &mut Vec<String>) {
    let abs = std::fs::canonicalize(path)
        .map(|p| p.to_string_lossy().into_owned())
        .unwrap_or_else(|_| {
            let p = PathBuf::from(path);
            if p.is_absolute() {
                path.to_string()
            } else {
                std::env::current_dir()
                    .map(|cwd| cwd.join(path).to_string_lossy().into_owned())
                    .unwrap_or_else(|_| path.to_string())
            }
        });
    recents.retain(|r| r != &abs);
    recents.insert(0, abs);
    recents.truncate(20);
}

// ── File scanning (for fuzzy finder) ──

pub(crate) fn scan_files(root: &PathBuf) -> Vec<String> {
    let mut results = Vec::new();
    scan_dir(root, root, 5, &mut results);
    results.sort();
    results.truncate(2000);
    results
}

pub(crate) fn scan_dir(
    root: &PathBuf,
    dir: &PathBuf,
    depth: usize,
    results: &mut Vec<String>,
) {
    if depth == 0 {
        return;
    }
    let entries = match std::fs::read_dir(dir) {
        Ok(e) => e,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("");
        if name.starts_with('.') {
            continue;
        }
        if matches!(
            name,
            "target" | "node_modules" | "__pycache__" | ".git" | "dist" | "build"
        ) {
            continue;
        }
        if path.is_dir() {
            scan_dir(root, &path, depth - 1, results);
        } else if let Ok(rel) = path.strip_prefix(root) {
            results.push(rel.to_string_lossy().to_string());
        }
    }
}

// ── Config editor ──

pub(crate) fn open_config_in_editor() {
    use crossterm::execute;
    use crossterm::terminal::{disable_raw_mode, enable_raw_mode};

    let config_path = {
        let base = std::env::var("XDG_CONFIG_HOME")
            .map(PathBuf::from)
            .unwrap_or_else(|_| {
                dirs::home_dir()
                    .unwrap_or_else(|| PathBuf::from("."))
                    .join(".config")
            });
        base.join("asat").join("config.toml")
    };

    if !config_path.exists() {
        let _ = asat_config::Config::write_default();
    }

    let editor = std::env::var("EDITOR")
        .or_else(|_| std::env::var("VISUAL"))
        .unwrap_or_else(|_| "nano".to_string());

    let _ = disable_raw_mode();
    let _ = execute!(std::io::stdout(), crossterm::terminal::LeaveAlternateScreen);

    let _ = std::process::Command::new(&editor)
        .arg(&config_path)
        .status();

    let _ = enable_raw_mode();
    let _ = execute!(std::io::stdout(), crossterm::terminal::EnterAlternateScreen);
}

// ── Go-to definition helpers ──

/// Parse a formula string and return the first cell reference.
pub(crate) fn find_first_cell_ref(formula: &str) -> Option<(Option<String>, u32, u32)> {
    let tokens = asat_formula::lexer::lex(formula).ok()?;
    let expr = asat_formula::parser::parse(&tokens).ok()?;
    find_first_ref_in_expr(&expr)
}

pub(crate) fn find_first_ref_in_expr(
    expr: &asat_formula::parser::Expr,
) -> Option<(Option<String>, u32, u32)> {
    use asat_formula::parser::Expr;
    match expr {
        Expr::CellRef {
            sheet, row, col, ..
        } => Some((sheet.clone(), *row, *col)),
        Expr::RangeRef {
            sheet, row1, col1, ..
        } => Some((sheet.clone(), *row1, *col1)),
        Expr::UnaryMinus(e) | Expr::UnaryPlus(e) => find_first_ref_in_expr(e),
        Expr::Add(a, b)
        | Expr::Sub(a, b)
        | Expr::Mul(a, b)
        | Expr::Div(a, b)
        | Expr::Pow(a, b)
        | Expr::Concat(a, b)
        | Expr::Eq(a, b)
        | Expr::Neq(a, b)
        | Expr::Lt(a, b)
        | Expr::Lte(a, b)
        | Expr::Gt(a, b)
        | Expr::Gte(a, b) => find_first_ref_in_expr(a).or_else(|| find_first_ref_in_expr(b)),
        Expr::Call { args, .. } => args.iter().find_map(find_first_ref_in_expr),
        _ => None,
    }
}

// ── Cell address parsing ──

/// Parse an Excel-style cell address like "B15" or "AA3" into (row, col) 0-indexed.
pub(crate) fn parse_cell_address(addr: &str) -> Option<(u32, u32)> {
    let addr = addr.trim().to_uppercase();
    let col_end = addr.find(|c: char| c.is_ascii_digit())?;
    if col_end == 0 {
        return None;
    }
    let col_str = &addr[..col_end];
    let row_str = &addr[col_end..];
    let col = asat_core::letter_to_col(col_str)?;
    let row: u32 = row_str.parse::<u32>().ok()?.checked_sub(1)?;
    Some((row, col))
}

/// Parse a range address like "A1:C10" into a CellRange.
pub(crate) fn parse_range_address(s: &str, sheet: usize) -> Option<asat_core::CellRange> {
    let s = s.trim().to_uppercase();
    if let Some(colon) = s.find(':') {
        let (a, b) = (&s[..colon], &s[colon + 1..]);
        let (r1, c1) = parse_cell_address(a)?;
        let (r2, c2) = parse_cell_address(b)?;
        Some(asat_core::CellRange::new(sheet, r1, c1, r2, c2))
    } else {
        let (r, c) = parse_cell_address(&s)?;
        Some(asat_core::CellRange::single(sheet, r, c))
    }
}

/// Check if a cell value matches a filter condition.
pub(crate) fn filter_row_matches(
    val: &CellValue,
    op: &str,
    val_str: &str,
    val_num: Option<f64>,
) -> bool {
    if let (Some(vn), Some(cn)) = (
        val_num,
        match val {
            CellValue::Number(n) => Some(*n),
            _ => None,
        },
    ) {
        return match op {
            ">" => cn > vn,
            "<" => cn < vn,
            ">=" => cn >= vn,
            "<=" => cn <= vn,
            "=" | "==" => (cn - vn).abs() < 1e-10,
            "<>" | "!=" => (cn - vn).abs() >= 1e-10,
            _ => false,
        };
    }
    let cell_text = val.display().to_lowercase();
    let needle = val_str.to_lowercase();
    match op {
        "=" | "==" => cell_text == needle,
        "<>" | "!=" => cell_text != needle,
        _ => cell_text.contains(&needle),
    }
}

/// Parse a range for CF rules, also accepting whole-column ranges like "A:A" or "A:C".
pub(crate) fn parse_range_address_cf(
    s: &str,
    sheet: usize,
    workbook: &Workbook,
) -> Option<asat_core::CellRange> {
    let upper = s.trim().to_uppercase();
    if let Some(colon) = upper.find(':') {
        let left = &upper[..colon];
        let right = &upper[colon + 1..];
        if left.chars().all(|c| c.is_ascii_alphabetic())
            && right.chars().all(|c| c.is_ascii_alphabetic())
        {
            let c1 = asat_core::letter_to_col(left)?;
            let c2 = asat_core::letter_to_col(right)?;
            let max_row = workbook.active().max_row().max(9999);
            return Some(asat_core::CellRange::new(
                sheet,
                0,
                c1.min(c2),
                max_row,
                c1.max(c2),
            ));
        }
    }
    parse_range_address(s, sheet)
}

// ── Date/text cycling helpers ──

pub(crate) fn cycle_text_sequence(text: &str, delta: i32) -> Option<String> {
    const MONTHS_LONG: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    const MONTHS_SHORT: [&str; 12] = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];
    const DAYS_LONG: [&str; 7] = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ];
    const DAYS_SHORT: [&str; 7] = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

    let apply_case = |template: &str| -> String {
        let all_upper = text.chars().all(|c| !c.is_alphabetic() || c.is_uppercase());
        let all_lower = text.chars().all(|c| !c.is_alphabetic() || c.is_lowercase());
        if all_upper {
            template.to_uppercase()
        } else if all_lower {
            template.to_lowercase()
        } else {
            template.to_string()
        }
    };

    let lower = text.to_lowercase();

    for lists in [
        (&MONTHS_LONG[..], 12usize),
        (&MONTHS_SHORT[..], 12),
        (&DAYS_LONG[..], 7),
        (&DAYS_SHORT[..], 7),
    ] {
        let (names, len) = lists;
        if let Some(idx) = names.iter().position(|&n| n.to_lowercase() == lower) {
            let next = ((idx as i32 + delta).rem_euclid(len as i32)) as usize;
            return Some(apply_case(names[next]));
        }
    }
    None
}

pub(crate) fn days_in_month(month: i32, year: i32) -> i32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}

pub(crate) fn add_days_to_date(
    mut day: i32,
    mut month: i32,
    mut year: i32,
    delta: i32,
) -> (i32, i32, i32) {
    day += delta;
    loop {
        if day < 1 {
            month -= 1;
            if month < 1 {
                month = 12;
                year -= 1;
            }
            day += days_in_month(month, year);
        } else if day > days_in_month(month, year) {
            day -= days_in_month(month, year);
            month += 1;
            if month > 12 {
                month = 1;
                year += 1;
            }
        } else {
            break;
        }
    }
    (day, month, year)
}

pub(crate) fn cycle_date(text: &str, delta: i32) -> Option<String> {
    let sep = if text.contains('.') && !text.contains('/') && !text.contains('-') {
        '.'
    } else if text.contains('/') && !text.contains('.') && !text.contains('-') {
        '/'
    } else if text.contains('-') && !text.contains('.') && !text.contains('/') {
        '-'
    } else {
        return None;
    };

    let parts: Vec<&str> = text.splitn(4, sep).collect();
    if parts.len() != 3 {
        return None;
    }
    if parts
        .iter()
        .any(|p| !p.chars().all(|c| c.is_ascii_digit()))
    {
        return None;
    }

    let n: Vec<i32> = parts.iter().filter_map(|p| p.parse().ok()).collect();
    if n.len() != 3 {
        return None;
    }

    #[derive(Clone, Copy)]
    #[allow(clippy::upper_case_acronyms)]
    enum Layout {
        DMY,
        MDY,
        YMD,
    }

    let (layout, two_digit_year) = if parts[0].len() == 4 {
        (Layout::YMD, false)
    } else if parts[2].len() == 4 {
        if n[0] <= 12 && n[1] > 12 {
            (Layout::MDY, false)
        } else {
            (Layout::DMY, false)
        }
    } else if parts[2].len() == 2 {
        (Layout::DMY, true)
    } else {
        return None;
    };

    let (day, month, full_year) = match layout {
        Layout::DMY => {
            let y = if two_digit_year {
                if n[2] < 70 {
                    2000 + n[2]
                } else {
                    1900 + n[2]
                }
            } else {
                n[2]
            };
            (n[0], n[1], y)
        }
        Layout::MDY => (n[1], n[0], n[2]),
        Layout::YMD => (n[2], n[1], n[0]),
    };

    if !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }

    let (nd, nm, ny) = add_days_to_date(day, month, full_year, delta);

    let fmt_field =
        |val: i32, orig: &str| -> String { format!("{:0>width$}", val, width = orig.len()) };

    Some(match layout {
        Layout::DMY => format!(
            "{}{}{}{}{}",
            fmt_field(nd, parts[0]),
            sep,
            fmt_field(nm, parts[1]),
            sep,
            fmt_field(if two_digit_year { ny % 100 } else { ny }, parts[2])
        ),
        Layout::MDY => format!(
            "{}{}{}{}{}",
            fmt_field(nm, parts[0]),
            sep,
            fmt_field(nd, parts[1]),
            sep,
            fmt_field(ny, parts[2])
        ),
        Layout::YMD => format!(
            "{}{}{}{}{}",
            fmt_field(ny, parts[0]),
            sep,
            fmt_field(nm, parts[1]),
            sep,
            fmt_field(nd, parts[2])
        ),
    })
}

// ── Auto-fill series detection ──

pub(crate) const WEEKDAYS_SHORT: &[&str] = &["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
pub(crate) const WEEKDAYS_FULL: &[&str] = &[
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
];
pub(crate) const MONTHS_SHORT_LIST: &[&str] = &[
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];
pub(crate) const MONTHS_FULL: &[&str] = &[
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

#[derive(Clone, Copy)]
pub(crate) enum CaseStyle {
    Lower,
    Upper,
    Title,
}

pub(crate) fn detect_case(s: &str) -> CaseStyle {
    if s.chars().all(|c| c.is_ascii_uppercase()) {
        CaseStyle::Upper
    } else if s
        .chars()
        .next()
        .is_some_and(|c| c.is_ascii_uppercase())
        && s.chars().skip(1).all(|c| c.is_ascii_lowercase())
    {
        CaseStyle::Title
    } else {
        CaseStyle::Lower
    }
}

pub(crate) fn apply_case_style(s: &str, style: CaseStyle) -> String {
    match style {
        CaseStyle::Lower => s.to_ascii_lowercase(),
        CaseStyle::Upper => s.to_ascii_uppercase(),
        CaseStyle::Title => {
            let mut c = s.chars();
            match c.next() {
                None => String::new(),
                Some(first) => {
                    first.to_uppercase().to_string() + &c.as_str().to_ascii_lowercase()
                }
            }
        }
    }
}

pub(crate) fn match_name_list(s: &str, short: &[&str], full: &[&str]) -> Option<(usize, bool)> {
    let lower = s.to_ascii_lowercase();
    if let Some(idx) = full
        .iter()
        .position(|n| n.to_ascii_lowercase() == lower)
    {
        return Some((idx, true));
    }
    if let Some(idx) = short
        .iter()
        .position(|n| n.to_ascii_lowercase() == lower)
    {
        return Some((idx, false));
    }
    None
}

/// Given a seed slice of CellValues, return a closure that produces
/// the i-th fill value (0 = first cell *after* the seed).
pub(crate) fn auto_fill_series(seed: &[CellValue]) -> Box<dyn Fn(usize) -> CellValue> {
    if seed.is_empty() {
        return Box::new(|_| CellValue::Empty);
    }

    // Try numeric arithmetic
    let nums: Option<Vec<f64>> = seed
        .iter()
        .map(|v| {
            if let CellValue::Number(n) = v {
                Some(*n)
            } else {
                None
            }
        })
        .collect();
    if let Some(ns) = nums {
        let start = ns[0];
        let step = if ns.len() >= 2 { ns[1] - ns[0] } else { 1.0 };
        let seed_len = ns.len() as f64;
        return Box::new(move |i| CellValue::Number(start + (seed_len + i as f64) * step));
    }

    // Try weekday cycle
    let wdays: Option<Vec<(usize, bool)>> = seed
        .iter()
        .map(|v| {
            if let CellValue::Text(s) = v {
                match_name_list(s, WEEKDAYS_SHORT, WEEKDAYS_FULL)
            } else {
                None
            }
        })
        .collect();
    if let Some(wd) = wdays {
        let seed_len = wd.len();
        let first = wd[0].0;
        let is_full = wd[0].1;
        let case_style = if let CellValue::Text(s) = &seed[0] {
            detect_case(s)
        } else {
            CaseStyle::Title
        };
        let list: &'static [&'static str] = if is_full {
            WEEKDAYS_FULL
        } else {
            WEEKDAYS_SHORT
        };
        let cycle_len = list.len();
        return Box::new(move |i| {
            let idx = (first + seed_len + i) % cycle_len;
            CellValue::Text(apply_case_style(list[idx], case_style))
        });
    }

    // Try month cycle
    let mths: Option<Vec<(usize, bool)>> = seed
        .iter()
        .map(|v| {
            if let CellValue::Text(s) = v {
                match_name_list(s, MONTHS_SHORT_LIST, MONTHS_FULL)
            } else {
                None
            }
        })
        .collect();
    if let Some(ms) = mths {
        let seed_len = ms.len();
        let first = ms[0].0;
        let is_full = ms[0].1;
        let case_style = if let CellValue::Text(s) = &seed[0] {
            detect_case(s)
        } else {
            CaseStyle::Title
        };
        let list: &'static [&'static str] = if is_full {
            MONTHS_FULL
        } else {
            MONTHS_SHORT_LIST
        };
        let cycle_len = list.len();
        return Box::new(move |i| {
            let idx = (first + seed_len + i) % cycle_len;
            CellValue::Text(apply_case_style(list[idx], case_style))
        });
    }

    // Fallback: cycle through seed values
    let owned: Vec<CellValue> = seed.to_vec();
    let seed_len = owned.len();
    Box::new(move |i| owned[(seed_len + i) % seed_len].clone())
}

/// Keep `input.subcmd_completions` in sync with the current command buffer.
pub(crate) fn update_subcmd_completions(input: &mut InputState) {
    use asat_input::Mode;
    if !matches!(input.mode, Mode::Command) {
        return;
    }
    let buf = input.command_buffer.clone();
    if let Some(space_idx) = buf.find(' ') {
        let verb = buf[..space_idx].to_ascii_lowercase();
        let arg_prefix = buf[space_idx + 1..].to_ascii_lowercase();

        match verb.as_str() {
            "theme" => {
                let themes = asat_config::builtin_themes();
                input.subcmd_completions = themes
                    .iter()
                    .filter(|t| {
                        let n = t.name.to_ascii_lowercase();
                        let id = t.id.to_ascii_lowercase();
                        arg_prefix.is_empty()
                            || n.starts_with(&arg_prefix)
                            || id.starts_with(&arg_prefix)
                            || n.contains(&arg_prefix)
                    })
                    .map(|t| t.name.to_string())
                    .collect();
            }
            _ => input.subcmd_completions.clear(),
        }
    } else {
        input.subcmd_completions.clear();
        input.subcmd_completion_idx = None;
    }
}
