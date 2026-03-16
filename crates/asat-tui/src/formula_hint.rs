use crate::{darken, parse_hex_color, RenderState};
use asat_input::{Mode, FN_NAMES};
use ratatui::{
    layout::Rect,
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Clear, Paragraph},
    Frame,
};

const ROW_GUTTER_WIDTH: u16 = 5;
const MIN_COL_WIDTH: u16 = 3;
const MAX_NAME_HINTS: usize = 6;

// ── Function signature database ───────────────────────────────────────────────

struct FnDef {
    name: &'static str,
    /// Individual argument tokens (shown one per slot, current arg is highlighted)
    args: &'static [&'static str],
    desc: &'static str,
}

macro_rules! fn_def {
    ($name:expr, [$($arg:expr),*], $desc:expr) => {
        FnDef { name: $name, args: &[$($arg),*], desc: $desc }
    };
}

static FN_DEFS: &[FnDef] = &[
    // Math
    fn_def!(
        "SUM",
        ["number1", "[number2, ...]"],
        "Sum all numbers in range"
    ),
    fn_def!(
        "AVERAGE",
        ["number1", "[number2, ...]"],
        "Average of all numbers"
    ),
    fn_def!(
        "AVG",
        ["number1", "[number2, ...]"],
        "Average of all numbers"
    ),
    fn_def!("MIN", ["number1", "[number2, ...]"], "Smallest value"),
    fn_def!("MAX", ["number1", "[number2, ...]"], "Largest value"),
    fn_def!("COUNT", ["value1", "[value2, ...]"], "Count numeric cells"),
    fn_def!(
        "COUNTA",
        ["value1", "[value2, ...]"],
        "Count non-empty cells"
    ),
    fn_def!("ABS", ["number"], "Absolute value"),
    fn_def!("ROUND", ["number", "digits"], "Round to N decimal places"),
    fn_def!("ROUNDUP", ["number", "digits"], "Round up (away from zero)"),
    fn_def!(
        "ROUNDDOWN",
        ["number", "digits"],
        "Round down (toward zero)"
    ),
    fn_def!(
        "FLOOR",
        ["number", "significance"],
        "Round down to multiple"
    ),
    fn_def!(
        "CEILING",
        ["number", "significance"],
        "Round up to multiple"
    ),
    fn_def!("MOD", ["number", "divisor"], "Remainder after division"),
    fn_def!("POWER", ["number", "power"], "Number raised to a power"),
    fn_def!("SQRT", ["number"], "Square root"),
    fn_def!("LN", ["number"], "Natural logarithm"),
    fn_def!("LOG", ["number", "[base]"], "Logarithm with optional base"),
    fn_def!("LOG10", ["number"], "Base-10 logarithm"),
    fn_def!("EXP", ["number"], "e raised to a power"),
    fn_def!("INT", ["number"], "Round down to integer"),
    fn_def!(
        "TRUNC",
        ["number", "[digits]"],
        "Truncate to N decimal places"
    ),
    fn_def!("SIGN", ["number"], "Sign: -1, 0, or 1"),
    // Text
    fn_def!("LEN", ["text"], "Number of characters"),
    fn_def!("LEFT", ["text", "[num_chars]"], "First N characters"),
    fn_def!("RIGHT", ["text", "[num_chars]"], "Last N characters"),
    fn_def!(
        "MID",
        ["text", "start", "num_chars"],
        "Substring from position"
    ),
    fn_def!("TRIM", ["text"], "Remove extra whitespace"),
    fn_def!("UPPER", ["text"], "Convert to uppercase"),
    fn_def!("LOWER", ["text"], "Convert to lowercase"),
    fn_def!("PROPER", ["text"], "Capitalize each word"),
    fn_def!(
        "CONCATENATE",
        ["text1", "text2", "[text3, ...]"],
        "Join text strings"
    ),
    fn_def!(
        "CONCAT",
        ["text1", "text2", "[text3, ...]"],
        "Join text strings"
    ),
    fn_def!("TEXT", ["value", "format"], "Format value as text"),
    fn_def!("VALUE", ["text"], "Parse text as number"),
    fn_def!(
        "FIND",
        ["find_text", "within_text", "[start]"],
        "Case-sensitive position search"
    ),
    fn_def!(
        "SEARCH",
        ["find_text", "within_text", "[start]"],
        "Case-insensitive position search"
    ),
    fn_def!(
        "SUBSTITUTE",
        ["text", "old_text", "new_text", "[nth]"],
        "Replace occurrences in text"
    ),
    fn_def!(
        "REPLACE",
        ["text", "start", "num_chars", "new_text"],
        "Replace by position"
    ),
    fn_def!("REPT", ["text", "times"], "Repeat text N times"),
    // Logic
    fn_def!(
        "IF",
        ["condition", "value_if_true", "[value_if_false]"],
        "Conditional value"
    ),
    fn_def!(
        "AND",
        ["condition1", "[condition2, ...]"],
        "True if all conditions true"
    ),
    fn_def!(
        "OR",
        ["condition1", "[condition2, ...]"],
        "True if any condition true"
    ),
    fn_def!("NOT", ["condition"], "Invert a boolean"),
    fn_def!("ISNUMBER", ["value"], "True if value is a number"),
    fn_def!("ISTEXT", ["value"], "True if value is text"),
    fn_def!("ISBLANK", ["value"], "True if cell is empty"),
    fn_def!("ISERROR", ["value"], "True if value is an error"),
    fn_def!(
        "IFERROR",
        ["value", "value_if_error"],
        "Return fallback on error"
    ),
    fn_def!("ISLOGICAL", ["value"], "True if value is boolean"),
    // Lookup
    fn_def!(
        "VLOOKUP",
        ["lookup_value", "table", "col_index", "[range_lookup]"],
        "Vertical lookup"
    ),
    fn_def!(
        "HLOOKUP",
        ["lookup_value", "table", "row_index", "[range_lookup]"],
        "Horizontal lookup"
    ),
    fn_def!(
        "XLOOKUP",
        [
            "lookup_value",
            "lookup_array",
            "return_array",
            "[if_not_found]",
            "[match_mode]"
        ],
        "Extended lookup"
    ),
    fn_def!(
        "INDEX",
        ["array", "row", "[col]"],
        "Value at row/col in array"
    ),
    fn_def!(
        "MATCH",
        ["lookup_value", "array", "[match_type]"],
        "Position of value in array"
    ),
    fn_def!(
        "OFFSET",
        ["reference", "rows", "cols", "[height]", "[width]"],
        "Shifted reference"
    ),
    fn_def!(
        "CHOOSE",
        ["index", "value1", "[value2, ...]"],
        "Pick value by index"
    ),
    // Date
    fn_def!("NOW", [], "Current date and time"),
    fn_def!("TODAY", [], "Today's date"),
    fn_def!("DATE", ["year", "month", "day"], "Build a date value"),
    fn_def!("YEAR", ["date"], "Year from date"),
    fn_def!("MONTH", ["date"], "Month (1-12) from date"),
    fn_def!("DAY", ["date"], "Day (1-31) from date"),
    // Statistical
    fn_def!(
        "SUMIF",
        ["range", "criteria", "[sum_range]"],
        "Sum cells matching criteria"
    ),
    fn_def!(
        "COUNTIF",
        ["range", "criteria"],
        "Count cells matching criteria"
    ),
    fn_def!(
        "SUMPRODUCT",
        ["array1", "[array2, ...]"],
        "Sum of products of arrays"
    ),
    fn_def!("LARGE", ["array", "k"], "K-th largest value"),
    fn_def!("SMALL", ["array", "k"], "K-th smallest value"),
    fn_def!("MEDIAN", ["number1", "[number2, ...]"], "Middle value"),
    fn_def!(
        "STDEV",
        ["number1", "[number2, ...]"],
        "Standard deviation (sample)"
    ),
    fn_def!("VAR", ["number1", "[number2, ...]"], "Variance (sample)"),
    fn_def!(
        "AVERAGEIF",
        ["range", "criteria", "[avg_range]"],
        "Average of matching cells"
    ),
    fn_def!(
        "MAXIFS",
        ["max_range", "criteria_range", "criteria"],
        "Max of matching cells"
    ),
    fn_def!(
        "MINIFS",
        ["min_range", "criteria_range", "criteria"],
        "Min of matching cells"
    ),
    fn_def!(
        "RANK",
        ["number", "ref", "[order]"],
        "Rank of value in list"
    ),
    fn_def!("PERCENTILE", ["array", "k"], "K-th percentile (0–1)"),
    fn_def!("QUARTILE", ["array", "quart"], "Quartile 0–4"),
    // Finance
    fn_def!(
        "PV",
        ["rate", "nper", "pmt", "[fv]", "[type]"],
        "Present value"
    ),
    fn_def!(
        "FV",
        ["rate", "nper", "pmt", "[pv]", "[type]"],
        "Future value"
    ),
    fn_def!(
        "PMT",
        ["rate", "nper", "pv", "[fv]", "[type]"],
        "Periodic payment"
    ),
    fn_def!(
        "NPER",
        ["rate", "pmt", "pv", "[fv]", "[type]"],
        "Number of periods"
    ),
    fn_def!(
        "RATE",
        ["nper", "pmt", "pv", "[fv]", "[type]", "[guess]"],
        "Interest rate per period"
    ),
    fn_def!(
        "NPV",
        ["rate", "value1", "[value2, ...]"],
        "Net present value"
    ),
    fn_def!("IRR", ["values", "[guess]"], "Internal rate of return"),
    fn_def!(
        "MIRR",
        ["values", "finance_rate", "reinvest_rate"],
        "Modified IRR"
    ),
    fn_def!(
        "IPMT",
        ["rate", "per", "nper", "pv", "[fv]", "[type]"],
        "Interest portion of payment"
    ),
    fn_def!(
        "PPMT",
        ["rate", "per", "nper", "pv", "[fv]", "[type]"],
        "Principal portion of payment"
    ),
    fn_def!(
        "SLN",
        ["cost", "salvage", "life"],
        "Straight-line depreciation"
    ),
    fn_def!(
        "DDB",
        ["cost", "salvage", "life", "period", "[factor]"],
        "Double-declining depreciation"
    ),
    fn_def!("EFFECT", ["nominal_rate", "npery"], "Effective annual rate"),
    fn_def!("NOMINAL", ["effect_rate", "npery"], "Nominal annual rate"),
    fn_def!(
        "CUMIPMT",
        ["rate", "nper", "pv", "start", "end", "type"],
        "Cumulative interest paid"
    ),
    fn_def!(
        "CUMPRINC",
        ["rate", "nper", "pv", "start", "end", "type"],
        "Cumulative principal paid"
    ),
    // Volatile
    fn_def!("RAND", [], "Random float in [0, 1)"),
    fn_def!(
        "RANDBETWEEN",
        ["lo", "hi"],
        "Random integer between lo and hi (inclusive)"
    ),
    // Constants
    fn_def!("TRUE", [], "Boolean true"),
    fn_def!("FALSE", [], "Boolean false"),
    fn_def!("PI", [], "π ≈ 3.14159"),
];

fn find_fn_def(name: &str) -> Option<&'static FnDef> {
    FN_DEFS.iter().find(|f| f.name == name)
}

// ── Parsing ───────────────────────────────────────────────────────────────────

/// If the cursor is inside a function call, return (FUNCTION_NAME, arg_index).
fn parse_fn_context(buf: &str, cursor_pos: usize) -> Option<(String, usize)> {
    let buf = &buf[..cursor_pos.min(buf.len())];
    let mut depth = 0i32;
    let mut arg_idx = 0usize;

    for (i, c) in buf.char_indices().rev() {
        match c {
            ')' => depth += 1,
            '(' => {
                if depth > 0 {
                    depth -= 1;
                } else {
                    let before = &buf[..i];
                    let name_start = before
                        .rfind(|c: char| !c.is_alphanumeric() && c != '_')
                        .map(|p| p + 1)
                        .unwrap_or(0);
                    let name = before[name_start..].to_uppercase();
                    if name.is_empty() {
                        return None;
                    }
                    return Some((name, arg_idx));
                }
            }
            ',' if depth == 0 => arg_idx += 1,
            _ => {}
        }
    }
    None
}

/// Return the identifier prefix being typed at `cursor_pos` (uppercased).
fn parse_fn_prefix(buf: &str, cursor_pos: usize) -> Option<String> {
    let buf = &buf[..cursor_pos.min(buf.len())];
    let name_start = buf
        .rfind(|c: char| !c.is_alphanumeric() && c != '_')
        .map(|p| p + 1)
        .unwrap_or(0);
    let prefix = buf[name_start..].to_uppercase();
    // Must start with a letter (not a digit)
    if prefix.is_empty()
        || prefix
            .chars()
            .next()
            .map(|c| c.is_ascii_digit())
            .unwrap_or(true)
    {
        None
    } else {
        Some(prefix)
    }
}

// ── Screen position ───────────────────────────────────────────────────────────

/// Compute the screen (x, y, cell_height) of the cursor cell within `grid_area`.
/// Returns None if the cursor is scrolled off screen.
fn cursor_screen_rect(grid_area: Rect, state: &RenderState<'_>) -> Option<(u16, u16, u16)> {
    let sheet = state.workbook.active();
    let cursor = state.input.cursor;
    let viewport = state.input.viewport;

    let freeze_rows = sheet.freeze_rows as u16;
    let freeze_row_sep: u16 = if freeze_rows > 0 { 1 } else { 0 };
    let data_y = grid_area.y + 1 + freeze_rows + freeze_row_sep;

    // Compute cursor row screen_y by summing row heights
    let mut screen_y = data_y;
    if cursor.row < viewport.top_row {
        return None; // scrolled above
    }
    for r in viewport.top_row..cursor.row {
        let rh = sheet.row_height(r).max(1);
        screen_y += rh;
        if screen_y >= grid_area.y + grid_area.height {
            return None; // off screen below
        }
    }
    let cell_h = sheet.row_height(cursor.row).max(1);

    // Compute cursor col screen_x
    let freeze_cols = sheet.freeze_cols;
    let mut screen_x = grid_area.x + ROW_GUTTER_WIDTH;

    // Check if cursor is in frozen columns
    if cursor.col < freeze_cols {
        for fc in 0..cursor.col {
            screen_x += sheet.col_width(fc).max(MIN_COL_WIDTH);
        }
        return Some((screen_x, screen_y, cell_h));
    }

    // Sum frozen col widths + separator
    for fc in 0..freeze_cols {
        screen_x += sheet.col_width(fc).max(MIN_COL_WIDTH);
    }
    if freeze_cols > 0 {
        screen_x += 1; // separator
    }

    // Scrollable columns
    let scroll_left = viewport.left_col.max(freeze_cols);
    if cursor.col < scroll_left {
        return None; // in frozen section but we already handled that
    }
    for c in scroll_left..cursor.col {
        let cw = sheet.col_width(c).max(MIN_COL_WIDTH);
        screen_x += cw;
        if screen_x >= grid_area.x + grid_area.width {
            return None; // off screen right
        }
    }

    Some((screen_x, screen_y, cell_h))
}

// ── Render ────────────────────────────────────────────────────────────────────

pub fn render(frame: &mut Frame, grid_area: Rect, state: &RenderState<'_>) {
    // Only in Insert mode, only for formulas
    if !matches!(state.input.mode, Mode::Insert { .. }) {
        return;
    }
    let buf = &state.input.edit_buffer;
    if !buf.starts_with('=') || buf.len() <= 1 {
        return;
    }

    let cursor_pos = state.input.edit_cursor_pos;

    // Determine what to show
    enum HintMode {
        Signature { fn_name: String, arg_idx: usize },
        Names { prefix: String },
    }

    let hint = if let Some((fn_name, arg_idx)) = parse_fn_context(buf, cursor_pos) {
        HintMode::Signature { fn_name, arg_idx }
    } else if let Some(prefix) = parse_fn_prefix(&buf[1..], cursor_pos.saturating_sub(1)) {
        // We skip the leading '=' when looking for a prefix
        if prefix.is_empty() {
            return;
        }
        HintMode::Names { prefix }
    } else {
        return;
    };

    // Get cursor screen position
    let (cell_x, cell_y, cell_h) = match cursor_screen_rect(grid_area, state) {
        Some(r) => r,
        None => return,
    };

    // Colors
    let theme = &state.config.theme;
    let cursor_c = parse_hex_color(&theme.cursor_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let border_c = darken(header_fg, 0.55);
    let teal = Color::Rgb(45, 212, 191);

    let bg = header_bg;
    let normal_style = Style::default().fg(header_fg).bg(bg);
    let fn_style = Style::default()
        .fg(cursor_c)
        .bg(bg)
        .add_modifier(Modifier::BOLD);
    let arg_style = Style::default()
        .fg(teal)
        .bg(bg)
        .add_modifier(Modifier::BOLD);
    let dim_style = Style::default().fg(darken(header_fg, 0.55)).bg(bg);
    let desc_style = Style::default().fg(header_fg).bg(bg);

    match hint {
        // ── Signature popup ────────────────────────────────────────────────────
        HintMode::Signature { fn_name, arg_idx } => {
            let fn_def = match find_fn_def(&fn_name) {
                Some(d) => d,
                None => return,
            };

            // Build signature line spans
            let mut sig_spans: Vec<Span> =
                vec![Span::styled(format!(" {}(", fn_def.name), fn_style)];
            for (i, &arg) in fn_def.args.iter().enumerate() {
                if i > 0 {
                    sig_spans.push(Span::styled(", ", dim_style));
                }
                if i == arg_idx {
                    sig_spans.push(Span::styled(arg, arg_style));
                } else {
                    sig_spans.push(Span::styled(arg, dim_style));
                }
            }
            sig_spans.push(Span::styled(") ", fn_style));

            let sig_text: String = fn_def
                .args
                .iter()
                .enumerate()
                .fold(format!(" {}(", fn_def.name), |acc, (i, &a)| {
                    acc + if i > 0 { ", " } else { "" } + a
                })
                + ") ";
            let desc_text = format!(" {} ", fn_def.desc);

            let inner_w = sig_text
                .len()
                .max(desc_text.len())
                .max(fn_def.name.len() + 4) as u16;
            let popup_w = (inner_w + 2).min(grid_area.width.saturating_sub(2));
            let popup_h = 4u16; // border top + sig + desc + border bottom

            let (popup_x, popup_y) =
                place_popup(cell_x, cell_y, cell_h, popup_w, popup_h, grid_area);

            let popup_rect = Rect {
                x: popup_x,
                y: popup_y,
                width: popup_w,
                height: popup_h,
            };
            if popup_rect.width < 6 || popup_rect.height < 3 {
                return;
            }

            let lines = vec![
                Line::from(sig_spans),
                Line::from(vec![Span::styled(desc_text, desc_style)]),
            ];

            let block = Block::default()
                .title(format!(" {} ", fn_def.name))
                .borders(Borders::ALL)
                .border_style(Style::default().fg(border_c))
                .title_style(fn_style)
                .style(normal_style);

            frame.render_widget(Clear, popup_rect);
            frame.render_widget(Paragraph::new(lines).block(block), popup_rect);
        }

        // ── Name list popup ────────────────────────────────────────────────────
        HintMode::Names { prefix } => {
            let matches: Vec<&str> = FN_NAMES
                .iter()
                .filter(|&&n| n.starts_with(prefix.as_str()))
                .copied()
                .take(MAX_NAME_HINTS)
                .collect();

            if matches.is_empty() {
                return;
            }

            // Build rows: NAME  description
            let name_col = matches.iter().map(|n| n.len()).max().unwrap_or(3) as u16;
            let desc_col = matches
                .iter()
                .filter_map(|&n| find_fn_def(n))
                .map(|d| d.desc.len())
                .max()
                .unwrap_or(10) as u16;

            // inner = 1 pad + name_col + 2 sep + desc_col + 1 pad
            let inner_w = (1 + name_col + 2 + desc_col + 1).min(grid_area.width.saturating_sub(4));
            let popup_w = inner_w + 2;
            let popup_h = (matches.len() as u16 + 2).min(grid_area.height.saturating_sub(2));

            let (popup_x, popup_y) =
                place_popup(cell_x, cell_y, cell_h, popup_w, popup_h, grid_area);

            let popup_rect = Rect {
                x: popup_x,
                y: popup_y,
                width: popup_w,
                height: popup_h,
            };
            if popup_rect.width < 8 || popup_rect.height < 3 {
                return;
            }

            let max_name_w = (inner_w.saturating_sub(4)) / 2;
            let max_desc_w = inner_w.saturating_sub(max_name_w + 4);

            let lines: Vec<Line> = matches
                .iter()
                .map(|&name| {
                    let desc = find_fn_def(name).map(|d| d.desc).unwrap_or("");
                    let name_pad = " ".repeat((max_name_w as usize).saturating_sub(name.len()));
                    let desc_trunc: String = desc.chars().take(max_desc_w as usize).collect();
                    Line::from(vec![
                        Span::styled(" ", dim_style),
                        Span::styled(name, fn_style),
                        Span::styled(name_pad, normal_style),
                        Span::styled("  ", dim_style),
                        Span::styled(desc_trunc, desc_style),
                    ])
                })
                .collect();

            let block = Block::default()
                .title(" formula ")
                .borders(Borders::ALL)
                .border_style(Style::default().fg(border_c))
                .title_style(Style::default().fg(header_fg).add_modifier(Modifier::BOLD))
                .style(normal_style);

            frame.render_widget(Clear, popup_rect);
            frame.render_widget(Paragraph::new(lines).block(block), popup_rect);
        }
    }
}

/// Choose x, y for a popup of given size so it appears below (or above) the cursor cell,
/// clamped inside `grid_area`.
fn place_popup(
    cell_x: u16,
    cell_y: u16,
    cell_h: u16,
    popup_w: u16,
    popup_h: u16,
    grid_area: Rect,
) -> (u16, u16) {
    let grid_bottom = grid_area.y + grid_area.height;

    // Prefer below the cell
    let y_below = cell_y + cell_h;
    let y = if y_below + popup_h <= grid_bottom {
        y_below
    } else {
        // Place above
        cell_y.saturating_sub(popup_h)
    };

    // Clamp x so the popup doesn't go past the right edge
    let x = cell_x.min(grid_area.x + grid_area.width.saturating_sub(popup_w));

    (x, y.max(grid_area.y))
}
