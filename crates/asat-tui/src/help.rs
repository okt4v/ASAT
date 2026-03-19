use crate::{darken, parse_hex_color, RenderState};
use asat_input::FN_NAMES;
use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, BorderType, Borders, Clear, Paragraph},
    Frame,
};

// ── Static keybinding data ────────────────────────────────────────────────────

struct Section {
    title: &'static str,
    entries: &'static [(&'static str, &'static str)],
}

struct HelpColors {
    amber: Color,
    header_fg: Color,
    dim: Color,
    bg: Color,
}

const KEYBIND_SECTIONS: &[Section] = &[
    Section {
        title: "Navigation",
        entries: &[
            ("h/j/k/l", "move left/down/up/right (count prefix)"),
            ("w / b", "next / prev non-empty cell →"),
            ("W / B", "next / prev non-empty cell ↓"),
            ("e", "next non-empty cell → (alias w)"),
            ("} / {", "next / prev paragraph (empty row)"),
            ("gg / G", "first / last row"),
            ("0 / Home", "first column"),
            ("$ / End", "last column"),
            ("H / M / L", "top / mid / bottom of view"),
            ("^d / ^u", "page down / up"),
            ("^f / ^b", "page down / up (alias)"),
            ("PgDn / PgUp", "page down / up"),
            ("zz / zt / zb", "center / top / bottom row"),
            ("g{A-Z}", "jump to column by letter"),
            ("Tab / BackTab", "move right / left"),
        ],
    },
    Section {
        title: "Editing",
        entries: &[
            ("i / Enter / F2", "edit cell"),
            ("a", "edit cell (append mode)"),
            ("s / cc", "clear cell + edit"),
            ("ci\"/({[<'", "change inner text object"),
            (".", "repeat last change"),
            ("r", "replace cell (single-cell edit)"),
            ("o / O", "open row below / above + edit"),
            ("x / D", "clear cell content"),
            ("dd (count)", "delete row(s)"),
            ("dc", "clear cell (alias for x)"),
            ("dC", "delete column"),
            ("dj / dk (count)", "delete row below / above"),
            ("~", "toggle case (text cells)"),
            ("^a / ^x", "increment / decrement number"),
            ("J", "join cell below into current"),
            ("U", "unmerge cell under cursor"),
            ("gw", "toggle line-wrap on cell"),
        ],
    },
    Section {
        title: "Column / Row",
        entries: &[
            (">> / << (count)", "widen / narrow column"),
            ("=", "auto-fit column width"),
            ("+ / - (count)", "taller / shorter row"),
            ("_", "reset row height (auto-fit)"),
        ],
    },
    Section {
        title: "Yank / Paste",
        entries: &[
            ("yy / yr (count)", "yank row(s)"),
            ("yc", "yank cell"),
            ("yC", "yank column"),
            ("yj / yk", "yank row below / above"),
            ("yS", "copy cell style"),
            ("p / P", "paste after / before"),
            ("pS", "paste style to cell"),
        ],
    },
    Section {
        title: "Macros",
        entries: &[
            ("q{a-z}", "start recording macro into register"),
            ("q (while recording)", "stop recording"),
            ("@{a-z} (count)", "play macro from register"),
            ("@@ (count)", "replay last macro"),
        ],
    },
    Section {
        title: "Marks",
        entries: &[
            ("m{a-z}", "set mark at cursor"),
            ("'{a-z}", "jump to mark"),
            ("''", "jump to previous position"),
        ],
    },
    Section {
        title: "Go-To",
        entries: &[
            ("gd", "go to first cell ref in formula"),
            ("gt / gT", "next / prev sheet"),
            ("gg", "first row"),
            ("G", "last row"),
            ("g{A-Z}", "jump to column by letter"),
        ],
    },
    Section {
        title: "Search",
        entries: &[
            ("/ / ?", "search forward / backward"),
            ("n / N (count)", "next / prev match"),
            ("*", "search cell content under cursor"),
            ("Esc", "clear search highlights"),
        ],
    },
    Section {
        title: "Visual Mode (v / V / ^v)",
        entries: &[
            ("v", "visual select (cell range)"),
            ("V", "V-ROW select (full rows)"),
            ("^v", "V-COL select (full columns)"),
            ("d / x / Del", "delete selection"),
            ("c / s", "clear selection + enter edit"),
            ("y", "yank selection"),
            ("M", "merge selection into one cell"),
            ("S", "insert =SUM() below/right"),
            ("^d", "fill down (copy anchor row)"),
            ("^r", "fill right (copy anchor col)"),
            ("^f", "auto-fill series (smart direction)"),
            ("> / < (count)", "widen / narrow columns"),
            (":", "enter command for selection range"),
            ("v / V", "swap visual mode / exit"),
            ("Esc", "exit visual mode"),
        ],
    },
    Section {
        title: "Insert Mode",
        entries: &[
            ("Enter", "confirm + move down"),
            ("Tab", "confirm + move right"),
            ("Esc", "confirm + stay in place"),
            ("Left / Right", "move edit cursor"),
            ("^a / ^e", "jump to start / end of buffer"),
            ("Backspace", "delete character left"),
            ("Delete", "delete character right"),
            ("^w", "delete word backward"),
            ("^u", "delete to start of buffer"),
            ("^k", "delete to end of buffer"),
            ("^v", "paste from system clipboard"),
            ("^r", "formula ref picker (when editing =...)"),
            ("Tab / BackTab", "cycle formula completions"),
        ],
    },
    Section {
        title: "Ex-Commands — File",
        entries: &[
            (":w", "save to current file"),
            (":w <file>", "save to a new file"),
            (":q", "quit (warns if unsaved)"),
            (":q!", "force quit without saving"),
            (":wq / :x", "save and quit"),
            (":e <file>", "open file"),
            (":home", "return to welcome screen"),
        ],
    },
    Section {
        title: "Ex-Commands — Sheets",
        entries: &[
            (":tabnew / :tabedit", "new sheet"),
            (":tabclose", "close current sheet"),
        ],
    },
    Section {
        title: "Ex-Commands — Row / Column",
        entries: &[
            (":ir", "insert row below cursor"),
            (":dr", "delete current row"),
            (":ic", "insert column left"),
            (":icr", "insert column right"),
            (":dc", "delete current column"),
            (":cw <N>", "set column width to N"),
            (":rh <N>", "set row height to N"),
        ],
    },
    Section {
        title: "Ex-Commands — Formatting",
        entries: &[
            (":bold", "toggle bold"),
            (":italic", "toggle italic"),
            (":underline", "toggle underline"),
            (":strike", "toggle strikethrough"),
            (":fg <color>", "set foreground colour"),
            (":bg <color>", "set background colour"),
            (":hl <color>", "highlight (bg + auto fg)"),
            (":hl", "clear highlight"),
            (":align <l/c/r>", "set text alignment"),
            (":fmt <spec>", "number format (%, $, 0.00, date…)"),
            (":copystyle / :cs", "copy style / clear styles"),
            (":pastestyle", "paste style to selection"),
            (":wrap", "toggle line-wrap"),
        ],
    },
    Section {
        title: "Ex-Commands — Data",
        entries: &[
            (":sort asc/desc", "sort rows by cursor column"),
            (":s /pat/repl/g", "find & replace in text cells"),
            (":filter <c> <op> <v>", "filter rows (e.g. :filter A >100)"),
            (":filter off", "clear row filter"),
            (":transpose", "transpose visual selection"),
            (":dedup", "remove duplicate rows"),
            (":filldown / :fillright", "fill selection down / right"),
            (":merge / :unmerge", "merge or unmerge cells"),
            (":name <n> <range>", "define named range"),
            (":colfmt <op> <v> <color>", "conditional format rule"),
        ],
    },
    Section {
        title: "Ex-Commands — Other",
        entries: &[
            (":help / :h", "open searchable help screen"),
            (":goto <cell>", "jump to cell (e.g. :goto B15)"),
            (":theme / :theme <name>", "open theme picker / apply theme"),
            (":set <option>", "set an option"),
            (":note <text>", "set note (= view, ! clear)"),
            (":freeze rows/cols <N>", "freeze panes"),
            (":freeze off", "clear freeze panes"),
            (":J", "join cell below into current"),
            (":plugin reload", "reload init.py"),
            (":plugins", "open plugin manager"),
        ],
    },
];

// ── Render ────────────────────────────────────────────────────────────────────

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let theme = &state.config.theme;
    let amber = parse_hex_color(&theme.cursor_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let bg = parse_hex_color(&theme.cell_bg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let dim = darken(header_fg, 0.55);
    let colors = HelpColors {
        amber,
        header_fg,
        dim,
        bg,
    };
    let border_c = darken(header_fg, 0.4);

    frame.render_widget(Paragraph::new("").style(Style::default().bg(bg)), area);

    // ── Outer layout: tab bar (1) + filter bar (3) + content (rest) ──────────
    let rows = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1),
            Constraint::Length(3),
            Constraint::Min(1),
        ])
        .split(area);

    // ── Tab bar ───────────────────────────────────────────────────────────────
    let tabs = ["  Keybindings  ", "  Formulas  "];
    let tab_line: Vec<Span> = tabs
        .iter()
        .enumerate()
        .flat_map(|(i, label)| {
            let active = i == state.input.help_tab;
            let style = if active {
                Style::default()
                    .fg(amber)
                    .bg(header_bg)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(dim).bg(bg)
            };
            [Span::styled(*label, style), Span::raw(" ")]
        })
        .collect();
    frame.render_widget(
        Paragraph::new(Line::from(tab_line)).style(Style::default().bg(bg)),
        rows[0],
    );

    // ── Filter / search bar ───────────────────────────────────────────────────
    let query = &state.input.help_query;
    let filter_block = Block::default()
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(border_c))
        .style(Style::default().bg(header_bg))
        .title(Span::styled(
            " Search (Tab to switch tabs, q/Esc to close) ",
            Style::default().fg(dim),
        ));
    let filter_line = Line::from(vec![
        Span::styled(
            " ❯ ",
            Style::default().fg(amber).add_modifier(Modifier::BOLD),
        ),
        Span::styled(query.as_str(), Style::default().fg(Color::White)),
        Span::styled("█", Style::default().fg(dim)),
    ]);
    frame.render_widget(Clear, rows[1]);
    frame.render_widget(Paragraph::new(filter_line).block(filter_block), rows[1]);

    // ── Content area ─────────────────────────────────────────────────────────
    let content_area = rows[2];
    let scroll = state.input.help_scroll;

    match state.input.help_tab {
        0 => render_keybindings(frame, content_area, query, scroll, &colors),
        _ => render_formulas(frame, content_area, query, scroll, &colors),
    }
}

fn render_keybindings(frame: &mut Frame, area: Rect, query: &str, scroll: usize, c: &HelpColors) {
    let (amber, header_fg, dim, bg) = (c.amber, c.header_fg, c.dim, c.bg);
    let q = query.to_ascii_lowercase();

    // Build all visible lines
    let mut lines: Vec<Line> = Vec::new();
    for section in KEYBIND_SECTIONS {
        let matching: Vec<_> = section
            .entries
            .iter()
            .filter(|(keys, desc)| {
                q.is_empty()
                    || keys.to_ascii_lowercase().contains(&q)
                    || desc.to_ascii_lowercase().contains(&q)
            })
            .collect();
        if matching.is_empty() {
            continue;
        }
        // Section header
        lines.push(Line::from(vec![Span::styled(
            format!("  {}  ", section.title),
            Style::default()
                .fg(amber)
                .add_modifier(Modifier::BOLD | Modifier::UNDERLINED),
        )]));
        for (keys, desc) in &matching {
            lines.push(Line::from(vec![
                Span::styled("  ", Style::default()),
                Span::styled(
                    format!("{:<24}", keys),
                    Style::default().fg(amber).add_modifier(Modifier::BOLD),
                ),
                Span::styled(*desc, Style::default().fg(header_fg)),
            ]));
        }
        lines.push(Line::from(""));
    }

    if lines.is_empty() {
        lines.push(Line::from(Span::styled(
            "  No results",
            Style::default().fg(dim),
        )));
    }

    let visible_h = area.height as usize;
    let max_scroll = lines.len().saturating_sub(visible_h);
    let scroll = scroll.min(max_scroll);

    let shown: Vec<Line> = lines.into_iter().skip(scroll).take(visible_h).collect();
    frame.render_widget(Paragraph::new(shown).style(Style::default().bg(bg)), area);
}

fn render_formulas(frame: &mut Frame, area: Rect, query: &str, scroll: usize, c: &HelpColors) {
    let (amber, header_fg, dim, bg) = (c.amber, c.header_fg, c.dim, c.bg);
    let q = query.to_ascii_uppercase();

    // Build formula doc entries from FN_NAMES
    let all_fns: Vec<(&str, &str)> = FN_NAMES
        .iter()
        .filter_map(|name| {
            if q.is_empty() || name.contains(&q) {
                Some((*name, formula_hint(name)))
            } else {
                None
            }
        })
        .collect();

    let mut lines: Vec<Line> = Vec::new();
    if all_fns.is_empty() {
        lines.push(Line::from(Span::styled(
            "  No results",
            Style::default().fg(dim),
        )));
    } else {
        lines.push(Line::from(Span::styled(
            format!("  {} functions", all_fns.len()),
            Style::default().fg(dim),
        )));
        lines.push(Line::from(""));
        for (name, hint) in &all_fns {
            lines.push(Line::from(vec![
                Span::styled("  ", Style::default()),
                Span::styled(
                    format!("{:<20}", name),
                    Style::default().fg(amber).add_modifier(Modifier::BOLD),
                ),
                Span::styled(*hint, Style::default().fg(header_fg)),
            ]));
        }
    }

    let visible_h = area.height as usize;
    let max_scroll = lines.len().saturating_sub(visible_h);
    let scroll = scroll.min(max_scroll);

    let shown: Vec<Line> = lines.into_iter().skip(scroll).take(visible_h).collect();
    frame.render_widget(Paragraph::new(shown).style(Style::default().bg(bg)), area);
}

/// Brief one-line description for each formula function.
fn formula_hint(name: &str) -> &'static str {
    match name {
        "SUM" => "Sum of values",
        "AVERAGE" | "AVG" => "Average of values",
        "COUNT" => "Count numeric values",
        "COUNTA" => "Count non-empty values",
        "MIN" => "Minimum value",
        "MAX" => "Maximum value",
        "IF" => "IF(cond, val_true, val_false)",
        "AND" => "True if all args are true",
        "OR" => "True if any arg is true",
        "NOT" => "Logical negation",
        "ABS" => "Absolute value",
        "ROUND" => "ROUND(num, digits)",
        "ROUNDUP" => "Round away from zero",
        "ROUNDDOWN" => "Round toward zero",
        "FLOOR" => "Round down to multiple",
        "CEILING" => "Round up to multiple",
        "MOD" => "MOD(num, divisor) — remainder",
        "POWER" => "POWER(base, exp)",
        "SQRT" => "Square root",
        "LN" => "Natural logarithm",
        "LOG" => "LOG(num, base) — logarithm",
        "LOG10" => "Base-10 logarithm",
        "EXP" => "e raised to the power",
        "INT" => "Integer part (floor for positive)",
        "TRUNC" => "Truncate to integer",
        "SIGN" => "Sign: -1, 0, or 1",
        "LEN" => "Length of text",
        "LEFT" => "LEFT(text, n) — first n chars",
        "RIGHT" => "RIGHT(text, n) — last n chars",
        "MID" => "MID(text, start, n)",
        "TRIM" => "Remove extra whitespace",
        "UPPER" => "Convert to uppercase",
        "LOWER" => "Convert to lowercase",
        "PROPER" => "Title case",
        "CONCATENATE" | "CONCAT" => "Join text values",
        "TEXT" => "TEXT(value, format) — format number as text",
        "VALUE" => "Parse text to number",
        "FIND" => "FIND(needle, haystack[, start]) — case-sensitive",
        "SEARCH" => "SEARCH(needle, haystack[, start]) — case-insensitive",
        "SUBSTITUTE" => "SUBSTITUTE(text, old, new[, n])",
        "REPLACE" => "REPLACE(text, start, len, new)",
        "REPT" => "REPT(text, n) — repeat text",
        "ISNUMBER" => "True if value is a number",
        "ISTEXT" => "True if value is text",
        "ISBLANK" => "True if cell is empty",
        "ISERROR" => "True if value is an error",
        "IFERROR" => "IFERROR(value, fallback)",
        "ISLOGICAL" => "True if value is TRUE/FALSE",
        "TRUE" => "Boolean true",
        "FALSE" => "Boolean false",
        "PI" => "π ≈ 3.14159…",
        "SUMIF" => "SUMIF(range, criteria, sum_range)",
        "COUNTIF" => "COUNTIF(range, criteria)",
        "SUMPRODUCT" => "Sum of products of arrays",
        "LARGE" => "LARGE(range, k) — k-th largest",
        "SMALL" => "SMALL(range, k) — k-th smallest",
        "MEDIAN" => "Median value",
        "STDEV" => "Standard deviation (sample)",
        "VAR" => "Variance (sample)",
        "AVERAGEIF" => "AVERAGEIF(range, criteria, avg_range)",
        "MAXIFS" => "Max with multiple criteria",
        "MINIFS" => "Min with multiple criteria",
        "RANK" => "RANK(value, range[, order])",
        "PERCENTILE" => "PERCENTILE(range, k)",
        "QUARTILE" => "QUARTILE(range, quart)",
        "XLOOKUP" => "XLOOKUP(val, lookup, return[, missing])",
        "CHOOSE" => "CHOOSE(index, val1, val2, …)",
        "PV" => "Present value",
        "FV" => "Future value",
        "PMT" => "Loan payment amount",
        "NPER" => "Number of periods",
        "RATE" => "Interest rate per period",
        "NPV" => "Net present value",
        "IRR" => "Internal rate of return",
        "MIRR" => "Modified IRR",
        "IPMT" => "Interest payment for period",
        "PPMT" => "Principal payment for period",
        "SLN" => "Straight-line depreciation",
        "DDB" => "Double-declining balance depreciation",
        "EFFECT" => "Effective annual interest rate",
        "NOMINAL" => "Nominal annual interest rate",
        "CUMIPMT" => "Cumulative interest paid",
        "CUMPRINC" => "Cumulative principal paid",
        "NOW" => "Current date and time (Excel serial)",
        "TODAY" => "Today's date (Excel serial integer)",
        "RAND" => "Random float in [0, 1)",
        "RANDBETWEEN" => "RANDBETWEEN(lo, hi) — random integer",
        _ => "",
    }
}
