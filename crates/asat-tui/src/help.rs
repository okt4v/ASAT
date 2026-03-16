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

const KEYBIND_SECTIONS: &[Section] = &[
    Section {
        title: "Navigation",
        entries: &[
            ("h/j/k/l", "move left/down/up/right"),
            ("w / b", "next / prev non-empty cell →"),
            ("W / B", "next / prev non-empty cell ↓"),
            ("} / {", "next / prev paragraph"),
            ("gg / G", "first / last row"),
            ("0 / $", "first / last column"),
            ("H / M / L", "top / mid / bottom of view"),
            ("^d / ^u", "page down / up"),
            ("zz / zt / zb", "center / top / bottom row"),
        ],
    },
    Section {
        title: "Editing",
        entries: &[
            ("i / Enter", "edit cell"),
            ("s / cc", "clear cell + edit"),
            ("ci\"/(/{[", "change inner text object"),
            (".", "repeat last change"),
            ("r", "replace cell"),
            ("o / O", "open row below / above"),
            ("x / D", "clear cell"),
            ("dd", "delete row"),
            ("dc", "clear cell"),
            ("dC", "delete column"),
            ("dj / dk", "delete row below / above"),
            ("~", "toggle case"),
        ],
    },
    Section {
        title: "Column / Row",
        entries: &[
            (">> / <<", "widen / narrow column"),
            ("=", "auto-fit column width"),
            ("+ / -", "taller / shorter row"),
            ("_", "reset row height"),
        ],
    },
    Section {
        title: "Yank / Paste",
        entries: &[
            ("yy / yr", "yank row"),
            ("yc", "yank cell"),
            ("yC", "yank column"),
            ("yj / yk", "yank row below / above"),
            ("p / P", "paste after / before"),
        ],
    },
    Section {
        title: "Macros",
        entries: &[
            ("q{a-z}", "start recording macro"),
            ("@{a-z}", "play macro"),
            ("@@", "replay last macro"),
        ],
    },
    Section {
        title: "Marks",
        entries: &[
            ("m{a-z}", "set mark"),
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
            ("g{A-Z}", "jump to column"),
        ],
    },
    Section {
        title: "Search",
        entries: &[
            ("/ / ?", "search forward / backward"),
            ("n / N", "next / prev match"),
            ("*", "search cell under cursor"),
        ],
    },
    Section {
        title: "Visual",
        entries: &[
            ("v", "visual select"),
            ("V", "V-ROW select"),
            ("^v", "V-COL select"),
            ("d / x", "delete selection"),
            ("y", "yank selection"),
            ("M", "merge selection"),
            ("^f / ^e", "auto-fill series down / right"),
        ],
    },
    Section {
        title: "Ex-Commands",
        entries: &[
            (":w", "save"),
            (":q / :q!", "quit / force quit"),
            (":wq", "save and quit"),
            (":e <file>", "open file"),
            (":tabnew", "new sheet"),
            (":tabclose", "close sheet"),
            (":help / :h", "open this help screen"),
            (":plugins", "open plugin manager"),
            (":theme", "open theme picker"),
            (":goto <cell>", "jump to cell (e.g. B15)"),
            (":sort asc/desc", "sort rows by column"),
            (":filter <col> <op> <val>", "filter rows"),
            (":filter off", "clear filter"),
            (":bold / :italic", "toggle bold / italic"),
            (":fg / :bg <color>", "set foreground / background"),
            (":fmt <spec>", "number format"),
            (":cf <range> cond val", "conditional format"),
            (":filldown / :fillright", "fill selection"),
            (":name <n> <range>", "define named range"),
            (":plugin reload", "reload init.py"),
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
        Span::styled(" ❯ ", Style::default().fg(amber).add_modifier(Modifier::BOLD)),
        Span::styled(query.as_str(), Style::default().fg(Color::White)),
        Span::styled("█", Style::default().fg(dim)),
    ]);
    frame.render_widget(Clear, rows[1]);
    frame.render_widget(Paragraph::new(filter_line).block(filter_block), rows[1]);

    // ── Content area ─────────────────────────────────────────────────────────
    let content_area = rows[2];
    let scroll = state.input.help_scroll;

    match state.input.help_tab {
        0 => render_keybindings(frame, content_area, query, scroll, amber, header_fg, dim, bg),
        _ => render_formulas(frame, content_area, query, scroll, amber, header_fg, dim, bg),
    }
}

fn render_keybindings(
    frame: &mut Frame,
    area: Rect,
    query: &str,
    scroll: usize,
    amber: ratatui::style::Color,
    header_fg: ratatui::style::Color,
    dim: ratatui::style::Color,
    bg: ratatui::style::Color,
) {
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
        lines.push(Line::from(vec![
            Span::styled(
                format!("  {}  ", section.title),
                Style::default()
                    .fg(amber)
                    .add_modifier(Modifier::BOLD | Modifier::UNDERLINED),
            ),
        ]));
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
    frame.render_widget(
        Paragraph::new(shown).style(Style::default().bg(bg)),
        area,
    );
}

fn render_formulas(
    frame: &mut Frame,
    area: Rect,
    query: &str,
    scroll: usize,
    amber: ratatui::style::Color,
    header_fg: ratatui::style::Color,
    dim: ratatui::style::Color,
    bg: ratatui::style::Color,
) {
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
    frame.render_widget(
        Paragraph::new(shown).style(Style::default().bg(bg)),
        area,
    );
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
