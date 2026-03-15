use crate::{darken, parse_hex_color, RenderState};
use asat_input::Mode;
use ratatui::{
    layout::Rect,
    style::{Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Clear, Paragraph},
    Frame,
};

// ── Hint definitions ──────────────────────────────────────────────────────────

struct Hint {
    keys: &'static str,
    desc: &'static str,
}

impl Hint {
    const fn new(keys: &'static str, desc: &'static str) -> Self {
        Hint { keys, desc }
    }
}

const NORMAL_HINTS: &[Hint] = &[
    // Navigation
    Hint::new("h/j/k/l", "move left/down/up/right"),
    Hint::new("w / b", "next / prev non-empty cell →"),
    Hint::new("W / B", "next / prev non-empty cell ↓"),
    Hint::new("} / {", "next / prev paragraph"),
    Hint::new("gg / G", "first / last row"),
    Hint::new("0 / $", "first / last column"),
    Hint::new("H / M / L", "top / mid / bottom of view"),
    Hint::new("^d / ^u", "page down / up"),
    Hint::new("zz/zt/zb", "center / top / bottom row"),
    // Editing
    Hint::new("i / Enter", "edit cell"),
    Hint::new("s / cc", "clear cell + edit"),
    Hint::new("r", "replace cell"),
    Hint::new("o / O", "open row below / above"),
    Hint::new("x / D", "clear cell"),
    Hint::new("dd", "delete row"),
    Hint::new("~", "toggle case"),
    // Column / row sizing
    Hint::new(">> / <<", "widen / narrow column"),
    Hint::new("=", "auto-fit column width"),
    Hint::new("+ / -", "taller / shorter row"),
    Hint::new("_", "reset row height"),
    // Yank / paste
    Hint::new("yy", "yank row"),
    Hint::new("p / P", "paste after / before"),
    // Macros
    Hint::new("q{a-z}", "start recording macro"),
    Hint::new("@{a-z}", "play macro"),
    Hint::new("@@", "replay last macro"),
    // Marks
    Hint::new("m{a-z}", "set mark"),
    Hint::new("'{a-z}", "jump to mark"),
    Hint::new("''", "jump to previous position"),
    // Undo
    Hint::new("u / ^r", "undo / redo"),
    // Visual modes
    Hint::new("v", "visual select"),
    Hint::new("V", "V-ROW select"),
    Hint::new("^v", "V-COL select"),
    // Commands / search
    Hint::new(":", "command mode"),
    Hint::new("/ / ?", "search forward / backward"),
    Hint::new("n / N", "next / prev match"),
    Hint::new("*", "search cell under cursor"),
    // Sheets
    Hint::new("gt / gT", "next / prev sheet"),
];

const VISUAL_HINTS: &[Hint] = &[
    Hint::new("h/j/k/l", "extend selection"),
    Hint::new("w/b/W/B", "extend by non-empty cell"),
    Hint::new("} / {", "extend by paragraph"),
    Hint::new("0 / $", "first / last column"),
    Hint::new("gg / G", "first / last row"),
    Hint::new("d / x", "delete selection"),
    Hint::new("y", "yank selection"),
    Hint::new("v / V", "swap mode / exit"),
    Hint::new("Esc", "exit visual"),
];

const INSERT_HINTS: &[Hint] = &[
    Hint::new("Enter", "confirm, move down"),
    Hint::new("Tab", "confirm, move right"),
    Hint::new("Esc", "confirm, stay"),
    Hint::new("←/→", "move cursor in cell"),
    Hint::new("Home/End", "start / end of cell"),
    Hint::new("Backspace", "delete left"),
    Hint::new("Del", "delete right"),
];

// ── Render ────────────────────────────────────────────────────────────────────

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    // Decide which hints to show and what the title should be
    let mode = &state.input.mode;
    let prefix = state.input.key_prefix();

    // Hide in modes where which-key would be distracting or wrong
    if matches!(
        mode,
        Mode::Command | Mode::Search { .. }
                     | Mode::Insert { .. }
                     | Mode::FormulaSelect { .. }  // formula ref picker has its own UI
                     | Mode::Welcome | Mode::FileFind | Mode::RecentFiles | Mode::ThemeManager
    ) {
        return;
    }
    if matches!(mode, Mode::Normal) && prefix.is_empty() {
        return;
    }

    let base_hints: &[Hint] = match mode {
        Mode::Insert { .. } => INSERT_HINTS,
        Mode::Visual { .. } | Mode::VisualLine => VISUAL_HINTS,
        _ => NORMAL_HINTS,
    };

    // In non-Normal modes show all hints; in Normal show only prefix completions
    let filtered: Vec<&Hint> = if prefix.is_empty() {
        base_hints.iter().collect()
    } else {
        base_hints
            .iter()
            .filter(|h| h.keys.starts_with(&prefix))
            .collect()
    };

    if filtered.is_empty() {
        return;
    }

    // Title shows current prefix or mode name
    let title = if prefix.is_empty() {
        mode.name().to_string()
    } else {
        format!(" {} ", prefix)
    };

    // Measure content to size the panel exactly
    // 1 space pad + key col + 1 space sep + desc col + 1 border each side = key+desc+4
    let key_col = filtered.iter().map(|h| h.keys.len()).max().unwrap_or(1);
    let desc_col = filtered.iter().map(|h| h.desc.len()).max().unwrap_or(1);
    let title_min = title.len() + 4; // borders + spaces
    let content_w = 1 + key_col + 1 + desc_col + 1; // pad + key + sep + desc + pad
    let inner_w = content_w.max(title_min);
    let panel_w = (inner_w + 2) as u16; // +2 for left/right borders
                                        // Cap at terminal width
    let panel_w = panel_w.min(area.width);

    let theme = &state.config.theme;
    let cursor = parse_hex_color(&theme.cursor_bg);
    let hdr_fg = parse_hex_color(&theme.header_fg);
    let cell_bg = parse_hex_color(&theme.cell_bg);
    let border_c = darken(hdr_fg, 0.6);

    let key_style = Style::default().fg(cursor).add_modifier(Modifier::BOLD);
    let desc_style = Style::default().fg(hdr_fg);
    let sep_style = Style::default().fg(border_c);

    let lines: Vec<Line> = filtered
        .iter()
        .map(|h| {
            Line::from(vec![
                Span::styled(" ", sep_style),
                Span::styled(format!("{:<width$}", h.keys, width = key_col), key_style),
                Span::styled(" ", sep_style),
                Span::styled(h.desc.to_string(), desc_style),
            ])
        })
        .collect();

    let panel_height = (filtered.len() as u16 + 2).min(area.height);

    // Position: bottom-right corner of the grid area
    let x = area.x + area.width.saturating_sub(panel_w);
    let y = area.y + area.height.saturating_sub(panel_height);
    let rect = Rect {
        x,
        y,
        width: panel_w,
        height: panel_height,
    };

    let block = Block::default()
        .title(format!(" {} ", title))
        .borders(Borders::ALL)
        .border_style(Style::default().fg(border_c))
        .title_style(Style::default().fg(cursor).add_modifier(Modifier::BOLD))
        .style(Style::default().bg(cell_bg));

    frame.render_widget(Clear, rect);
    frame.render_widget(Paragraph::new(lines).block(block), rect);
}
