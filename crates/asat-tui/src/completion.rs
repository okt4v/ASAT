use crate::{darken, is_dark_color, parse_hex_color, RenderState};
use asat_input::{get_command_completions, Mode};
use ratatui::{
    layout::Rect,
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Clear, Paragraph},
    Frame,
};

const MAX_VISIBLE: usize = 10;

pub fn render(frame: &mut Frame, screen: Rect, state: &RenderState<'_>) {
    // Only show in Command mode
    if !matches!(state.input.mode, Mode::Command) {
        return;
    }

    // ── Sub-command completions (e.g. :theme <name>) take priority ────────
    let cmd_buf = &state.input.command_buffer;
    if cmd_buf.contains(' ') {
        // We're completing a command argument
        let subcmds = &state.input.subcmd_completions;
        if !subcmds.is_empty() {
            render_subcmd_popup(
                frame,
                screen,
                state,
                subcmds,
                state.input.subcmd_completion_idx,
            );
        }
        return;
    }

    // ── Verb-level completions ─────────────────────────────────────────────
    let prefix = &state.input.completion_prefix;
    let query = if state.input.completion_idx.is_some() {
        prefix.as_str()
    } else {
        state
            .input
            .command_buffer
            .split_whitespace()
            .next()
            .unwrap_or("")
    };

    let completions = get_command_completions(query);
    if completions.is_empty() {
        return;
    }

    let selected = state.input.completion_idx;
    let visible: Vec<_> = completions.iter().enumerate().take(MAX_VISIBLE).collect();

    // ── Sizing ────────────────────────────────────────────────────────────
    let cmd_col_w = visible
        .iter()
        .map(|(_, (cmd, _))| cmd.len())
        .max()
        .unwrap_or(4) as u16;
    let desc_col_w = visible
        .iter()
        .map(|(_, (_, desc))| desc.len())
        .max()
        .unwrap_or(10) as u16;

    // inner width: 1 pad + cmd + 2 sep + desc + 1 pad
    let inner_w = (1 + cmd_col_w + 2 + desc_col_w + 1).min(screen.width.saturating_sub(4));
    let popup_w = inner_w + 2; // +2 for borders
    let popup_h = (visible.len() as u16 + 2).min(screen.height.saturating_sub(4)); // +2 for borders

    // Position: just above the command line (bottom row) on the left side
    // The command line takes the very last row(s); command bar is at screen.bottom()-1
    // We offset 2 rows up from the very bottom: 1 for command line + 1 gap
    let y = screen.y + screen.height.saturating_sub(popup_h + 2);
    let x = screen.x + 1; // one cell from the left edge (aligns under the ':' prompt)

    let popup_rect = Rect {
        x,
        y,
        width: popup_w.min(screen.width.saturating_sub(x)),
        height: popup_h,
    };
    if popup_rect.width < 8 || popup_rect.height < 3 {
        return;
    }

    // ── Colors ────────────────────────────────────────────────────────────
    let theme = &state.config.theme;
    let cursor_bg = parse_hex_color(&theme.cursor_bg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let sel_bg = parse_hex_color(&theme.selection_bg);
    let border_c = darken(header_fg, 0.55);
    let sel_fg = if is_dark_color(sel_bg) {
        Color::White
    } else {
        Color::Black
    };

    let bg = header_bg; // popup panel uses header_bg surface
    let normal_style = Style::default().fg(Color::White).bg(bg);
    let cmd_style = Style::default()
        .fg(cursor_bg)
        .bg(bg)
        .add_modifier(Modifier::BOLD);
    let desc_style = Style::default().fg(header_fg).bg(bg);
    let sel_cmd_style = Style::default()
        .fg(cursor_bg)
        .bg(sel_bg)
        .add_modifier(Modifier::BOLD);
    let sel_desc_style = Style::default().fg(sel_fg).bg(sel_bg);
    let sep_style = Style::default().fg(border_c).bg(bg);

    // ── Lines ─────────────────────────────────────────────────────────────
    let max_cmd_w = (inner_w.saturating_sub(4)) / 2;
    let max_desc_w = inner_w.saturating_sub(max_cmd_w + 4);

    let lines: Vec<Line> = visible
        .iter()
        .map(|(local_idx, (cmd, desc))| {
            let is_sel = selected.map(|s| s == *local_idx).unwrap_or(false);

            let (cs, ds, pad_s) = if is_sel {
                (sel_cmd_style, sel_desc_style, Style::default().bg(sel_bg))
            } else {
                (cmd_style, desc_style, normal_style)
            };

            let indicator = if is_sel { "▶ " } else { "  " };
            let cmd_txt = truncate(cmd, max_cmd_w as usize);
            let desc_txt = truncate(desc, max_desc_w as usize);
            let padding = " ".repeat((max_cmd_w as usize).saturating_sub(cmd_txt.len()));

            Line::from(vec![
                Span::styled(indicator, if is_sel { sel_cmd_style } else { sep_style }),
                Span::styled(cmd_txt, cs),
                Span::styled(padding, pad_s),
                Span::styled("  ", sep_style),
                Span::styled(desc_txt, ds),
            ])
        })
        .collect();

    // Show hint when more completions exist beyond MAX_VISIBLE
    let overflow = completions.len().saturating_sub(MAX_VISIBLE);

    let title = if overflow > 0 {
        format!(" completions (+{} more) ", overflow)
    } else {
        format!(
            " {} completion{} ",
            completions.len(),
            if completions.len() == 1 { "" } else { "s" }
        )
    };

    let block = Block::default()
        .title(title)
        .borders(Borders::ALL)
        .border_style(Style::default().fg(border_c))
        .title_style(Style::default().fg(header_fg).add_modifier(Modifier::BOLD))
        .style(normal_style);

    // Clear the area first so the popup draws over the grid
    frame.render_widget(Clear, popup_rect);
    frame.render_widget(Paragraph::new(lines).block(block), popup_rect);
}

fn truncate(s: &str, max: usize) -> &str {
    if max == 0 {
        return "";
    }
    if s.len() <= max {
        s
    } else {
        &s[..max.saturating_sub(1)]
    }
}

/// Popup for sub-command completions (e.g. theme names after `:theme `).
fn render_subcmd_popup(
    frame: &mut Frame,
    screen: Rect,
    state: &RenderState<'_>,
    completions: &[String],
    selected: Option<usize>,
) {
    let theme = &state.config.theme;
    let cursor_bg = parse_hex_color(&theme.cursor_bg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let sel_bg = parse_hex_color(&theme.selection_bg);
    let border_c = darken(header_fg, 0.55);
    let sel_fg = if is_dark_color(sel_bg) {
        Color::White
    } else {
        Color::Black
    };

    let visible: Vec<_> = completions.iter().enumerate().take(MAX_VISIBLE).collect();
    let max_w = visible.iter().map(|(_, s)| s.len()).max().unwrap_or(4) as u16;

    let inner_w = (max_w + 4).min(screen.width.saturating_sub(4));
    let popup_w = inner_w + 2;
    let popup_h = (visible.len() as u16 + 2).min(screen.height.saturating_sub(4));

    // Position just above command line (same anchor as main popup)
    let y = screen.y + screen.height.saturating_sub(popup_h + 2);
    let x = screen.x + 1;
    let popup_rect = Rect {
        x,
        y,
        width: popup_w.min(screen.width.saturating_sub(x)),
        height: popup_h,
    };
    if popup_rect.width < 6 || popup_rect.height < 3 {
        return;
    }

    let overflow = completions.len().saturating_sub(MAX_VISIBLE);
    let title = if overflow > 0 {
        format!(" (+{} more) ", overflow)
    } else {
        format!(" {} ", completions.len())
    };

    let lines: Vec<Line> = visible
        .iter()
        .map(|(local_idx, name)| {
            let is_sel = selected.map(|s| s == *local_idx).unwrap_or(false);
            if is_sel {
                Line::from(vec![
                    Span::styled(" ▶ ", Style::default().fg(cursor_bg).bg(sel_bg)),
                    Span::styled(
                        name.as_str(),
                        Style::default()
                            .fg(sel_fg)
                            .bg(sel_bg)
                            .add_modifier(Modifier::BOLD),
                    ),
                ])
                .style(Style::default().bg(sel_bg))
            } else {
                Line::from(vec![
                    Span::styled("   ", Style::default().bg(header_bg)),
                    Span::styled(name.as_str(), Style::default().fg(header_fg).bg(header_bg)),
                ])
            }
        })
        .collect();

    let block = Block::default()
        .title(title)
        .borders(Borders::ALL)
        .border_style(Style::default().fg(border_c))
        .title_style(Style::default().fg(cursor_bg).add_modifier(Modifier::BOLD))
        .style(Style::default().bg(header_bg));

    frame.render_widget(Clear, popup_rect);
    frame.render_widget(Paragraph::new(lines).block(block), popup_rect);
}
