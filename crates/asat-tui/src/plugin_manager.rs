use crate::{darken, parse_hex_color, RenderState};
use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, BorderType, Borders, Clear, List, ListItem, Paragraph},
    Frame,
};

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let theme = &state.config.theme;
    let amber = parse_hex_color(&theme.cursor_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let sel_bg = parse_hex_color(&theme.selection_bg);
    let bg = parse_hex_color(&theme.cell_bg);
    let dim = darken(header_fg, 0.55);
    let border_c = darken(header_fg, 0.4);

    frame.render_widget(Paragraph::new("").style(Style::default().bg(bg)), area);

    // ── Build the list of "entries" to display ────────────────────────────────
    // Entry 0: engine status
    // Entry 1+: custom functions registered by plugins
    let custom_fns = &state.plugin_custom_fns;
    let total_entries = 1 + custom_fns.len();
    let selected = state.input.plugin_selected.min(total_entries.saturating_sub(1));

    // ── Panel ─────────────────────────────────────────────────────────────────
    let panel_w = area.width.min(80).max(50);
    let panel_h = area.height.min(30).max(10);
    let panel_x = area.x + area.width.saturating_sub(panel_w) / 2;
    let panel_y = area.y + area.height.saturating_sub(panel_h) / 2;
    let panel = Rect {
        x: panel_x,
        y: panel_y,
        width: panel_w,
        height: panel_h,
    };

    // Split into left list (60%) and right detail (40%)
    let cols = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Percentage(60), Constraint::Percentage(40)])
        .split(panel);

    // ── Left: plugin list ─────────────────────────────────────────────────────
    let list_block = Block::default()
        .title(Span::styled(
            " Plugin Manager ",
            Style::default().fg(amber).add_modifier(Modifier::BOLD),
        ))
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(amber))
        .style(Style::default().bg(header_bg));

    let hint_block = list_block.title_bottom(Span::styled(
        " r reload  q close ",
        Style::default().fg(dim),
    ));

    let mut items: Vec<ListItem> = Vec::new();

    // Entry 0: engine status
    let engine_label = if state.plugin_info.contains("not compiled") {
        "✗ Python engine (not compiled)"
    } else if state.plugin_info.contains("active") {
        "● Python engine  (active)"
    } else {
        "○ Python engine  (inactive)"
    };
    let (fg0, bg0) = if selected == 0 {
        (Color::White, sel_bg)
    } else {
        (header_fg, bg)
    };
    items.push(
        ListItem::new(Line::from(vec![
            Span::styled(
                if selected == 0 { " ▶ " } else { "   " },
                Style::default().fg(amber),
            ),
            Span::styled(
                engine_label,
                Style::default().fg(fg0).add_modifier(Modifier::BOLD),
            ),
        ]))
        .style(Style::default().bg(bg0)),
    );

    // Entry 1+: custom functions
    if custom_fns.is_empty() {
        items.push(ListItem::new(Line::from(Span::styled(
            "   No custom functions",
            Style::default().fg(dim),
        ))));
    } else {
        for (i, name) in custom_fns.iter().enumerate() {
            let idx = i + 1;
            let is_sel = selected == idx;
            let (fg, item_bg) = if is_sel {
                (Color::White, sel_bg)
            } else {
                (header_fg, bg)
            };
            items.push(
                ListItem::new(Line::from(vec![
                    Span::styled(
                        if is_sel { " ▶ " } else { "   " },
                        Style::default().fg(amber),
                    ),
                    Span::styled(
                        format!("fn {}", name),
                        Style::default().fg(fg),
                    ),
                ]))
                .style(Style::default().bg(item_bg)),
            );
        }
    }

    frame.render_widget(Clear, cols[0]);
    frame.render_widget(List::new(items).block(hint_block), cols[0]);

    // ── Right: detail pane ────────────────────────────────────────────────────
    let detail_block = Block::default()
        .title(Span::styled(" Detail ", Style::default().fg(dim)))
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(border_c))
        .style(Style::default().bg(header_bg));

    let detail_text: Vec<Line> = if selected == 0 {
        // Engine status detail
        state
            .plugin_info
            .split("  ")
            .flat_map(|chunk| {
                chunk.split(',').map(|s| {
                    Line::from(Span::styled(
                        format!(" {}", s.trim()),
                        Style::default().fg(header_fg),
                    ))
                })
            })
            .collect()
    } else {
        let fn_name = &custom_fns[selected - 1];
        vec![
            Line::from(Span::styled(
                format!(" {}", fn_name),
                Style::default().fg(amber).add_modifier(Modifier::BOLD),
            )),
            Line::from(""),
            Line::from(Span::styled(
                " Custom formula function",
                Style::default().fg(header_fg),
            )),
            Line::from(Span::styled(
                " registered by init.py",
                Style::default().fg(dim),
            )),
            Line::from(""),
            Line::from(Span::styled(
                format!(" Usage: ={fn_name}(...)"),
                Style::default().fg(header_fg),
            )),
        ]
    };

    frame.render_widget(Clear, cols[1]);
    frame.render_widget(Paragraph::new(detail_text).block(detail_block), cols[1]);
}
