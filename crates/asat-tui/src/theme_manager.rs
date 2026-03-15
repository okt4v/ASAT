use asat_config::builtin_themes;
use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, BorderType, Borders, List, ListItem, Paragraph},
    Frame,
};
use crate::{darken, parse_hex_color, RenderState};

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let themes  = builtin_themes();
    let selected = state.input.theme_selected.min(themes.len().saturating_sub(1));

    let t_cur  = &state.config.theme;
    let amber  = parse_hex_color(&t_cur.cursor_bg);
    let bg     = parse_hex_color(&t_cur.cell_bg);
    let hdr_bg = parse_hex_color(&t_cur.header_bg);
    let hdr_fg = parse_hex_color(&t_cur.header_fg);
    let border = darken(hdr_fg, 0.55);

    // Background fill
    frame.render_widget(Paragraph::new("").style(Style::default().bg(bg)), area);

    // ── Outer layout: title bar + body ────────────────────────────────────
    let outer = Layout::default()
        .direction(Direction::Vertical)
        .constraints([Constraint::Length(3), Constraint::Min(1)])
        .split(area);

    // Title bar
    let title_line = Line::from(vec![
        Span::styled(
            "  🎨  Theme Manager  ",
            Style::default().fg(amber).bg(hdr_bg).add_modifier(Modifier::BOLD),
        ),
        Span::styled(
            "  ↑/↓ or j/k to browse  ·  Enter to apply  ·  Esc to cancel  ",
            Style::default().fg(border).bg(hdr_bg),
        ),
    ]);
    frame.render_widget(
        Paragraph::new(title_line).style(Style::default().bg(hdr_bg)),
        outer[0],
    );

    // ── Body: left list + right preview ──────────────────────────────────
    let body = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Length(28), Constraint::Min(1)])
        .split(outer[1]);

    // ── Left: theme list ─────────────────────────────────────────────────
    let sel_bg  = parse_hex_color(&t_cur.selection_bg);

    let list_items: Vec<ListItem> = themes.iter().enumerate().map(|(i, t)| {
        let is_sel = i == selected;
        let dark_indicator = if t.dark { "◐" } else { "◑" };
        if is_sel {
            ListItem::new(Line::from(vec![
                Span::styled(" ▶ ", Style::default().fg(amber)),
                Span::styled(t.name, Style::default().fg(Color::White).add_modifier(Modifier::BOLD)),
                Span::styled(format!(" {}", dark_indicator), Style::default().fg(darken(hdr_fg, 0.7))),
            ])).style(Style::default().bg(sel_bg))
        } else {
            ListItem::new(Line::from(vec![
                Span::styled("   ", Style::default()),
                Span::styled(t.name, Style::default().fg(hdr_fg)),
                Span::styled(format!(" {}", dark_indicator), Style::default().fg(darken(hdr_fg, 0.5))),
            ]))
        }
    }).collect();

    let list_block = Block::default()
        .title(" Themes ")
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(border))
        .title_style(Style::default().fg(amber).add_modifier(Modifier::BOLD))
        .style(Style::default().bg(bg));

    frame.render_widget(List::new(list_items).block(list_block), body[0]);

    // ── Right: preview pane ───────────────────────────────────────────────
    let preset = &themes[selected];
    let t = &preset.config;

    let cursor_bg  = parse_hex_color(&t.cursor_bg);
    let header_bg  = parse_hex_color(&t.header_bg);
    let header_fg  = parse_hex_color(&t.header_fg);
    let cell_bg    = parse_hex_color(&t.cell_bg);
    let sel_bg     = parse_hex_color(&t.selection_bg);
    let normal_c   = parse_hex_color(&t.normal_mode_color);
    let insert_c   = parse_hex_color(&t.insert_mode_color);
    let visual_c   = parse_hex_color(&t.visual_mode_color);
    let command_c  = parse_hex_color(&t.command_mode_color);

    let preview_block = Block::default()
        .title(format!(" Preview — {} ", preset.name))
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(border))
        .title_style(Style::default().fg(amber).add_modifier(Modifier::BOLD))
        .style(Style::default().bg(bg));

    let inner = preview_block.inner(body[1]);
    frame.render_widget(preview_block, body[1]);

    if inner.width < 20 || inner.height < 8 { return; }

    // Layout within preview: mock spreadsheet + description
    let preview_rows = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(7),  // mock grid
            Constraint::Length(1),  // spacer
            Constraint::Length(3),  // description
            Constraint::Length(1),  // spacer
            Constraint::Length(1),  // mode badges
            Constraint::Min(0),
        ])
        .split(inner);

    let normal_c = parse_hex_color(&t.normal_mode_color);
    let number_c = parse_hex_color(&t.number_color);
    render_mock_grid(frame, preview_rows[0], cursor_bg, header_bg, header_fg, cell_bg, sel_bg, normal_c, number_c);

    // Description
    let badge = if preset.dark { "DARK" } else { "LIGHT" };
    frame.render_widget(
        Paragraph::new(vec![
            Line::from(vec![
                Span::styled(format!(" {} ", badge), Style::default()
                    .fg(Color::Black).bg(cursor_bg)
                    .add_modifier(Modifier::BOLD)),
                Span::styled(format!("  {}", preset.name), Style::default()
                    .fg(Color::White)
                    .add_modifier(Modifier::BOLD)),
            ]),
            Line::from(""),
            Line::from(Span::styled(
                format!(" {}", preset.description),
                Style::default().fg(parse_hex_color(&t_cur.header_fg)),
            )),
        ]).style(Style::default().bg(bg)),
        preview_rows[2],
    );

    // Mode badge preview
    let badges = Line::from(vec![
        Span::styled(" NORMAL ",  Style::default().fg(Color::Black).bg(normal_c).add_modifier(Modifier::BOLD)),
        Span::raw(" "),
        Span::styled(" INSERT ",  Style::default().fg(Color::Black).bg(insert_c).add_modifier(Modifier::BOLD)),
        Span::raw(" "),
        Span::styled(" VISUAL ",  Style::default().fg(Color::Black).bg(visual_c).add_modifier(Modifier::BOLD)),
        Span::raw(" "),
        Span::styled(" COMMAND ", Style::default().fg(Color::Black).bg(command_c).add_modifier(Modifier::BOLD)),
    ]);
    frame.render_widget(
        Paragraph::new(badges).style(Style::default().bg(bg)),
        preview_rows[4],
    );
}

/// Draw a tiny mock spreadsheet using the theme's colours.
fn render_mock_grid(
    frame: &mut Frame,
    area: Rect,
    cursor_bg:    Color,
    header_bg:    Color,
    header_fg:    Color,
    cell_bg:      Color,
    sel_bg:       Color,
    normal_color: Color,
    number_color: Color,
) {
    if area.height < 4 { return; }

    let header_style  = Style::default().fg(header_fg).bg(header_bg).add_modifier(Modifier::BOLD);
    let cursor_hdr    = Style::default().fg(Color::Black).bg(cursor_bg).add_modifier(Modifier::BOLD);
    let cell_normal   = Style::default().fg(Color::White).bg(cell_bg);
    let cell_cursor   = Style::default().fg(Color::Black).bg(cursor_bg).add_modifier(Modifier::BOLD);
    let cell_number   = Style::default().fg(number_color).bg(cell_bg);
    let cell_selected = Style::default().fg(Color::Black).bg(sel_bg);
    let gutter_style  = Style::default().fg(header_fg).bg(header_bg);
    let gutter_cur    = Style::default().fg(Color::Black).bg(cursor_bg).add_modifier(Modifier::BOLD);

    // Column widths: gutter(4) + A(12) + B(10) + C(9)
    let rows: &[(u16, &str, &[(Style, &str, bool)])] = &[
        // (gutter_style_idx, gutter_label, [(cell_style, text, right_align)])
        (0, "    ", &[
            (header_style,  "  A           ", false),
            (cursor_hdr,    "  B       ", false),
            (header_style,  "  C      ", false),
        ]),
        (1, "  1 ", &[
            (cell_normal,   " Revenue      ", false),
            (cell_selected, " 84,200   ", true),
            (cell_number,   "  12.5%  ", true),
        ]),
        (0, "  2 ", &[
            (cell_cursor,   " Q4 Budget    ", false),
            (cell_number,   " -31,450  ", true),
            (cell_number,   "   8.3%  ", true),
        ]),
        (0, "  3 ", &[
            (cell_normal,   " Net Profit   ", false),
            (cell_number,   "  52,750  ", true),
            (cell_number,   "  20.8%  ", true),
        ]),
        (0, "  4 ", &[
            (cell_normal,   "              ", false),
            (cell_normal,   "          ", false),
            (cell_normal,   "         ", false),
        ]),
    ];

    let gutter_styles = [gutter_style, gutter_cur];

    for (row_idx, (gs_idx, gutter, cells)) in rows.iter().enumerate() {
        let y = area.y + row_idx as u16;
        if y >= area.y + area.height { break; }

        // Gutter
        let gs = gutter_styles[*gs_idx as usize];
        let gw = 4u16.min(area.width);
        frame.render_widget(
            Paragraph::new(Line::from(Span::styled(*gutter, gs))),
            Rect { x: area.x, y, width: gw, height: 1 },
        );

        let mut x = area.x + gw;
        for (style, text, _right) in cells.iter() {
            let w = text.len() as u16;
            if x + w > area.x + area.width { break; }
            frame.render_widget(
                Paragraph::new(Line::from(Span::styled(*text, *style))),
                Rect { x, y, width: w, height: 1 },
            );
            x += w;
        }
    }

    // Status-bar mock at bottom of grid
    let sy = area.y + rows.len() as u16;
    if sy < area.y + area.height {
        let status_bg = Style::default().fg(header_fg).bg(header_bg);
        let normal_badge = Style::default().fg(Color::Black).bg(normal_color).add_modifier(Modifier::BOLD);
        frame.render_widget(
            Paragraph::new(Line::from(vec![
                Span::styled(" NORMAL ", normal_badge),
                Span::styled(" file.csv · B2/Sheet1", status_bg),
            ])).style(status_bg),
            Rect { x: area.x, y: sy, width: area.width, height: 1 },
        );
    }
}
