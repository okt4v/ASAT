use crate::{parse_hex_color, RenderState};
use ratatui::{
    layout::Rect,
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::Paragraph,
    Frame,
};

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let workbook = state.workbook;
    let active = workbook.active_sheet;
    let theme = &state.config.theme;

    let cursor_bg = parse_hex_color(&theme.cursor_bg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let header_fg = parse_hex_color(&theme.header_fg);

    let active_style = Style::default()
        .fg(Color::Black)
        .bg(cursor_bg)
        .add_modifier(Modifier::BOLD);
    let inactive_style = Style::default().fg(header_fg).bg(header_bg);

    let mut spans: Vec<Span> = Vec::new();
    for (i, sheet) in workbook.sheets.iter().enumerate() {
        let dirty_marker = if workbook.dirty && i == active {
            " ●"
        } else {
            ""
        };
        let label = format!(" {}{} ", &sheet.name, dirty_marker);
        if i == active {
            spans.push(Span::styled(label, active_style));
        } else {
            spans.push(Span::styled(label, inactive_style));
        }
        spans.push(Span::styled(" ", inactive_style));
    }
    spans.push(Span::styled(" + ", inactive_style));

    let line = Line::from(spans);
    frame.render_widget(Paragraph::new(line).style(inactive_style), area);
}
