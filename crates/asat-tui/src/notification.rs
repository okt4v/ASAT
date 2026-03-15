use crate::{is_dark_color, parse_hex_color, RenderState};
use ratatui::{
    layout::Rect,
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, BorderType, Borders, Clear, Paragraph},
    Frame,
};

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let msg = match state.status_message {
        Some(m) if !m.is_empty() => m,
        _ => return,
    };

    let theme = &state.config.theme;
    let accent = parse_hex_color(&theme.cursor_bg);
    let bg = parse_hex_color(&theme.header_bg);
    let text_fg = if is_dark_color(bg) {
        Color::White
    } else {
        Color::Black
    };

    // Prefix with an icon based on message content
    let icon = if msg.to_ascii_lowercase().contains("error")
        || msg.to_ascii_lowercase().contains("unknown")
        || msg.to_ascii_lowercase().contains("failed")
    {
        "✗ "
    } else if msg.to_ascii_lowercase().contains("saved")
        || msg.to_ascii_lowercase().contains("written")
        || msg.to_ascii_lowercase().contains("applied")
        || msg.to_ascii_lowercase().contains("copied")
        || msg.to_ascii_lowercase().contains("yanked")
    {
        "✓ "
    } else {
        "● "
    };

    // Max 50 visible chars; truncate with ellipsis
    let max_chars = 50usize;
    let full = format!("{}{}", icon, msg);
    let display: String = if full.chars().count() > max_chars {
        format!("{}…", full.chars().take(max_chars - 1).collect::<String>())
    } else {
        full
    };

    // Popup sizing: content + 1 space padding each side + 2 borders
    let content_w = display.chars().count() as u16;
    let popup_w = (content_w + 4).min(area.width.saturating_sub(2));
    let popup_h = 3u16;

    // Position: top-right, 1 row and 1 col from the edge
    let x = area.x + area.width.saturating_sub(popup_w + 1);
    let y = area.y + 1;

    if popup_w < 6 || area.height < 4 {
        return;
    }

    let rect = Rect {
        x,
        y,
        width: popup_w,
        height: popup_h,
    };

    let line = Line::from(Span::styled(
        format!(" {} ", display),
        Style::default().fg(text_fg).bg(bg),
    ));

    let block = Block::default()
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(accent).add_modifier(Modifier::BOLD))
        .style(Style::default().bg(bg));

    frame.render_widget(Clear, rect);
    frame.render_widget(Paragraph::new(line).block(block), rect);
}
