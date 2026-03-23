use crate::RenderState;
use ratatui::{
    layout::Rect,
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, BorderType, Borders, Clear, Paragraph, Wrap},
    Frame,
};

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let note = match &state.input.note_popup {
        Some(n) => n.as_str(),
        None => return,
    };

    let popup_w = (area.width * 2 / 3).clamp(30, 80);
    // Estimate height: count wrapped lines + 2 borders + 1 footer
    let inner_w = popup_w.saturating_sub(4) as usize; // 2 borders + 2 padding
    let wrapped_lines = wrap_text(note, inner_w);
    let content_h = wrapped_lines.len() as u16;
    let popup_h = (content_h + 4).min(area.height.saturating_sub(4)); // +2 border +1 title +1 footer

    let x = area.x + (area.width.saturating_sub(popup_w)) / 2;
    let y = area.y + (area.height.saturating_sub(popup_h)) / 2;

    if popup_w < 10 || popup_h < 4 {
        return;
    }

    let rect = Rect {
        x,
        y,
        width: popup_w,
        height: popup_h,
    };

    let amber = Color::Rgb(255, 200, 50);
    let bg = Color::Rgb(18, 20, 35);
    let fg = Color::Rgb(230, 235, 255);
    let dim = Color::Rgb(120, 125, 160);

    let mut lines: Vec<Line> = wrapped_lines
        .iter()
        .map(|l| {
            Line::from(Span::styled(
                format!(" {} ", l),
                Style::default().fg(fg).bg(bg),
            ))
        })
        .collect();

    // Footer hint
    lines.push(Line::from(Span::styled(
        " press any key to close ",
        Style::default()
            .fg(dim)
            .bg(bg)
            .add_modifier(Modifier::ITALIC),
    )));

    let block = Block::default()
        .title(Span::styled(
            " ▸ Cell Note ",
            Style::default()
                .fg(amber)
                .bg(bg)
                .add_modifier(Modifier::BOLD),
        ))
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(amber))
        .style(Style::default().bg(bg));

    frame.render_widget(Clear, rect);
    frame.render_widget(
        Paragraph::new(lines)
            .block(block)
            .wrap(Wrap { trim: false }),
        rect,
    );
}

/// Simple word-wrap: split `text` into lines of at most `width` chars.
fn wrap_text(text: &str, width: usize) -> Vec<String> {
    if width == 0 {
        return vec![text.to_string()];
    }
    let mut lines = Vec::new();
    for paragraph in text.split('\n') {
        if paragraph.is_empty() {
            lines.push(String::new());
            continue;
        }
        let mut current = String::new();
        for word in paragraph.split_whitespace() {
            if current.is_empty() {
                current.push_str(word);
            } else if current.chars().count() + 1 + word.chars().count() <= width {
                current.push(' ');
                current.push_str(word);
            } else {
                lines.push(current.clone());
                current = word.to_string();
            }
        }
        if !current.is_empty() {
            lines.push(current);
        }
    }
    if lines.is_empty() {
        lines.push(String::new());
    }
    lines
}
