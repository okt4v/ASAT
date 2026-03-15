use crate::{parse_hex_color, RenderState};
use asat_input::Mode;
use ratatui::{
    layout::Rect,
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::Paragraph,
    Frame,
};

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let theme = &state.config.theme;
    let header_bg = parse_hex_color(&theme.header_bg);
    let amber = parse_hex_color(&theme.cursor_bg);

    let (prefix, content, is_command) = match &state.input.mode {
        Mode::Command => (":", state.input.command_buffer.as_str(), true),
        Mode::Search { forward: true } => ("/", state.input.search_buffer.as_str(), false),
        Mode::Search { forward: false } => ("?", state.input.search_buffer.as_str(), false),
        _ => return,
    };

    let bg_style = Style::default().fg(Color::White).bg(header_bg);

    // Prefix (':' '/' '?') is amber/teal depending on mode
    let prefix_color = if is_command {
        amber
    } else {
        Color::Rgb(42, 161, 152) // teal for search
    };
    let prefix_style = Style::default()
        .fg(prefix_color)
        .bg(header_bg)
        .add_modifier(Modifier::BOLD);

    // If a completion is active, tint the buffer text to hint it was auto-filled
    let content_style = if is_command && state.input.completion_idx.is_some() {
        Style::default()
            .fg(amber)
            .bg(header_bg)
            .add_modifier(Modifier::ITALIC)
    } else {
        Style::default().fg(Color::White).bg(header_bg)
    };

    let mut spans = vec![
        Span::styled(prefix, prefix_style),
        Span::styled(content, content_style),
        Span::styled("█", Style::default().fg(Color::White).bg(header_bg)),
    ];

    // Show a small hint when completions are available but not yet cycling
    if is_command && state.input.completion_idx.is_none() && !content.is_empty() {
        use asat_input::get_command_completions;
        let first_word = content.split_whitespace().next().unwrap_or("");
        let n = get_command_completions(first_word).len();
        if n > 0 {
            spans.push(Span::styled(
                format!("  [tab: {} match{}]", n, if n == 1 { "" } else { "es" }),
                Style::default().fg(Color::Rgb(50, 80, 110)).bg(header_bg),
            ));
        }
    }

    frame.render_widget(Paragraph::new(Line::from(spans)).style(bg_style), area);
}
