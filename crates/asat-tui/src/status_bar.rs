use asat_input::Mode;
use ratatui::{
    layout::Rect,
    style::{Color, Modifier, Style}, // Color needed for search mode teal
    text::{Line, Span},
    widgets::Paragraph,
    Frame,
};
use crate::{parse_hex_color, RenderState};
use asat_core::cell_address;

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let mode = &state.input.mode;
    let cursor = state.input.cursor;
    let workbook = state.workbook;
    let sheet = workbook.active();
    let theme = &state.config.theme;

    let normal_color  = parse_hex_color(&theme.normal_mode_color);
    let insert_color  = parse_hex_color(&theme.insert_mode_color);
    let visual_color  = parse_hex_color(&theme.visual_mode_color);
    let command_color = parse_hex_color(&theme.command_mode_color);
    let header_bg     = parse_hex_color(&theme.header_bg);
    let header_fg     = parse_hex_color(&theme.header_fg);

    // Mode indicator color
    let (mode_str, mode_color) = match mode {
        Mode::Normal                    => ("NORMAL",   normal_color),
        Mode::Insert { replace: false } => ("INSERT",   insert_color),
        Mode::Insert { replace: true  } => ("REPLACE",  command_color),
        Mode::Visual { block: false }   => ("VISUAL",   visual_color),
        Mode::Visual { block: true  }   => ("V-COL",    visual_color),
        Mode::VisualLine                => ("V-ROW",    visual_color),
        Mode::Command                   => ("COMMAND",  command_color),
        Mode::Search { forward: true  } => ("SEARCH↓",  Color::Rgb(42, 161, 152)),
        Mode::Search { forward: false } => ("SEARCH↑",  Color::Rgb(42, 161, 152)),
        Mode::Recording { .. }          => ("REC",      Color::Red),
        Mode::Welcome                   => ("WELCOME",  normal_color),
        Mode::FileFind                  => ("FIND",     insert_color),
        Mode::RecentFiles               => ("RECENT",   insert_color),
        Mode::ThemeManager                          => ("THEMES",   visual_color),
        Mode::FormulaSelect { anchor: None }        => ("F-REF",    insert_color),
        Mode::FormulaSelect { anchor: Some(_) }     => ("F-RANGE",  normal_color),
    };

    let mode_style = Style::default()
        .fg(Color::Black)
        .bg(mode_color)
        .add_modifier(Modifier::BOLD);
    let bg_style = Style::default()
        .fg(header_fg)
        .bg(header_bg);
    let right_style = Style::default()
        .fg(Color::White)
        .bg(header_bg);

    let file_name = workbook.file_name().unwrap_or("[No Name]");
    let dirty_marker = if workbook.dirty { " ●" } else { "" };
    let file_info = format!(" {}{} ", file_name, dirty_marker);

    let pos_info = format!(
        " {}  {} ",
        cell_address(cursor.row, cursor.col),
        &sheet.name,
    );

    // Status bar: [mode] [file_info] [spacer fills] [pos_info]
    let line = Line::from(vec![
        Span::styled(format!(" {} ", mode_str), mode_style),
        Span::styled(" ", bg_style),
        Span::styled(file_info, bg_style),
        Span::styled("", bg_style),   // flex spacer
        Span::styled(pos_info, right_style),
    ]);

    frame.render_widget(
        Paragraph::new(line).style(bg_style),
        area,
    );
}
