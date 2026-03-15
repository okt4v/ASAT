use crate::{parse_hex_color, RenderState};
use asat_core::cell_address;
use asat_input::Mode;
use ratatui::{
    layout::Rect,
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::Paragraph,
    Frame,
};

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let cursor = state.input.cursor;
    let sheet = state.workbook.active();
    let theme = &state.config.theme;

    // In FormulaSelect mode, the address shows where we started editing (formula_origin),
    // not where the selection cursor is now.
    let addr = if let Some((or, oc)) = state.input.formula_origin {
        cell_address(or, oc)
    } else {
        cell_address(cursor.row, cursor.col)
    };

    let cursor_bg = parse_hex_color(&theme.cursor_bg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let normal_c = parse_hex_color(&theme.normal_mode_color);
    let insert_c = parse_hex_color(&theme.insert_mode_color);

    let addr_style = Style::default()
        .fg(Color::Black)
        .bg(cursor_bg)
        .add_modifier(Modifier::BOLD);
    let sep_style = Style::default().fg(header_fg).bg(header_bg);
    let content_style = Style::default().fg(Color::White).bg(header_bg);

    let spans = match &state.input.mode {
        Mode::Insert { .. } => {
            // Show the edit buffer live while typing
            vec![
                Span::styled(format!(" {:>6} ", addr), addr_style),
                Span::styled(" │ ", sep_style),
                Span::styled(state.input.edit_buffer.clone(), content_style),
            ]
        }

        Mode::FormulaSelect { anchor } => {
            // Show the formula being built + a live preview of the ref under the cursor
            let formula_so_far = state.input.edit_buffer.clone();

            let preview_ref = if let Some((ar, ac)) = anchor {
                // Range in progress
                format!(
                    "{}:{}",
                    cell_address(*ar, *ac),
                    cell_address(cursor.row, cursor.col)
                )
            } else {
                // Single cell
                cell_address(cursor.row, cursor.col)
            };

            let ref_style = Style::default()
                .fg(Color::Black)
                .bg(normal_c)
                .add_modifier(Modifier::BOLD);

            let hint_style = Style::default()
                .fg(insert_c)
                .bg(header_bg)
                .add_modifier(Modifier::DIM);

            vec![
                Span::styled(format!(" {:>6} ", addr), addr_style),
                Span::styled(" │ ", sep_style),
                Span::styled(formula_so_far, content_style),
                Span::styled(format!(" {}", preview_ref), ref_style),
                Span::styled(
                    if anchor.is_some() {
                        "  [Enter] insert range  [:] re-anchor  [Esc] cancel"
                    } else {
                        "  [Enter] insert ref  [:] start range  [Esc] cancel"
                    },
                    hint_style,
                ),
            ]
        }

        _ => {
            // Normal mode — show raw formula, not computed value
            let raw = sheet
                .get_raw_value(cursor.row, cursor.col)
                .formula_bar_display();
            vec![
                Span::styled(format!(" {:>6} ", addr), addr_style),
                Span::styled(" │ ", sep_style),
                Span::styled(raw, content_style),
            ]
        }
    };

    frame.render_widget(
        Paragraph::new(Line::from(spans)).style(Style::default().bg(header_bg)),
        area,
    );
}
