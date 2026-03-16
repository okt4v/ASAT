use crate::{darken, parse_hex_color, RenderState};
use asat_core::{col_to_letter, CellStyle, CellValue};
use asat_input::{Mode, VisualAnchor};
use ratatui::{
    buffer::Buffer,
    layout::Rect,
    style::{Color, Modifier, Style},
    widgets::Widget,
    Frame,
};
use unicode_width::UnicodeWidthStr;

/// Width of the row number gutter
const ROW_GUTTER_WIDTH: u16 = 5;
/// Minimum column width
const MIN_COL_WIDTH: u16 = 3;

pub fn render(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let widget = GridWidget { state };
    frame.render_widget(widget, area);
}

struct GridWidget<'a> {
    state: &'a RenderState<'a>,
}

impl<'a> Widget for GridWidget<'a> {
    fn render(self, area: Rect, buf: &mut Buffer) {
        if area.width < 4 || area.height < 2 {
            return;
        }

        let workbook = self.state.workbook;
        let input = self.state.input;
        let sheet = workbook.active();
        let viewport = input.viewport;
        let cursor = input.cursor;
        let theme = &self.state.config.theme;

        // Resolve theme colors
        let cursor_bg = parse_hex_color(&theme.cursor_bg);
        let header_bg = parse_hex_color(&theme.header_bg);
        let header_fg = parse_hex_color(&theme.header_fg);
        let cell_bg = parse_hex_color(&theme.cell_bg);
        let selection_bg = parse_hex_color(&theme.selection_bg);
        let number_color = parse_hex_color(&theme.number_color);
        let cmd_color = parse_hex_color(&theme.command_mode_color);
        let vis_color = parse_hex_color(&theme.visual_mode_color);
        let insert_color = parse_hex_color(&theme.insert_mode_color);
        let normal_color = parse_hex_color(&theme.normal_mode_color);

        // ── Freeze pane metrics ──────────────────────────────────────────────
        let freeze_rows = sheet.freeze_rows as u16;
        let freeze_cols = sheet.freeze_cols;

        // Frozen column widths (always rendered at fixed left position)
        let available_width = area.width.saturating_sub(ROW_GUTTER_WIDTH);
        let mut frozen_cols: Vec<(u32, u16)> = Vec::new();
        let mut frozen_col_width: u16 = 0;
        for fc in 0..freeze_cols {
            let w = sheet.col_width(fc).max(MIN_COL_WIDTH);
            let w = w.min(available_width.saturating_sub(frozen_col_width));
            frozen_cols.push((fc, w));
            frozen_col_width += w;
            if frozen_col_width >= available_width {
                break;
            }
        }
        // 1-char separator between frozen and scrollable cols (only if freeze_cols > 0)
        let freeze_col_sep = if freeze_cols > 0 && frozen_col_width < available_width {
            1u16
        } else {
            0u16
        };
        let scroll_col_available = available_width
            .saturating_sub(frozen_col_width)
            .saturating_sub(freeze_col_sep);

        // Layout: top row = column headers, then freeze_rows frozen rows, then data rows
        let header_y = area.y;
        // 1-char separator row between frozen rows and scrollable rows
        let freeze_row_sep: u16 = if freeze_rows > 0 { 1 } else { 0 };
        let data_y = area.y + 1 + freeze_rows + freeze_row_sep;
        let data_height = area.height.saturating_sub(1 + freeze_rows + freeze_row_sep);

        // Scrollable column list (starts at viewport.left_col, must be >= freeze_cols)
        let scroll_left = viewport.left_col.max(freeze_cols);
        let mut scroll_cols: Vec<(u32, u16)> = Vec::new();
        {
            let mut x_offset = 0u16;
            let mut c = scroll_left;
            loop {
                if x_offset >= scroll_col_available {
                    break;
                }
                let col_width = sheet.col_width(c).max(MIN_COL_WIDTH);
                scroll_cols.push((c, col_width.min(scroll_col_available - x_offset)));
                x_offset += col_width.min(scroll_col_available - x_offset);
                if x_offset >= scroll_col_available {
                    break;
                }
                c += 1;
                if c > 10000 {
                    break;
                }
            }
        }

        // Combined column list for header rendering (frozen + separator + scroll)
        let corner_style = Style::default()
            .fg(header_fg)
            .bg(header_bg)
            .add_modifier(Modifier::BOLD);

        // ── Header row ───────────────────────────────────────────────────────
        render_cell_str(
            buf,
            area.x,
            header_y,
            ROW_GUTTER_WIDTH,
            "     ",
            corner_style,
        );

        let mut x = area.x + ROW_GUTTER_WIDTH;
        // Frozen col headers
        for (col_idx, col_width) in &frozen_cols {
            let label = col_to_letter(*col_idx);
            let is_cursor_col = *col_idx == cursor.col;
            let hdr_style = if is_cursor_col {
                Style::default()
                    .fg(Color::Black)
                    .bg(cursor_bg)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default()
                    .fg(header_fg)
                    .bg(darken(header_bg, 0.7))
                    .add_modifier(Modifier::BOLD)
            };
            render_cell_centered(buf, x, header_y, *col_width, &label, hdr_style);
            x += col_width;
        }
        // Separator column header
        if freeze_col_sep > 0 {
            let sep_style = Style::default().fg(header_fg).bg(darken(header_bg, 0.5));
            render_cell_str(buf, x, header_y, 1, "┃", sep_style);
            x += 1;
        }
        // Scrollable col headers
        for (col_idx, col_width) in &scroll_cols {
            let label = col_to_letter(*col_idx);
            let is_cursor_col = *col_idx == cursor.col;
            let hdr_style = if is_cursor_col {
                Style::default()
                    .fg(Color::Black)
                    .bg(cursor_bg)
                    .add_modifier(Modifier::BOLD)
            } else {
                corner_style
            };
            render_cell_centered(buf, x, header_y, *col_width, &label, hdr_style);
            x += col_width;
        }

        // Build combined column list used by both frozen rows and scrollable rows
        let mut all_col_groups: Vec<(u32, u16)> = frozen_cols.clone();
        if freeze_col_sep > 0 {
            all_col_groups.push((u32::MAX, freeze_col_sep));
        }
        all_col_groups.extend(scroll_cols.iter().copied());

        let note_marker_style = Style::default().fg(Color::Rgb(255, 200, 50));

        // ── Frozen rows ───────────────────────────────────────────────────────
        for frozen_r in 0..freeze_rows {
            let row_idx = frozen_r as u32;
            let screen_y = area.y + 1 + frozen_r;
            let is_cursor_row = row_idx == cursor.row;
            let gutter_style = if is_cursor_row {
                Style::default()
                    .fg(Color::Black)
                    .bg(cursor_bg)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(header_fg).bg(darken(header_bg, 0.7))
            };
            let gutter_text = format!("{:>4}┃", row_idx + 1);
            render_cell_str(
                buf,
                area.x,
                screen_y,
                ROW_GUTTER_WIDTH,
                &gutter_text,
                gutter_style,
            );

            let mut x = area.x + ROW_GUTTER_WIDTH;
            for &(col_idx, col_width) in &all_col_groups {
                if col_idx == u32::MAX {
                    let sep_style = Style::default()
                        .fg(darken(header_fg, 0.6))
                        .bg(darken(header_bg, 0.5));
                    render_cell_str(buf, x, screen_y, col_width, "┃", sep_style);
                    x += col_width;
                    continue;
                }
                if let Some(m) = sheet.merge_at(row_idx, col_idx) {
                    if row_idx == m.row_start && col_idx == m.col_start {
                        // Anchor cell — render with full merged width
                        let actual_width = all_col_groups
                            .iter()
                            .filter(|(c, _)| *c != u32::MAX && *c >= m.col_start && *c <= m.col_end)
                            .map(|(_, w)| *w)
                            .sum::<u16>()
                            .max(col_width);
                        render_data_cell(
                            buf, x, screen_y, actual_width,
                            row_idx, col_idx, cursor, &input.mode,
                            input.visual_anchor.as_ref(),
                            input.search_highlight(row_idx, col_idx),
                            sheet, &input.edit_buffer, input.formula_origin,
                            cursor_bg, cell_bg, selection_bg,
                            number_color, cmd_color, vis_color, insert_color, normal_color,
                        );
                        if sheet.notes.contains_key(&(row_idx, col_idx)) {
                            if let Some(cell) = buf.cell_mut((x + col_width.saturating_sub(1), screen_y)) {
                                cell.set_char('▸');
                                cell.set_style(note_marker_style);
                            }
                        }
                    } else if row_idx > m.row_start {
                        // Covered cell in a row below anchor — paint matching background
                        let is_anchor_cursor = m.row_start == cursor.row && m.col_start == cursor.col;
                        let bg = if is_anchor_cursor {
                            cursor_bg
                        } else {
                            sheet.get_cell(m.row_start, m.col_start)
                                .and_then(|c| c.style.as_ref())
                                .and_then(|s| s.bg.as_ref())
                                .map(|c| Color::Rgb(c.r, c.g, c.b))
                                .unwrap_or(cell_bg)
                        };
                        render_cell_str(buf, x, screen_y, col_width, "", Style::default().bg(bg));
                    }
                    // row_idx == m.row_start but col != m.col_start: same-row covered cell,
                    // anchor's wide render already painted this area — just advance x.
                } else {
                    // Normal (non-merged) cell
                    render_data_cell(
                        buf, x, screen_y, col_width,
                        row_idx, col_idx, cursor, &input.mode,
                        input.visual_anchor.as_ref(),
                        input.search_highlight(row_idx, col_idx),
                        sheet, &input.edit_buffer, input.formula_origin,
                        cursor_bg, cell_bg, selection_bg,
                        number_color, cmd_color, vis_color, insert_color, normal_color,
                    );
                    if sheet.notes.contains_key(&(row_idx, col_idx)) {
                        if let Some(cell) = buf.cell_mut((x + col_width.saturating_sub(1), screen_y)) {
                            cell.set_char('▸');
                            cell.set_style(note_marker_style);
                        }
                    }
                }
                x += col_width;
            }
        }

        // ── Frozen row separator ─────────────────────────────────────────────
        if freeze_rows > 0 && freeze_row_sep > 0 {
            let sep_y = area.y + 1 + freeze_rows;
            let sep_style = Style::default()
                .fg(darken(header_fg, 0.6))
                .bg(darken(header_bg, 0.5));
            for dx in 0..area.width {
                if let Some(cell) = buf.cell_mut((area.x + dx, sep_y)) {
                    cell.set_char('━');
                    cell.set_style(sep_style);
                }
            }
        }

        // ── Data rows ────────────────────────────────────────────────────────
        let mut screen_y = data_y;
        let mut row_idx = viewport.top_row.max(freeze_rows as u32);
        while screen_y < data_y + data_height {
            // Skip hidden rows (filter)
            if sheet
                .row_meta
                .get(&row_idx)
                .map(|m| m.hidden)
                .unwrap_or(false)
            {
                row_idx += 1;
                continue;
            }
            // Skip rows already shown in frozen pane
            if row_idx < freeze_rows as u32 {
                row_idx = freeze_rows as u32;
                continue;
            }

            let row_h = {
                let stored = sheet.row_height(row_idx).max(1);
                // Auto-expand height for any wrapped cell in this row
                let wrap_h = all_col_groups
                    .iter()
                    .filter_map(|(c, w)| {
                        if *c == u32::MAX {
                            return None;
                        }
                        // Covered cells are handled by their anchor
                        if sheet.is_covered(row_idx, *c) {
                            return None;
                        }
                        let has_wrap = sheet
                            .get_cell(row_idx, *c)
                            .and_then(|cell| cell.style.as_ref())
                            .map(|s| s.wrap)
                            .unwrap_or(false);
                        if !has_wrap {
                            return None;
                        }
                        // Use merged width if this cell is a merge anchor
                        let cw = if let Some(m) = sheet.merge_at(row_idx, *c) {
                            all_col_groups
                                .iter()
                                .filter(|(c2, _)| {
                                    *c2 != u32::MAX && *c2 >= m.col_start && *c2 <= m.col_end
                                })
                                .map(|(_, mw)| *mw)
                                .sum::<u16>()
                                .max(1) as usize
                        } else {
                            (*w).max(1) as usize
                        };
                        let text_len = sheet.display_value(row_idx, *c).chars().count();
                        Some(((text_len + cw - 1) / cw).max(1) as u16)
                    })
                    .max()
                    .unwrap_or(1);
                stored.max(wrap_h)
            };
            let is_cursor_row = row_idx == cursor.row;

            // ── First line: gutter + cell content ────────────────────────────
            let gutter_label = format!("{:>4}", row_idx + 1);
            // Show height indicator when row is taller than 1
            let gutter_text = if row_h > 1 {
                format!("{:>3}↕ ", row_idx + 1)
            } else {
                format!("{} ", gutter_label)
            };
            let gutter_style = if is_cursor_row {
                Style::default()
                    .fg(Color::Black)
                    .bg(cursor_bg)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(header_fg).bg(header_bg)
            };
            render_cell_str(
                buf,
                area.x,
                screen_y,
                ROW_GUTTER_WIDTH,
                &gutter_text,
                gutter_style,
            );

            let mut x = area.x + ROW_GUTTER_WIDTH;
            for (col_idx, col_width) in &all_col_groups {
                if *col_idx == u32::MAX {
                    // Freeze separator column
                    let sep_style = Style::default()
                        .fg(darken(header_fg, 0.6))
                        .bg(darken(header_bg, 0.5));
                    render_cell_str(buf, x, screen_y, *col_width, "┃", sep_style);
                    x += col_width;
                    continue;
                }

                if let Some(m) = sheet.merge_at(row_idx, *col_idx) {
                    if row_idx == m.row_start && *col_idx == m.col_start {
                        // Anchor cell — render with full merged width
                        let actual_width = all_col_groups
                            .iter()
                            .filter(|(c, _)| *c != u32::MAX && *c >= m.col_start && *c <= m.col_end)
                            .map(|(_, w)| *w)
                            .sum::<u16>()
                            .max(*col_width);
                        render_data_cell(
                            buf, x, screen_y, actual_width,
                            row_idx, *col_idx, cursor, &input.mode,
                            input.visual_anchor.as_ref(),
                            input.search_highlight(row_idx, *col_idx),
                            sheet, &input.edit_buffer, input.formula_origin,
                            cursor_bg, cell_bg, selection_bg,
                            number_color, cmd_color, vis_color, insert_color, normal_color,
                        );
                        if sheet.notes.contains_key(&(row_idx, *col_idx)) {
                            if let Some(cell) = buf.cell_mut((x + col_width.saturating_sub(1), screen_y)) {
                                cell.set_char('▸');
                                cell.set_style(note_marker_style);
                            }
                        }
                    } else if row_idx > m.row_start {
                        // Covered cell in a row below anchor — paint matching background
                        let is_anchor_cursor = m.row_start == cursor.row && m.col_start == cursor.col;
                        let bg = if is_anchor_cursor {
                            cursor_bg
                        } else {
                            sheet.get_cell(m.row_start, m.col_start)
                                .and_then(|c| c.style.as_ref())
                                .and_then(|s| s.bg.as_ref())
                                .map(|c| Color::Rgb(c.r, c.g, c.b))
                                .unwrap_or(cell_bg)
                        };
                        render_cell_str(buf, x, screen_y, *col_width, "", Style::default().bg(bg));
                    }
                    // row == m.row_start, col != m.col_start: same-row covered, anchor painted it.
                } else {
                    // Normal (non-merged) cell
                    render_data_cell(
                        buf, x, screen_y, *col_width,
                        row_idx, *col_idx, cursor, &input.mode,
                        input.visual_anchor.as_ref(),
                        input.search_highlight(row_idx, *col_idx),
                        sheet, &input.edit_buffer, input.formula_origin,
                        cursor_bg, cell_bg, selection_bg,
                        number_color, cmd_color, vis_color, insert_color, normal_color,
                    );
                    if sheet.notes.contains_key(&(row_idx, *col_idx)) {
                        if let Some(cell) = buf.cell_mut((x + col_width.saturating_sub(1), screen_y)) {
                            cell.set_char('▸');
                            cell.set_style(note_marker_style);
                        }
                    }
                }
                x += col_width;
            }

            // ── Extra lines for tall rows (background fill) ───────────────────
            for extra in 1..row_h {
                let ey = screen_y + extra;
                if ey >= data_y + data_height {
                    break;
                }

                // Gutter continuation marker
                let cont_style = if is_cursor_row {
                    gutter_style
                } else {
                    Style::default().fg(darken(header_fg, 0.6)).bg(header_bg)
                };
                render_cell_str(buf, area.x, ey, ROW_GUTTER_WIDTH, "    │", cont_style);

                // Fill cells with background; render wrapped text for wrap-enabled cells
                let mut ex = area.x + ROW_GUTTER_WIDTH;
                for (col_idx, col_width) in &all_col_groups {
                    if *col_idx == u32::MAX {
                        let sep_style = Style::default()
                            .fg(darken(header_fg, 0.6))
                            .bg(darken(header_bg, 0.5));
                        render_cell_str(buf, ex, ey, *col_width, "", sep_style);
                        ex += col_width;
                        continue;
                    }
                    // Skip covered cells (same-row: anchor's wide render covers them)
                    if sheet.is_covered(row_idx, *col_idx) {
                        ex += col_width;
                        continue;
                    }
                    // For merged anchors, use the full merged width
                    let render_width = if let Some(m) = sheet.merge_at(row_idx, *col_idx) {
                        all_col_groups
                            .iter()
                            .filter(|(c, _)| *c != u32::MAX && *c >= m.col_start && *c <= m.col_end)
                            .map(|(_, w)| *w)
                            .sum::<u16>()
                            .max(*col_width)
                    } else {
                        *col_width
                    };
                    let wrap_on = sheet
                        .get_cell(row_idx, *col_idx)
                        .and_then(|c| c.style.as_ref())
                        .map(|s| s.wrap)
                        .unwrap_or(false);
                    if wrap_on {
                        let display = sheet.display_value(row_idx, *col_idx);
                        let chunk_size = render_width.max(1) as usize;
                        let chunk: String = display
                            .chars()
                            .skip(chunk_size * extra as usize)
                            .take(chunk_size)
                            .collect();
                        let is_cursor_cell = is_cursor_row && *col_idx == cursor.col;
                        let style = if is_cursor_cell {
                            Style::default().fg(Color::Black).bg(cursor_bg)
                        } else {
                            let user_bg = sheet
                                .get_cell(row_idx, *col_idx)
                                .and_then(|c| c.style.as_ref())
                                .and_then(|s| s.bg.as_ref())
                                .map(|c| Color::Rgb(c.r, c.g, c.b))
                                .unwrap_or(cell_bg);
                            Style::default().fg(Color::White).bg(user_bg)
                        };
                        render_cell_str(buf, ex, ey, render_width, &chunk, style);
                    } else {
                        let bg_style = if is_cursor_row {
                            Style::default().fg(Color::Black).bg(cell_bg)
                        } else {
                            Style::default().fg(Color::White).bg(cell_bg)
                        };
                        render_cell_str(buf, ex, ey, *col_width, "", bg_style);
                    }
                    ex += col_width;
                }
            }

            screen_y += row_h;
            row_idx += 1;
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn render_data_cell(
    buf: &mut Buffer,
    x: u16,
    y: u16,
    col_width: u16,
    row_idx: u32,
    col_idx: u32,
    cursor: asat_input::Cursor,
    mode: &Mode,
    visual_anchor: Option<&asat_input::VisualAnchor>,
    search_hl: Option<bool>,
    sheet: &asat_core::Sheet,
    edit_buffer: &str,
    formula_origin: Option<(u32, u32)>,
    cursor_bg: Color,
    cell_bg: Color,
    selection_bg: Color,
    number_color: Color,
    cmd_color: Color,
    vis_color: Color,
    insert_color: Color,
    normal_color: Color,
) {
    let is_cursor = row_idx == cursor.row && col_idx == cursor.col;

    let is_fref_cursor = matches!(mode, Mode::FormulaSelect { .. })
        && row_idx == cursor.row
        && col_idx == cursor.col;
    let is_fref_range = if let Mode::FormulaSelect {
        anchor: Some((ar, ac)),
    } = mode
    {
        let r_min = (*ar).min(cursor.row);
        let r_max = (*ar).max(cursor.row);
        let c_min = (*ac).min(cursor.col);
        let c_max = (*ac).max(cursor.col);
        row_idx >= r_min && row_idx <= r_max && col_idx >= c_min && col_idx <= c_max
    } else {
        false
    };

    let is_visual = is_in_visual_selection(row_idx, col_idx, visual_anchor, cursor, mode);

    // In FormulaSelect, show edit_buffer on the cell being edited (formula_origin),
    // not on the navigator cursor which has moved away.
    let is_formula_origin =
        matches!(mode, Mode::FormulaSelect { .. }) && formula_origin == Some((row_idx, col_idx));
    let live_edit = (is_cursor && matches!(mode, Mode::Insert { .. })) || is_formula_origin;
    let wrap_on = sheet
        .get_cell(row_idx, col_idx)
        .and_then(|c| c.style.as_ref())
        .map(|s| s.wrap)
        .unwrap_or(false);
    let display = if live_edit {
        edit_buffer.to_string()
    } else if wrap_on {
        // First line of a wrapped cell: take exactly col_width chars so render_cell_str
        // doesn't see it as overflowing and doesn't add an ellipsis.
        sheet
            .display_value(row_idx, col_idx)
            .chars()
            .take(col_width as usize)
            .collect()
    } else {
        sheet.display_value(row_idx, col_idx)
    };
    let raw_value = sheet.get_value(row_idx, col_idx);

    let cell_style = if is_fref_cursor {
        Style::default()
            .fg(Color::Black)
            .bg(normal_color)
            .add_modifier(Modifier::BOLD)
    } else if is_fref_range {
        Style::default()
            .fg(Color::Black)
            .bg(darken(insert_color, 0.55))
    } else if is_cursor {
        Style::default()
            .fg(Color::Black)
            .bg(cursor_bg)
            .add_modifier(Modifier::BOLD)
    } else if is_visual {
        Style::default().fg(Color::Black).bg(selection_bg)
    } else if search_hl == Some(true) {
        Style::default()
            .fg(Color::Black)
            .bg(cmd_color)
            .add_modifier(Modifier::BOLD)
    } else if search_hl == Some(false) {
        Style::default()
            .fg(Color::Black)
            .bg(darken(vis_color, 0.75))
    } else {
        let user_style: Option<&CellStyle> = sheet
            .get_cell(row_idx, col_idx)
            .and_then(|c| c.style.as_ref());
        let bg = user_style
            .and_then(|s| s.bg.as_ref())
            .map(|c| Color::Rgb(c.r, c.g, c.b))
            .unwrap_or(cell_bg);
        let default_fg = if matches!(raw_value, CellValue::Error(_)) {
            cmd_color
        } else if matches!(raw_value, CellValue::Number(_) | CellValue::Boolean(_)) {
            number_color
        } else {
            Color::White
        };
        let fg = user_style
            .and_then(|s| s.fg.as_ref())
            .map(|c| Color::Rgb(c.r, c.g, c.b))
            .unwrap_or(default_fg);
        let mut s = Style::default().fg(fg).bg(bg);
        if let Some(us) = user_style {
            if us.bold {
                s = s.add_modifier(Modifier::BOLD);
            }
            if us.italic {
                s = s.add_modifier(Modifier::ITALIC);
            }
            if us.underline {
                s = s.add_modifier(Modifier::UNDERLINED);
            }
            if us.strikethrough {
                s = s.add_modifier(Modifier::CROSSED_OUT);
            }
        }
        s
    };

    use asat_core::Alignment;
    let user_align = sheet
        .get_cell(row_idx, col_idx)
        .and_then(|c| c.style.as_ref())
        .map(|s| s.align);
    match user_align {
        Some(Alignment::Right) => render_cell_right(buf, x, y, col_width, &display, cell_style),
        Some(Alignment::Center) => render_cell_centered(buf, x, y, col_width, &display, cell_style),
        Some(Alignment::Left) => render_cell_str(buf, x, y, col_width, &display, cell_style),
        _ => {
            if matches!(raw_value, CellValue::Number(_) | CellValue::Boolean(_)) {
                render_cell_right(buf, x, y, col_width, &display, cell_style);
            } else {
                render_cell_str(buf, x, y, col_width, &display, cell_style);
            }
        }
    }
}

fn is_in_visual_selection(
    row: u32,
    col: u32,
    anchor: Option<&VisualAnchor>,
    cursor: asat_input::Cursor,
    mode: &Mode,
) -> bool {
    let anchor = match anchor {
        Some(a) => a,
        None => return false,
    };
    let r_min = cursor.row.min(anchor.row);
    let r_max = cursor.row.max(anchor.row);
    match mode {
        Mode::VisualLine => row >= r_min && row <= r_max,
        Mode::Visual { block: true } => {
            // Column select: all rows, bounded by column range only
            let c_min = cursor.col.min(anchor.col);
            let c_max = cursor.col.max(anchor.col);
            col >= c_min && col <= c_max
        }
        Mode::Visual { block: false } => {
            let c_min = cursor.col.min(anchor.col);
            let c_max = cursor.col.max(anchor.col);
            row >= r_min && row <= r_max && col >= c_min && col <= c_max
        }
        _ => false,
    }
}

/// Render a fixed-width cell left-aligned. Appends "…" (single char) if content is wider.
fn render_cell_str(buf: &mut Buffer, x: u16, y: u16, width: u16, content: &str, style: Style) {
    if width == 0 {
        return;
    }
    // Fill background
    for dx in 0..width {
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(' ');
            cell.set_style(style);
        }
    }
    let content_w = UnicodeWidthStr::width(content) as u16;
    let truncated = content_w > width;
    // Reserve 1 char for the ellipsis indicator when truncating
    let text_budget = if truncated && width > 1 {
        width - 1
    } else {
        width
    };

    let mut dx = 0u16;
    for ch in content.chars() {
        if dx >= text_budget {
            break;
        }
        let cw = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(1) as u16;
        if dx + cw > text_budget {
            break;
        }
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(ch);
            cell.set_style(style);
        }
        dx += cw;
    }
    // Ellipsis — dim version of the cell's own style so it reads as "not content"
    if truncated && width > 1 {
        let ellipsis_style = style.add_modifier(Modifier::DIM);
        if let Some(cell) = buf.cell_mut((x + width - 1, y)) {
            cell.set_char('…');
            cell.set_style(ellipsis_style);
        }
    }
}

/// Right-aligned cell render. Shows "####" when a number won't fit (Excel convention).
fn render_cell_right(buf: &mut Buffer, x: u16, y: u16, width: u16, content: &str, style: Style) {
    if width == 0 {
        return;
    }
    for dx in 0..width {
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(' ');
            cell.set_style(style);
        }
    }
    let display_width = UnicodeWidthStr::width(content) as u16;
    if display_width > width {
        // Too wide — show #### like Excel (reversed style = always readable)
        let hash_style = style
            .add_modifier(Modifier::REVERSED)
            .add_modifier(Modifier::BOLD);
        for dx in 0..width {
            if let Some(cell) = buf.cell_mut((x + dx, y)) {
                cell.set_char('#');
                cell.set_style(hash_style);
            }
        }
        return;
    }
    let start_dx = width - display_width;
    let mut dx = start_dx;
    for ch in content.chars() {
        if dx >= width {
            break;
        }
        let cw = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(1) as u16;
        if dx + cw > width {
            break;
        }
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(ch);
            cell.set_style(style);
        }
        dx += cw;
    }
}

/// Center-aligned cell render (for column headers)
fn render_cell_centered(buf: &mut Buffer, x: u16, y: u16, width: u16, content: &str, style: Style) {
    if width == 0 {
        return;
    }
    for dx in 0..width {
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(' ');
            cell.set_style(style);
        }
    }
    let display_width = UnicodeWidthStr::width(content) as u16;
    let start_dx = if display_width < width {
        (width - display_width) / 2
    } else {
        0
    };
    let mut dx = start_dx;
    for ch in content.chars() {
        if dx >= width {
            break;
        }
        let ch_width = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(1) as u16;
        if dx + ch_width > width {
            break;
        }
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(ch);
            cell.set_style(style);
        }
        dx += ch_width;
    }
}
