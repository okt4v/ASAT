use asat_core::{col_to_letter, CellValue, CellStyle};
use asat_input::{Mode, VisualAnchor};
use ratatui::{
    buffer::Buffer,
    layout::Rect,
    style::{Color, Modifier, Style},
    widgets::Widget,
    Frame,
};
use unicode_width::UnicodeWidthStr;
use crate::{darken, parse_hex_color, RenderState};

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
        let cursor_bg    = parse_hex_color(&theme.cursor_bg);
        let header_bg    = parse_hex_color(&theme.header_bg);
        let header_fg    = parse_hex_color(&theme.header_fg);
        let cell_bg      = parse_hex_color(&theme.cell_bg);
        let selection_bg = parse_hex_color(&theme.selection_bg);
        let number_color = parse_hex_color(&theme.number_color);
        let cmd_color    = parse_hex_color(&theme.command_mode_color);
        let vis_color    = parse_hex_color(&theme.visual_mode_color);
        let insert_color = parse_hex_color(&theme.insert_mode_color);
        let normal_color = parse_hex_color(&theme.normal_mode_color);

        // Layout: top row = column headers, rest = data rows
        let header_y = area.y;
        let data_y = area.y + 1;
        let data_height = area.height.saturating_sub(1);

        // Compute column widths that fit in available area
        let available_width = area.width.saturating_sub(ROW_GUTTER_WIDTH);
        let mut cols: Vec<(u32, u16)> = Vec::new(); // (col_idx, display_width)
        let mut x_offset = 0u16;

        let mut c = viewport.left_col;
        loop {
            if x_offset >= available_width { break; }
            let col_width = sheet.col_width(c).max(MIN_COL_WIDTH);
            cols.push((c, col_width.min(available_width - x_offset)));
            x_offset += col_width.min(available_width - x_offset);
            if x_offset >= available_width { break; }
            c += 1;
            if c > 10000 { break; } // safety
        }

        // ── Header row ───────────────────────────────────────────────────────
        let corner_style = Style::default()
            .fg(header_fg)
            .bg(header_bg)
            .add_modifier(Modifier::BOLD);
        render_cell_str(buf, area.x, header_y, ROW_GUTTER_WIDTH, "     ", corner_style);

        let mut x = area.x + ROW_GUTTER_WIDTH;
        for (col_idx, col_width) in &cols {
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

        // ── Data rows ────────────────────────────────────────────────────────
        let mut screen_y = data_y;
        let mut row_idx = viewport.top_row;
        while screen_y < data_y + data_height {
            let row_h = sheet.row_height(row_idx).max(1) as u16;
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
                Style::default()
                    .fg(header_fg)
                    .bg(header_bg)
            };
            render_cell_str(buf, area.x, screen_y, ROW_GUTTER_WIDTH, &gutter_text, gutter_style);

            let mut x = area.x + ROW_GUTTER_WIDTH;
            for (col_idx, col_width) in &cols {
                let is_cursor = row_idx == cursor.row && *col_idx == cursor.col;

                // Formula reference selection highlighting
                let is_fref_cursor = matches!(&input.mode, Mode::FormulaSelect { .. })
                    && row_idx == cursor.row && *col_idx == cursor.col;
                let is_fref_range = if let Mode::FormulaSelect { anchor: Some((ar, ac)) } = &input.mode {
                    let r_min = (*ar).min(cursor.row);
                    let r_max = (*ar).max(cursor.row);
                    let c_min = (*ac).min(cursor.col);
                    let c_max = (*ac).max(cursor.col);
                    row_idx >= r_min && row_idx <= r_max
                        && *col_idx >= c_min && *col_idx <= c_max
                } else { false };

                let is_visual = is_in_visual_selection(
                    row_idx, *col_idx,
                    input.visual_anchor.as_ref(),
                    cursor,
                    &input.mode,
                );
                let search_hl = input.search_highlight(row_idx, *col_idx);

                // In Insert/FormulaSelect mode, show the live edit buffer for the cursor cell
                let live_edit = is_cursor && matches!(&input.mode, Mode::Insert { .. } | Mode::FormulaSelect { .. });
                let display = if live_edit {
                    input.edit_buffer.clone()
                } else {
                    sheet.display_value(row_idx, *col_idx)
                };
                let raw_value = sheet.get_value(row_idx, *col_idx);

                let cell_style = if is_fref_cursor {
                    // Formula reference picker — bright green (normal_color) for cursor cell
                    Style::default()
                        .fg(Color::Black)
                        .bg(normal_color)
                        .add_modifier(Modifier::BOLD)
                } else if is_fref_range {
                    // Range being selected for formula
                    Style::default()
                        .fg(Color::Black)
                        .bg(darken(insert_color, 0.55))
                } else if is_cursor {
                    Style::default()
                        .fg(Color::Black)
                        .bg(cursor_bg)
                        .add_modifier(Modifier::BOLD)
                } else if is_visual {
                    Style::default()
                        .fg(Color::Black)
                        .bg(selection_bg)
                } else if search_hl == Some(true) {
                    // Current search match — uses command_mode_color
                    Style::default()
                        .fg(Color::Black)
                        .bg(cmd_color)
                        .add_modifier(Modifier::BOLD)
                } else if search_hl == Some(false) {
                    // Other search matches — dimmed command_mode_color
                    Style::default()
                        .fg(Color::Black)
                        .bg(darken(vis_color, 0.75))
                } else {
                    // Normal cell — apply CellStyle if present
                    let user_style: Option<&CellStyle> = sheet
                        .get_cell(row_idx, *col_idx)
                        .and_then(|c| c.style.as_ref());

                    let cell_bg = user_style
                        .and_then(|s| s.bg.as_ref())
                        .map(|c| Color::Rgb(c.r, c.g, c.b))
                        .unwrap_or(cell_bg);

                    let default_fg = if matches!(raw_value, CellValue::Error(_)) {
                        cmd_color // errors use command colour (usually red/orange)
                    } else if matches!(raw_value, CellValue::Number(_) | CellValue::Boolean(_)) {
                        number_color
                    } else {
                        Color::White
                    };

                    let cell_fg = user_style
                        .and_then(|s| s.fg.as_ref())
                        .map(|c| Color::Rgb(c.r, c.g, c.b))
                        .unwrap_or(default_fg);

                    let mut s = Style::default().fg(cell_fg).bg(cell_bg);
                    if let Some(us) = user_style {
                        if us.bold          { s = s.add_modifier(Modifier::BOLD);        }
                        if us.italic        { s = s.add_modifier(Modifier::ITALIC);      }
                        if us.underline     { s = s.add_modifier(Modifier::UNDERLINED);  }
                        if us.strikethrough { s = s.add_modifier(Modifier::CROSSED_OUT); }
                    }
                    s
                };

                // Alignment: user style overrides; default is right for numbers/booleans
                use asat_core::Alignment;
                let user_align = sheet.get_cell(row_idx, *col_idx)
                    .and_then(|c| c.style.as_ref())
                    .map(|s| s.align.clone());
                match user_align {
                    Some(Alignment::Right) =>
                        render_cell_right(buf, x, screen_y, *col_width, &display, cell_style),
                    Some(Alignment::Center) =>
                        render_cell_centered(buf, x, screen_y, *col_width, &display, cell_style),
                    Some(Alignment::Left) =>
                        render_cell_str(buf, x, screen_y, *col_width, &display, cell_style),
                    _ => {
                        if matches!(raw_value, CellValue::Number(_) | CellValue::Boolean(_)) {
                            render_cell_right(buf, x, screen_y, *col_width, &display, cell_style);
                        } else {
                            render_cell_str(buf, x, screen_y, *col_width, &display, cell_style);
                        }
                    }
                }

                x += col_width;
            }

            // ── Extra lines for tall rows (background fill) ───────────────────
            for extra in 1..row_h {
                let ey = screen_y + extra;
                if ey >= data_y + data_height { break; }

                // Gutter continuation marker
                let cont_style = if is_cursor_row {
                    gutter_style
                } else {
                    Style::default()
                        .fg(darken(header_fg, 0.6))
                        .bg(header_bg)
                };
                render_cell_str(buf, area.x, ey, ROW_GUTTER_WIDTH, "    │", cont_style);

                // Fill cells with background
                let mut ex = area.x + ROW_GUTTER_WIDTH;
                for (_, col_width) in &cols {
                    let bg_style = if is_cursor_row {
                        Style::default().fg(Color::Black).bg(cell_bg)
                    } else {
                        Style::default().fg(Color::White).bg(cell_bg)
                    };
                    render_cell_str(buf, ex, ey, *col_width, "", bg_style);
                    ex += col_width;
                }
            }

            screen_y += row_h;
            row_idx += 1;
        }
    }
}

fn is_in_visual_selection(
    row: u32, col: u32,
    anchor: Option<&VisualAnchor>,
    cursor: asat_input::Cursor,
    mode: &Mode,
) -> bool {
    let anchor = match anchor { Some(a) => a, None => return false };
    let r_min = cursor.row.min(anchor.row);
    let r_max = cursor.row.max(anchor.row);
    match mode {
        Mode::VisualLine => {
            row >= r_min && row <= r_max
        }
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
    if width == 0 { return; }
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
    let text_budget = if truncated && width > 1 { width - 1 } else { width };

    let mut dx = 0u16;
    for ch in content.chars() {
        if dx >= text_budget { break; }
        let cw = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(1) as u16;
        if dx + cw > text_budget { break; }
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
    if width == 0 { return; }
    for dx in 0..width {
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(' ');
            cell.set_style(style);
        }
    }
    let display_width = UnicodeWidthStr::width(content) as u16;
    if display_width > width {
        // Too wide — show #### like Excel (reversed style = always readable)
        let hash_style = style.add_modifier(Modifier::REVERSED).add_modifier(Modifier::BOLD);
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
        if dx >= width { break; }
        let cw = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(1) as u16;
        if dx + cw > width { break; }
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(ch);
            cell.set_style(style);
        }
        dx += cw;
    }
}

/// Center-aligned cell render (for column headers)
fn render_cell_centered(buf: &mut Buffer, x: u16, y: u16, width: u16, content: &str, style: Style) {
    if width == 0 { return; }
    for dx in 0..width {
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(' ');
            cell.set_style(style);
        }
    }
    let display_width = UnicodeWidthStr::width(content) as u16;
    let start_dx = if display_width < width { (width - display_width) / 2 } else { 0 };
    let mut dx = start_dx;
    for ch in content.chars() {
        if dx >= width { break; }
        let ch_width = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(1) as u16;
        if dx + ch_width > width { break; }
        if let Some(cell) = buf.cell_mut((x + dx, y)) {
            cell.set_char(ch);
            cell.set_style(style);
        }
        dx += ch_width;
    }
}
