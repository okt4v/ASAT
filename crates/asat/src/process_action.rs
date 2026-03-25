#![allow(clippy::too_many_arguments)]

use std::path::PathBuf;

use arboard::Clipboard;
use crossterm::event::KeyEvent;
use regex::Regex;

use asat_commands::{
    Command, DeleteCol, InsertCol, MergeCells, SetCell, UndoStack, UnmergeCells,
};
use asat_config::Config;
use asat_core::{cell_address, CellStyle, CellValue, Workbook};
use asat_input::{AppAction, InputState, Mode};
use asat_plugins::{PluginEvent, PluginManager};

use crate::app::{
    apply_style_sel, auto_fill_series, cells_to_tsv, copy_to_clipboard,
    cycle_date, cycle_text_sequence, find_first_cell_ref, open_config_in_editor,
    parse_cell_address, push_recent, save_recent_files, scan_files, set_status,
    ActionResult,
};
use crate::ex_commands::handle_ex_command;

pub(crate) fn process_action(
    action: AppAction,
    workbook: &mut Workbook,
    input: &mut InputState,
    undo: &mut UndoStack,
    status: &mut Option<(String, std::time::Instant)>,
    config: &mut Config,
    plugins: &mut PluginManager,
    clipboard: &mut Option<Clipboard>,
    visible_rows: u32,
    visible_cols: u32,
    edit_count: &mut u32,
) -> ActionResult {
    match action {
        AppAction::NoOp => {}

        // ── Navigation ──
        AppAction::MoveCursor {
            row_delta,
            col_delta,
        } => {
            let new_row = (input.cursor.row as i64 + row_delta as i64).max(0) as u32;
            let new_col = (input.cursor.col as i64 + col_delta as i64).max(0) as u32;
            let (sr, sc) = if let Some(m) = workbook.active().merge_at(new_row, new_col) {
                let at_anchor = input.cursor.row == m.row_start && input.cursor.col == m.col_start;
                if at_anchor {
                    // Moving out of the anchor: jump past the far edge of the merge
                    let r = if row_delta > 0 {
                        m.row_end + 1
                    } else if row_delta < 0 {
                        m.row_start.saturating_sub(1)
                    } else {
                        new_row
                    };
                    let c = if col_delta > 0 {
                        m.col_end + 1
                    } else if col_delta < 0 {
                        m.col_start.saturating_sub(1)
                    } else {
                        new_col
                    };
                    // The position after the merge might itself be covered by another merge
                    workbook.active().snap_to_anchor(r, c)
                } else {
                    // Entering the merge from outside → land on anchor
                    (m.row_start, m.col_start)
                }
            } else {
                (new_row, new_col)
            };
            input.cursor.row = sr;
            input.cursor.col = sc;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::MoveCursorTo { row, col } => {
            let (sr, sc) = workbook.active().snap_to_anchor(row, col);
            input.cursor.row = sr;
            input.cursor.col = sc;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::MoveToFirstRow => {
            let (sr, sc) = workbook.active().snap_to_anchor(0, input.cursor.col);
            input.cursor.row = sr;
            input.cursor.col = sc;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::MoveToLastRow => {
            let last = workbook.active().max_row();
            let (sr, sc) = workbook.active().snap_to_anchor(last, input.cursor.col);
            input.cursor.row = sr;
            input.cursor.col = sc;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::MoveToFirstCol => {
            let (sr, sc) = workbook.active().snap_to_anchor(input.cursor.row, 0);
            input.cursor.row = sr;
            input.cursor.col = sc;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::MoveToLastCol => {
            let last = workbook.active().max_col();
            let (sr, sc) = workbook.active().snap_to_anchor(input.cursor.row, last);
            input.cursor.row = sr;
            input.cursor.col = sc;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::PageDown => {
            let new_row = (input.cursor.row + visible_rows).min(workbook.active().max_row());
            let (sr, sc) = workbook.active().snap_to_anchor(new_row, input.cursor.col);
            input.cursor.row = sr;
            input.cursor.col = sc;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::PageUp => {
            let new_row = input.cursor.row.saturating_sub(visible_rows);
            let (sr, sc) = workbook.active().snap_to_anchor(new_row, input.cursor.col);
            input.cursor.row = sr;
            input.cursor.col = sc;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }

        // ── Mode ──
        AppAction::EnterInsert { replace } => {
            input.mode = Mode::Insert { replace };
        }

        AppAction::EnterVisual { block } => {
            input.mode = Mode::Visual { block };
        }
        AppAction::EnterCommand => {
            input.mode = Mode::Command;
        }
        AppAction::EnterSearch { forward } => {
            input.mode = Mode::Search { forward };
        }
        AppAction::ExitMode => {
            input.visual_command_range = None;
            input.visual_anchor = None;
            // Don't reset the mode if ExecuteCommand2 just switched to a screen overlay.
            // Those modes handle their own Exit (q/Esc sets mode back to Normal themselves).
            if !matches!(
                input.mode,
                Mode::Help
                    | Mode::PluginManager
                    | Mode::ThemeManager
                    | Mode::FileFind
                    | Mode::RecentFiles
                    | Mode::Welcome
            ) {
                input.mode = Mode::Normal;
            }
        }

        // ── Cell editing ──
        AppAction::SetCell {
            sheet: _,
            row,
            col,
            value,
        } => {
            let sheet_idx = workbook.active_sheet;
            let old_val = workbook.active().get_value(row, col).clone();
            let cmd = Box::new(SetCell::new(workbook, sheet_idx, row, col, value.clone()));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                *edit_count += 1;
                plugins.push_event(PluginEvent::CellChange {
                    sheet: sheet_idx,
                    row,
                    col,
                    old: old_val,
                    new: value,
                });
            }
        }
        AppAction::DeleteCellContent => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            let col = input.cursor.col;
            // Yank deleted content + style to register + clipboard (like Vim)
            let val = workbook.active().get_raw_value(row, col).clone();
            let sty = workbook
                .active()
                .get_cell(row, col)
                .and_then(|c| c.style.clone());
            if !matches!(val, CellValue::Empty) || sty.is_some() {
                let text = val.display();
                copy_to_clipboard(clipboard, &text);
                input
                    .registers
                    .yank_at(None, vec![vec![val]], vec![vec![sty]], false, row, col);
            }
            let mut cmd = Box::new(SetCell::new(
                workbook,
                sheet_idx,
                row,
                col,
                CellValue::Empty,
            ));
            cmd.new_style = None;
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
            }
        }
        AppAction::DeleteCellRange {
            row_start,
            col_start,
            row_end,
            col_end,
        } => {
            let sheet_idx = workbook.active_sheet;
            // Clamp to actual data bounds first — row_end/col_end may be u32::MAX
            // for V-ROW (col_end=MAX) and V-COL (row_end=MAX) modes.
            let (row_end, col_end) = {
                let s = workbook.active();
                (row_end.min(s.max_row()), col_end.min(s.max_col()))
            };
            // Yank selection values + styles to register + clipboard before deleting
            {
                let sheet = workbook.active();
                let mut cells: Vec<Vec<CellValue>> = Vec::new();
                let mut styles: Vec<Vec<Option<CellStyle>>> = Vec::new();
                for r in row_start..=row_end {
                    let row_vals: Vec<CellValue> = (col_start..=col_end)
                        .map(|c| sheet.get_raw_value(r, c).clone())
                        .collect();
                    let row_styles: Vec<Option<CellStyle>> = (col_start..=col_end)
                        .map(|c| sheet.get_cell(r, c).and_then(|cell| cell.style.clone()))
                        .collect();
                    cells.push(row_vals);
                    styles.push(row_styles);
                }
                let tsv = cells_to_tsv(&cells);
                copy_to_clipboard(clipboard, &tsv);
                input
                    .registers
                    .yank_at(None, cells, styles, false, row_start, col_start);
            }
            // Only clear cells that actually have data (sparse-safe, avoids huge loops)
            let coords: Vec<(u32, u32)> = workbook
                .active()
                .cells
                .keys()
                .filter(|(r, c)| {
                    *r >= row_start && *r <= row_end && *c >= col_start && *c <= col_end
                })
                .cloned()
                .collect();
            if coords.is_empty() {
                return ActionResult::Continue;
            }
            let cmds: Vec<Box<dyn asat_commands::Command>> = coords
                .into_iter()
                .map(|(r, c)| {
                    let mut cmd =
                        Box::new(SetCell::new(workbook, sheet_idx, r, c, CellValue::Empty));
                    cmd.new_style = None;
                    cmd as Box<dyn asat_commands::Command>
                })
                .collect();
            let grouped = Box::new(asat_commands::GroupedCommand {
                description: "delete selection".to_string(),
                commands: cmds,
            });
            if let Err(e) = grouped.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(grouped);
            }
        }

        // ── Row/Col operations ──
        AppAction::InsertRowAbove => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            let cmd = Box::new(asat_commands::InsertRow {
                sheet: sheet_idx,
                row,
            });
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
            }
        }
        AppAction::DeleteCurrentRow => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            // Yank the row values + styles to register + clipboard before deleting
            let sheet = workbook.active();
            let max_col = sheet.max_col();
            let cells: Vec<Vec<CellValue>> = vec![(0..=max_col)
                .map(|c| sheet.get_raw_value(row, c).clone())
                .collect()];
            let styles: Vec<Vec<Option<CellStyle>>> = vec![(0..=max_col)
                .map(|c| sheet.get_cell(row, c).and_then(|cell| cell.style.clone()))
                .collect()];
            let tsv = cells_to_tsv(&cells);
            copy_to_clipboard(clipboard, &tsv);
            input.registers.yank_at(None, cells, styles, true, row, 0);
            let cmd = Box::new(asat_commands::DeleteRow::new(workbook, sheet_idx, row));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
            }
        }
        AppAction::DeleteRowAt { row } => {
            let sheet_idx = workbook.active_sheet;
            let max_row = workbook.active().max_row();
            if row > max_row {
                set_status(status, "No row there".to_string());
            } else {
                let cmd = Box::new(asat_commands::DeleteRow::new(workbook, sheet_idx, row));
                if let Err(e) = cmd.execute(workbook) {
                    set_status(status, format!("Error: {}", e));
                } else {
                    undo.push(cmd);
                    set_status(status, format!("Deleted row {}", row + 1));
                }
            }
        }

        // ── Undo/Redo ──
        AppAction::Undo => match undo.undo(workbook) {
            Ok(Some((sheet, row, col))) => {
                workbook.active_sheet = sheet;
                input.cursor.row = row;
                input.cursor.col = col;
                input.scroll_to_cursor(visible_rows, visible_cols);
                set_status(status, "Undo".to_string());
            }
            Ok(None) => set_status(status, "Already at oldest change".to_string()),
            Err(e) => set_status(status, format!("Undo error: {}", e)),
        },
        AppAction::Redo => match undo.redo(workbook) {
            Ok(Some((sheet, row, col))) => {
                workbook.active_sheet = sheet;
                input.cursor.row = row;
                input.cursor.col = col;
                input.scroll_to_cursor(visible_rows, visible_cols);
                set_status(status, "Redo".to_string());
            }
            Ok(None) => set_status(status, "Already at newest change".to_string()),
            Err(e) => set_status(status, format!("Redo error: {}", e)),
        },

        // ── Command execution ──
        AppAction::ExecuteCommand2(cmd_str) => {
            return handle_ex_command(&cmd_str, workbook, input, undo, status, config, plugins);
        }

        // ── Sheet navigation ──
        AppAction::NextSheet => {
            if workbook.active_sheet + 1 < workbook.sheets.len() {
                let from = workbook.active_sheet;
                workbook.active_sheet += 1;
                plugins.push_event(PluginEvent::SheetChange {
                    from,
                    to: workbook.active_sheet,
                });
            }
        }
        AppAction::PrevSheet => {
            if workbook.active_sheet > 0 {
                let from = workbook.active_sheet;
                workbook.active_sheet -= 1;
                plugins.push_event(PluginEvent::SheetChange {
                    from,
                    to: workbook.active_sheet,
                });
            }
        }

        // ── Quit / Save ──
        AppAction::Quit => return ActionResult::Quit,
        AppAction::QuitForce => return ActionResult::ForceQuit,
        AppAction::Save => {
            if let Some(path) = workbook.file_path.clone() {
                let path_str = path.display().to_string();
                plugins.push_event(PluginEvent::PreSave {
                    path: path_str.clone(),
                });
                if config.backup_on_save && path.exists() {
                    let bak = path.with_extension(format!(
                        "{}.bak",
                        path.extension().and_then(|e| e.to_str()).unwrap_or("")
                    ));
                    let _ = std::fs::copy(&path, &bak);
                }
                match asat_io::save(workbook, &path) {
                    Ok(_) => {
                        workbook.dirty = false;
                        plugins.push_event(PluginEvent::PostSave { path: path_str });
                        set_status(
                            status,
                            format!("Saved {:?}", path.file_name().unwrap_or_default()),
                        );
                    }
                    Err(e) => set_status(status, format!("Save error: {}", e)),
                }
            } else {
                set_status(status, "No file path. Use :w <filename>".to_string());
            }
        }

        // ── Mode transitions ──
        AppAction::EnterVisualLine => {
            input.mode = Mode::VisualLine;
        }

        // ── New navigation ──
        AppAction::MoveToNextNonEmptyH { forward } => {
            let sheet = workbook.active();
            let row = input.cursor.row;
            let max_col = sheet.max_col();
            let col = if forward {
                let mut c = input.cursor.col + 1;
                while c <= max_col && sheet.get_value(row, c).is_empty() {
                    c += 1;
                }
                c.min(max_col)
            } else {
                let mut c = input.cursor.col.saturating_sub(1);
                loop {
                    if !sheet.get_value(row, c).is_empty() {
                        break;
                    }
                    if c == 0 {
                        break;
                    }
                    c -= 1;
                }
                c
            };
            input.cursor.col = col;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::MoveToNextNonEmptyV { forward } => {
            let sheet = workbook.active();
            let col = input.cursor.col;
            let max_row = sheet.max_row();
            let row = if forward {
                let mut r = input.cursor.row + 1;
                while r <= max_row && sheet.get_value(r, col).is_empty() {
                    r += 1;
                }
                r.min(max_row)
            } else {
                let mut r = input.cursor.row.saturating_sub(1);
                loop {
                    if !sheet.get_value(r, col).is_empty() {
                        break;
                    }
                    if r == 0 {
                        break;
                    }
                    r -= 1;
                }
                r
            };
            input.cursor.row = row;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::MoveToNextParagraph { forward } => {
            // Jump past any non-empty rows to the next empty row (or vice-versa)
            let sheet = workbook.active();
            let max_row = sheet.max_row();
            let col = input.cursor.col;
            let row = if forward {
                let mut r = input.cursor.row + 1;
                // Skip current non-empty block
                while r <= max_row && !sheet.get_value(r, col).is_empty() {
                    r += 1;
                }
                // Skip empty rows
                while r <= max_row && sheet.get_value(r, col).is_empty() {
                    r += 1;
                }
                r.min(max_row)
            } else {
                let mut r = input.cursor.row.saturating_sub(1);
                loop {
                    if sheet.get_value(r, col).is_empty() {
                        break;
                    }
                    if r == 0 {
                        break;
                    }
                    r -= 1;
                }
                loop {
                    if !sheet.get_value(r, col).is_empty() {
                        break;
                    }
                    if r == 0 {
                        break;
                    }
                    r -= 1;
                }
                r
            };
            input.save_position(workbook.active_sheet);
            input.cursor.row = row;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::JumpHighRow => {
            input.cursor.row = input.viewport.top_row;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::JumpMidRow => {
            input.cursor.row = input.viewport.top_row + visible_rows / 2;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::JumpLowRow => {
            input.cursor.row = (input.viewport.top_row + visible_rows).saturating_sub(1);
            input.scroll_to_cursor(visible_rows, visible_cols);
        }
        AppAction::ScrollCenter => {
            input.viewport.top_row = input.cursor.row.saturating_sub(visible_rows / 2);
        }
        AppAction::ScrollTop => {
            input.viewport.top_row = input.cursor.row;
        }
        AppAction::ScrollBottom => {
            input.viewport.top_row = input
                .cursor
                .row
                .saturating_sub(visible_rows.saturating_sub(1));
        }

        // ── Marks ──
        AppAction::SetMark { ch } => {
            input
                .marks
                .insert(ch, (workbook.active_sheet, input.cursor));
            set_status(status, format!("Mark '{}'", ch));
        }
        AppAction::JumpToMark { ch } => {
            if let Some((sheet_idx, cursor)) = input.marks.get(&ch).cloned() {
                input.save_position(workbook.active_sheet);
                workbook.active_sheet = sheet_idx;
                input.cursor = cursor;
                input.scroll_to_cursor(visible_rows, visible_cols);
            } else {
                set_status(status, format!("Mark '{}' not set", ch));
            }
        }
        AppAction::JumpToPrevPos => {
            if let Some((sheet_idx, cursor)) = input.prev_position.take() {
                let cur_sheet = workbook.active_sheet;
                let cur_cursor = input.cursor;
                workbook.active_sheet = sheet_idx;
                input.cursor = cursor;
                input.prev_position = Some((cur_sheet, cur_cursor));
                input.scroll_to_cursor(visible_rows, visible_cols);
            }
        }

        // ── Cell operations ──
        AppAction::ChangeCell => {
            // Clear cell and enter insert with empty buffer
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            let col = input.cursor.col;
            let cmd = Box::new(SetCell::new(
                workbook,
                sheet_idx,
                row,
                col,
                CellValue::Empty,
            ));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                input.edit_buffer.clear();
                input.edit_cursor_pos = 0;
                input.mode = Mode::Insert { replace: false };
            }
        }
        AppAction::ToggleCase => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            let col = input.cursor.col;
            let current = workbook.active().get_value(row, col).clone();
            if let CellValue::Text(s) = &current {
                let toggled = if s.chars().next().map(|c| c.is_uppercase()).unwrap_or(false) {
                    s.to_lowercase()
                } else {
                    s.to_uppercase()
                };
                let cmd = Box::new(SetCell::new(
                    workbook,
                    sheet_idx,
                    row,
                    col,
                    CellValue::Text(toggled),
                ));
                if let Err(e) = cmd.execute(workbook) {
                    set_status(status, format!("Error: {}", e));
                } else {
                    undo.push(cmd);
                }
            }
        }

        // ── Row operations ──
        AppAction::OpenRowBelow => {
            let sheet_idx = workbook.active_sheet;
            let new_row = input.cursor.row + 1;
            let cmd = Box::new(asat_commands::InsertRow {
                sheet: sheet_idx,
                row: new_row,
            });
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                input.cursor.row = new_row;
                input.cursor.col = 0;
                input.edit_buffer.clear();
                input.edit_cursor_pos = 0;
                input.mode = Mode::Insert { replace: false };
                input.scroll_to_cursor(visible_rows, visible_cols);
            }
        }
        AppAction::OpenRowAbove => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            let cmd = Box::new(asat_commands::InsertRow {
                sheet: sheet_idx,
                row,
            });
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                input.cursor.col = 0;
                input.edit_buffer.clear();
                input.edit_cursor_pos = 0;
                input.mode = Mode::Insert { replace: false };
                input.scroll_to_cursor(visible_rows, visible_cols);
            }
        }

        // ── Column width ──
        AppAction::IncreaseColWidth { col } => {
            let sheet = workbook.active_mut();
            let meta = sheet.col_meta.entry(col).or_default();
            meta.width = Some(meta.width.unwrap_or(10) + 2);
            workbook.dirty = true;
            let w = workbook.active().col_width(col);
            set_status(status, format!("Column width: {}", w));
        }
        AppAction::DecreaseColWidth { col } => {
            let sheet = workbook.active_mut();
            let meta = sheet.col_meta.entry(col).or_default();
            let current = meta.width.unwrap_or(10);
            meta.width = Some(current.saturating_sub(2).max(3));
            workbook.dirty = true;
            let w = workbook.active().col_width(col);
            set_status(status, format!("Column width: {}", w));
        }
        AppAction::AutoFitCol { col } => {
            let sheet = workbook.active();
            let max_row = sheet.max_row();
            let max_w = (0..=max_row)
                .map(|r| sheet.display_value(r, col).len() as u16)
                .max()
                .unwrap_or(3)
                .max(3)
                + 1;
            workbook.active_mut().col_meta.entry(col).or_default().width = Some(max_w);
            workbook.dirty = true;
            set_status(status, format!("Column width: {} (auto-fit)", max_w));
        }

        // ── Row height ──
        AppAction::IncreaseRowHeight { row } => {
            let sheet = workbook.active_mut();
            let meta = sheet.row_meta.entry(row).or_default();
            meta.height = Some(meta.height.unwrap_or(1) + 1);
            workbook.dirty = true;
            let h = workbook.active().row_height(row);
            set_status(status, format!("Row height: {}", h));
        }
        AppAction::DecreaseRowHeight { row } => {
            let sheet = workbook.active_mut();
            let meta = sheet.row_meta.entry(row).or_default();
            let current = meta.height.unwrap_or(1);
            let new_h = current.saturating_sub(1).max(1);
            meta.height = Some(new_h);
            workbook.dirty = true;
            set_status(status, format!("Row height: {}", new_h));
        }
        AppAction::AutoFitRow { row } => {
            if let Some(meta) = workbook.active_mut().row_meta.get_mut(&row) {
                meta.height = None;
            }
            workbook.dirty = true;
            set_status(status, "Row height: auto (1)".to_string());
        }

        // ── Search ──
        AppAction::SearchCurrentCell => {
            let val = workbook
                .active()
                .get_value(input.cursor.row, input.cursor.col)
                .display();
            if !val.is_empty() {
                input.last_search = Some((val.clone(), true));
                set_status(status, format!("/{}", val));
            }
        }

        // ── Yank ──
        AppAction::YankRow => {
            let sheet = workbook.active();
            let row = input.cursor.row;
            let max_col = sheet.max_col();
            let cells: Vec<Vec<CellValue>> = vec![(0..=max_col)
                .map(|c| sheet.get_raw_value(row, c).clone())
                .collect()];
            let styles: Vec<Vec<Option<CellStyle>>> = vec![(0..=max_col)
                .map(|c| sheet.get_cell(row, c).and_then(|cell| cell.style.clone()))
                .collect()];
            let tsv = cells_to_tsv(&cells);
            copy_to_clipboard(clipboard, &tsv);
            input.registers.yank_at(None, cells, styles, true, row, 0);
            set_status(status, format!("Yanked row {} → clipboard", row + 1));
        }
        AppAction::YankCell => {
            let sheet = workbook.active();
            let (row, col) = (input.cursor.row, input.cursor.col);
            let val = sheet.get_raw_value(row, col).clone();
            let sty = sheet.get_cell(row, col).and_then(|c| c.style.clone());
            let text = val.display();
            copy_to_clipboard(clipboard, &text);
            input
                .registers
                .yank_at(None, vec![vec![val]], vec![vec![sty]], false, row, col);
            set_status(status, format!("Yanked cell → clipboard: {}", text));
        }
        AppAction::YankRowAt { row } => {
            let sheet = workbook.active();
            let max_row = sheet.max_row();
            if row > max_row {
                set_status(status, "No row there".to_string());
            } else {
                let max_col = sheet.max_col();
                let cells: Vec<Vec<CellValue>> = vec![(0..=max_col)
                    .map(|c| sheet.get_raw_value(row, c).clone())
                    .collect()];
                let styles: Vec<Vec<Option<CellStyle>>> = vec![(0..=max_col)
                    .map(|c| sheet.get_cell(row, c).and_then(|cell| cell.style.clone()))
                    .collect()];
                let tsv = cells_to_tsv(&cells);
                copy_to_clipboard(clipboard, &tsv);
                input.registers.yank_at(None, cells, styles, true, row, 0);
                set_status(status, format!("Yanked row {} → clipboard", row + 1));
            }
        }
        AppAction::YankCol => {
            let sheet = workbook.active();
            let col = input.cursor.col;
            let max_row = sheet.max_row();
            let cells: Vec<Vec<CellValue>> = (0..=max_row)
                .map(|r| vec![sheet.get_raw_value(r, col).clone()])
                .collect();
            let styles: Vec<Vec<Option<CellStyle>>> = (0..=max_row)
                .map(|r| vec![sheet.get_cell(r, col).and_then(|c| c.style.clone())])
                .collect();
            let tsv = cells_to_tsv(&cells);
            copy_to_clipboard(clipboard, &tsv);
            input.registers.yank_at(None, cells, styles, false, 0, col);
            set_status(status, format!("Yanked column {} → clipboard", col + 1));
        }
        AppAction::YankCellRange {
            row_start,
            col_start,
            row_end,
            col_end,
            is_line,
        } => {
            let sheet = workbook.active();
            // Clamp both axes — u32::MAX is used for V-ROW (col_end) and V-COL (row_end)
            let row_end = row_end.min(sheet.max_row());
            let col_end = col_end.min(sheet.max_col());
            if row_start > row_end || col_start > col_end {
                set_status(status, "Nothing to yank".to_string());
                return ActionResult::Continue;
            }
            let cells: Vec<Vec<CellValue>> = (row_start..=row_end)
                .map(|r| {
                    (col_start..=col_end)
                        .map(|c| sheet.get_raw_value(r, c).clone())
                        .collect()
                })
                .collect();
            let styles: Vec<Vec<Option<CellStyle>>> = (row_start..=row_end)
                .map(|r| {
                    (col_start..=col_end)
                        .map(|c| sheet.get_cell(r, c).and_then(|cell| cell.style.clone()))
                        .collect()
                })
                .collect();
            let rows = cells.len();
            let cols = cells.first().map(|r| r.len()).unwrap_or(0);
            let tsv = cells_to_tsv(&cells);
            copy_to_clipboard(clipboard, &tsv);
            input
                .registers
                .yank_at(None, cells, styles, is_line, row_start, col_start);
            set_status(status, format!("Yanked {}x{} → clipboard", rows, cols));
        }

        // ── Paste ──
        AppAction::PasteAfter => do_paste(workbook, input, undo, status, true),
        AppAction::PasteBefore => do_paste(workbook, input, undo, status, false),

        // ── Search ──
        AppAction::ExecuteSearch => {
            if let Some((pattern, forward)) = input.last_search.clone() {
                let sheet = workbook.active();
                let max_row = sheet.max_row();
                let max_col = sheet.max_col();

                // Try to compile as regex (case-insensitive); fall back to literal substring
                let regex = Regex::new(&format!("(?i){}", &pattern))
                    .or_else(|_| Regex::new(&format!("(?i){}", regex::escape(&pattern))))
                    .ok();

                let mut matches: Vec<(u32, u32)> = Vec::new();
                for r in 0..=max_row {
                    for c in 0..=max_col {
                        let val = sheet.display_value(r, c);
                        let hit = if let Some(ref re) = regex {
                            re.is_match(&val)
                        } else {
                            val.to_lowercase().contains(&pattern.to_lowercase())
                        };
                        if hit {
                            matches.push((r, c));
                        }
                    }
                }

                if matches.is_empty() {
                    input.search_matches.clear();
                    input.search_match_set.clear();
                    set_status(status, format!("Pattern not found: {}", pattern));
                } else {
                    let cr = input.cursor.row;
                    let cc = input.cursor.col;
                    let idx = if forward {
                        matches
                            .iter()
                            .position(|&(r, c)| r > cr || (r == cr && c > cc))
                            .unwrap_or(0)
                    } else {
                        matches
                            .iter()
                            .rposition(|&(r, c)| r < cr || (r == cr && c < cc))
                            .unwrap_or(matches.len() - 1)
                    };
                    input.save_position(workbook.active_sheet);
                    input.cursor.row = matches[idx].0;
                    input.cursor.col = matches[idx].1;
                    input.scroll_to_cursor(visible_rows, visible_cols);
                    let total = matches.len();
                    input.search_match_set = matches.iter().copied().collect();
                    input.search_matches = matches;
                    input.search_match_idx = idx;
                    set_status(status, format!("/{} [{}/{}]", pattern, idx + 1, total));
                }
            }
        }
        AppAction::ClearSearch => {
            input.search_matches.clear();
            input.search_match_set.clear();
        }
        AppAction::FindNext => {
            if input.search_matches.is_empty() {
                set_status(status, "No active search — press / to search".to_string());
            } else {
                input.search_match_idx = (input.search_match_idx + 1) % input.search_matches.len();
                let (r, c) = input.search_matches[input.search_match_idx];
                input.cursor.row = r;
                input.cursor.col = c;
                input.scroll_to_cursor(visible_rows, visible_cols);
                if let Some((pat, _)) = &input.last_search {
                    set_status(
                        status,
                        format!(
                            "/{} [{}/{}]",
                            pat,
                            input.search_match_idx + 1,
                            input.search_matches.len()
                        ),
                    );
                }
            }
        }
        AppAction::FindPrev => {
            if input.search_matches.is_empty() {
                set_status(status, "No active search — press / to search".to_string());
            } else {
                input.search_match_idx = if input.search_match_idx == 0 {
                    input.search_matches.len() - 1
                } else {
                    input.search_match_idx - 1
                };
                let (r, c) = input.search_matches[input.search_match_idx];
                input.cursor.row = r;
                input.cursor.col = c;
                input.scroll_to_cursor(visible_rows, visible_cols);
                if let Some((pat, _)) = &input.last_search {
                    set_status(
                        status,
                        format!(
                            "/{} [{}/{}]",
                            pat,
                            input.search_match_idx + 1,
                            input.search_matches.len()
                        ),
                    );
                }
            }
        }

        // ── Welcome screen actions ──
        AppAction::WelcomeNewFile => {
            if workbook.dirty {
                set_status(
                    status,
                    "Unsaved changes — use :w to save or :q! to discard".to_string(),
                );
            } else {
                *workbook = Workbook::new();
                input.mode = Mode::Normal;
                set_status(
                    status,
                    "New workbook — press i to start editing".to_string(),
                );
            }
        }
        AppAction::WelcomeEnterFileFind => {
            input.finder_query.clear();
            input.finder_selected = 0;
            input.finder_files =
                scan_files(&std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")));
            input.mode = Mode::FileFind;
        }
        AppAction::WelcomeEnterRecent => {
            input.recent_selected = 0;
            input.mode = Mode::RecentFiles;
        }
        AppAction::WelcomeOpenThemes => {
            input.theme_selected = 0;
            // Pre-select by theme_name id, falling back to cursor_bg match
            let themes = asat_config::builtin_themes();
            if let Some(idx) = themes
                .iter()
                .position(|t| t.id == config.theme_name)
                .or_else(|| {
                    themes
                        .iter()
                        .position(|t| t.config.cursor_bg == config.theme.cursor_bg)
                })
            {
                input.theme_selected = idx;
            }
            input.mode = Mode::ThemeManager;
        }
        AppAction::ThemeApply(idx) => {
            let themes = asat_config::builtin_themes();
            input.theme_selected = idx.min(themes.len().saturating_sub(1));
            if let Some(preset) = themes.get(input.theme_selected) {
                config.theme_name = preset.id.to_string();
                config.theme = preset.config.clone();
                match config.save() {
                    Ok(_) => set_status(
                        status,
                        format!("Theme \"{}\" applied and saved", preset.name),
                    ),
                    Err(e) => set_status(status, format!("Theme applied but couldn't save: {}", e)),
                }
            }
            input.mode = Mode::Welcome;
        }
        AppAction::ThemeManagerCancel => {
            input.mode = Mode::Welcome;
        }

        AppAction::OpenHelp => {
            input.help_tab = 0;
            input.help_scroll = 0;
            input.help_query.clear();
            input.mode = Mode::Help;
        }

        AppAction::OpenPluginManager => {
            input.plugin_selected = 0;
            input.plugin_show_output = false;
            input.mode = Mode::PluginManager;
        }

        AppAction::GotoDefinition => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            let col = input.cursor.col;
            if let CellValue::Formula(formula) = workbook.active().get_raw_value(row, col) {
                let formula = formula.clone();
                if let Some((target_sheet_name, target_row, target_col)) =
                    find_first_cell_ref(&formula)
                {
                    let target_sheet_idx = if let Some(name) = target_sheet_name {
                        workbook
                            .sheets
                            .iter()
                            .position(|s| s.name.eq_ignore_ascii_case(&name))
                            .unwrap_or(sheet_idx)
                    } else {
                        sheet_idx
                    };
                    input.save_position(sheet_idx);
                    workbook.active_sheet = target_sheet_idx;
                    input.cursor.row = target_row;
                    input.cursor.col = target_col;
                    input.scroll_to_cursor(visible_rows, visible_cols);
                } else {
                    set_status(status, "No cell reference in formula".to_string());
                }
            }
        }

        // ── Style copy / paste ──
        AppAction::YankStyle => {
            let (r, c) = (input.cursor.row, input.cursor.col);
            input.style_clipboard = workbook
                .active()
                .get_cell(r, c)
                .and_then(|cell| cell.style.clone());
            set_status(
                status,
                if input.style_clipboard.is_some() {
                    "Style copied".to_string()
                } else {
                    "Cell has no style to copy".to_string()
                },
            );
        }
        AppAction::PasteStyle => {
            if let Some(style) = input.style_clipboard.clone() {
                apply_style_sel(
                    workbook,
                    input,
                    undo,
                    status,
                    &move |s| *s = style.clone(),
                    "Style pasted",
                );
            } else {
                set_status(status, "No style in clipboard (use yS to copy)".to_string());
            }
        }

        // ── Formula reference selection ──
        AppAction::EnterFormulaSelect => {
            // Remember where we are so we can restore on cancel
            input.formula_origin = Some((input.cursor.row, input.cursor.col));
            input.mode = Mode::FormulaSelect { anchor: None };
        }
        AppAction::FormulaSelectStartRange => {
            // Mark the current cell as the range anchor
            if matches!(input.mode, Mode::FormulaSelect { anchor: None }) {
                input.mode = Mode::FormulaSelect {
                    anchor: Some((input.cursor.row, input.cursor.col)),
                };
            }
        }
        AppAction::FormulaSelectConfirm => {
            if let Mode::FormulaSelect { anchor } = input.mode.clone() {
                let ref_str = if let Some((ar, ac)) = anchor {
                    format!(
                        "{}:{}",
                        cell_address(ar, ac),
                        cell_address(input.cursor.row, input.cursor.col)
                    )
                } else {
                    cell_address(input.cursor.row, input.cursor.col)
                };
                // Insert the reference at the current edit cursor position
                input
                    .edit_buffer
                    .insert_str(input.edit_cursor_pos, &ref_str);
                input.edit_cursor_pos += ref_str.len();
            }
            // Restore cursor to the cell being edited and return to Insert
            if let Some((or, oc)) = input.formula_origin.take() {
                input.cursor.row = or;
                input.cursor.col = oc;
                input.scroll_to_cursor(visible_rows, visible_cols);
            }
            input.mode = Mode::Insert { replace: false };
        }
        AppAction::FormulaSelectCancel => {
            // Restore cursor to where we started
            if let Some((or, oc)) = input.formula_origin.take() {
                input.cursor.row = or;
                input.cursor.col = oc;
                input.scroll_to_cursor(visible_rows, visible_cols);
            }
            input.mode = Mode::Insert { replace: false };
        }
        AppAction::WelcomeOpenConfig => {
            open_config_in_editor();
            // Reload config after editor exits
            if let Ok(new_cfg) = asat_config::Config::load() {
                input.scroll_padding = new_cfg.scroll_padding;
            }
            set_status(
                status,
                "Config saved — restart for all changes to take effect".to_string(),
            );
            return ActionResult::ClearTerminal;
        }

        // ── File finder actions ──
        AppAction::FinderMoveUp => {
            if input.finder_selected > 0 {
                input.finder_selected -= 1;
            }
        }
        AppAction::FinderMoveDown => {
            let n = input.filtered_finder_files().len();
            if input.finder_selected + 1 < n {
                input.finder_selected += 1;
            }
        }
        AppAction::FinderOpen => {
            let path_str = input
                .filtered_finder_files()
                .get(input.finder_selected)
                .map(|s| s.to_string());
            if let Some(p) = path_str {
                let path = PathBuf::from(&p);
                match asat_io::load(&path) {
                    Ok(wb) => {
                        *workbook = wb;
                        push_recent(&p, &mut input.recent_files);
                        save_recent_files(&input.recent_files);
                        input.cursor = asat_input::Cursor::new();
                        input.viewport = asat_input::Viewport::default();
                        input.mode = Mode::Normal;
                        input.finder_query.clear();
                        plugins.push_event(PluginEvent::Open {
                            path: Some(p.clone()),
                        });
                        set_status(status, format!("Opened \"{}\"", p));
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }
        AppAction::FinderCancel => {
            input.finder_query.clear();
            input.finder_selected = 0;
            input.mode = Mode::Welcome;
        }

        // ── Recent files actions ──
        AppAction::RecentMoveUp => {
            if input.recent_selected > 0 {
                input.recent_selected -= 1;
            }
        }
        AppAction::RecentMoveDown => {
            if input.recent_selected + 1 < input.recent_files.len() {
                input.recent_selected += 1;
            }
        }
        AppAction::RecentOpen => {
            let path_str = input.recent_files.get(input.recent_selected).cloned();
            if let Some(p) = path_str {
                let path = PathBuf::from(&p);
                match asat_io::load(&path) {
                    Ok(wb) => {
                        *workbook = wb;
                        push_recent(&p, &mut input.recent_files);
                        save_recent_files(&input.recent_files);
                        input.cursor = asat_input::Cursor::new();
                        input.viewport = asat_input::Viewport::default();
                        input.mode = Mode::Normal;
                        set_status(status, format!("Opened \"{}\"", p));
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }
        AppAction::RecentRemove => {
            if !input.recent_files.is_empty() {
                input.recent_files.remove(input.recent_selected);
                save_recent_files(&input.recent_files);
                if input.recent_selected >= input.recent_files.len() && input.recent_selected > 0 {
                    input.recent_selected -= 1;
                }
                // If list is now empty, go back to welcome
                if input.recent_files.is_empty() {
                    input.mode = Mode::Welcome;
                }
            }
        }
        AppAction::RecentCancel => {
            input.recent_selected = 0;
            input.mode = Mode::Welcome;
        }

        // ── Row/col stubs now implemented ──
        AppAction::InsertRowBelow => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row + 1;
            let cmd = Box::new(asat_commands::InsertRow {
                sheet: sheet_idx,
                row,
            });
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
            }
        }
        AppAction::InsertColLeft => {
            let sheet_idx = workbook.active_sheet;
            let col = input.cursor.col;
            let cmd = Box::new(InsertCol {
                sheet: sheet_idx,
                col,
            });
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                set_status(status, "Column inserted".to_string());
            }
        }
        AppAction::InsertColRight => {
            let sheet_idx = workbook.active_sheet;
            let col = input.cursor.col + 1;
            let cmd = Box::new(InsertCol {
                sheet: sheet_idx,
                col,
            });
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                set_status(status, "Column inserted".to_string());
            }
        }
        AppAction::DeleteCurrentCol => {
            let sheet_idx = workbook.active_sheet;
            let col = input.cursor.col;
            let cmd = Box::new(DeleteCol::new(workbook, sheet_idx, col));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                set_status(status, "Column deleted".to_string());
            }
        }
        AppAction::SaveAs(path_str) => {
            let path = PathBuf::from(&path_str);
            match asat_io::save(workbook, &path) {
                Ok(_) => {
                    workbook.file_path = Some(path.clone());
                    workbook.dirty = false;
                    set_status(status, format!("\"{}\" written", path.display()));
                }
                Err(e) => set_status(status, format!("Save error: {}", e)),
            }
        }
        AppAction::OpenFile(path_str) => {
            let path = PathBuf::from(&path_str);
            match asat_io::load(&path) {
                Ok(wb) => {
                    *workbook = wb;
                    push_recent(&path_str, &mut input.recent_files);
                    save_recent_files(&input.recent_files);
                    input.cursor = asat_input::Cursor::new();
                    input.viewport = asat_input::Viewport::default();
                    input.mode = Mode::Normal;
                    set_status(status, format!("Opened \"{}\"", path_str));
                }
                Err(e) => set_status(status, format!("Error: {}", e)),
            }
        }

        // ── Macro recording / playback ──
        AppAction::StartRecording { register } => {
            input.recording_buffer.clear();
            input.macro_recording = Some(register);
            set_status(
                status,
                format!("Recording to \"{}\"... (q to stop)", register),
            );
        }
        AppAction::StopRecording => {
            if let Some(register) = input.macro_recording.take() {
                let count = input.recording_buffer.len();
                input
                    .macro_registers
                    .insert(register, input.recording_buffer.clone());
                input.last_macro_register = Some(register);
                input.recording_buffer.clear();
                set_status(
                    status,
                    format!("Recorded {} keys to \"{}\"", count, register),
                );
            }
        }
        AppAction::PlayMacro { register } => {
            if let Some(keys) = input.macro_registers.get(&register).cloned() {
                input.last_macro_register = Some(register);
                replay_macro(
                    keys,
                    workbook,
                    input,
                    undo,
                    status,
                    config,
                    plugins,
                    clipboard,
                    visible_rows,
                    visible_cols,
                    edit_count,
                );
            } else {
                set_status(status, format!("Register \"{}\" is empty", register));
            }
        }

        // ── Cell arithmetic ──
        AppAction::IncrementCell => {
            let sheet_idx = workbook.active_sheet;
            let (row, col) = (input.cursor.row, input.cursor.col);
            let val = workbook.active().get_value(row, col).clone();
            let new_val = match &val {
                CellValue::Number(n) => CellValue::Number(n + 1.0),
                CellValue::Empty => CellValue::Number(1.0),
                CellValue::Text(s) => {
                    if let Some(cycled) = cycle_date(s, 1) {
                        CellValue::Text(cycled)
                    } else if let Some(cycled) = cycle_text_sequence(s, 1) {
                        CellValue::Text(cycled)
                    } else {
                        val.clone()
                    }
                }
                _ => val.clone(),
            };
            let cmd = Box::new(SetCell::new(workbook, sheet_idx, row, col, new_val));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
            }
        }
        AppAction::DecrementCell => {
            let sheet_idx = workbook.active_sheet;
            let (row, col) = (input.cursor.row, input.cursor.col);
            let val = workbook.active().get_value(row, col).clone();
            let new_val = match &val {
                CellValue::Number(n) => CellValue::Number(n - 1.0),
                CellValue::Empty => CellValue::Number(-1.0),
                CellValue::Text(s) => {
                    if let Some(cycled) = cycle_date(s, -1) {
                        CellValue::Text(cycled)
                    } else if let Some(cycled) = cycle_text_sequence(s, -1) {
                        CellValue::Text(cycled)
                    } else {
                        val.clone()
                    }
                }
                _ => val.clone(),
            };
            let cmd = Box::new(SetCell::new(workbook, sheet_idx, row, col, new_val));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
            }
        }

        // ── Column jump ──
        AppAction::JumpToCol { col } => {
            input.cursor.col = col;
            input.scroll_to_cursor(visible_rows, visible_cols);
        }

        // ── Join cell below ──
        AppAction::JoinCellBelow => {
            let sheet_idx = workbook.active_sheet;
            let (row, col) = (input.cursor.row, input.cursor.col);
            let above = workbook.active().get_value(row, col).display();
            let below = workbook.active().get_value(row + 1, col).display();
            let joined = if above.is_empty() {
                below.clone()
            } else if below.is_empty() {
                above.clone()
            } else {
                format!("{} {}", above, below)
            };
            let set_above = Box::new(SetCell::new(
                workbook,
                sheet_idx,
                row,
                col,
                asat_input::parse_cell_value(&joined),
            ));
            let clear_below = Box::new(SetCell::new(
                workbook,
                sheet_idx,
                row + 1,
                col,
                CellValue::Empty,
            ));
            let grouped = Box::new(asat_commands::GroupedCommand {
                description: "join cells".to_string(),
                commands: vec![set_above, clear_below],
            });
            if let Err(e) = grouped.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(grouped);
            }
        }

        // ── Visual clear + insert ──
        AppAction::VisualClearAndInsert => {
            // The DeleteCellRange was already dispatched; just enter insert mode at cursor
            input.edit_buffer.clear();
            input.edit_cursor_pos = 0;
            input.mode = Mode::Insert { replace: false };
        }

        // ── Clipboard paste in insert mode ──
        AppAction::PasteFromClipboard => {
            if let Some(cb) = clipboard.as_mut() {
                if let Ok(text) = cb.get_text() {
                    // Insert at current edit cursor position
                    input.edit_buffer.insert_str(input.edit_cursor_pos, &text);
                    input.edit_cursor_pos += text.len();
                }
            }
        }

        // ── Fill Down ──
        AppAction::FillDown {
            anchor_row,
            col_start,
            col_end,
            row_end,
        } => {
            let sheet_idx = workbook.active_sheet;
            let col_end_c = col_end.min(workbook.active().max_col());
            let row_end_c = row_end;
            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for c in col_start..=col_end_c {
                let anchor_val = workbook.active().get_raw_value(anchor_row, c).clone();
                let anchor_style = workbook
                    .active()
                    .get_cell(anchor_row, c)
                    .and_then(|cell| cell.style.clone());
                for r in (anchor_row + 1)..=row_end_c {
                    let old_cell = workbook.active().get_cell(r, c);
                    let old_val = old_cell
                        .map(|c| c.value.clone())
                        .unwrap_or(CellValue::Empty);
                    let old_style = old_cell.and_then(|c| c.style.clone());
                    cmds.push(Box::new(SetCell {
                        sheet: sheet_idx,
                        row: r,
                        col: c,
                        old_value: old_val,
                        new_value: anchor_val.clone(),
                        old_style,
                        new_style: anchor_style.clone(),
                    }));
                }
            }
            if !cmds.is_empty() {
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: "fill down".to_string(),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, "Fill down".to_string());
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }

        // ── Fill Right ──
        AppAction::FillRight {
            anchor_col,
            row_start,
            row_end,
            col_end,
        } => {
            let sheet_idx = workbook.active_sheet;
            let row_end_c = row_end;
            let col_end_c = col_end;
            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for r in row_start..=row_end_c {
                let anchor_val = workbook.active().get_raw_value(r, anchor_col).clone();
                let anchor_style = workbook
                    .active()
                    .get_cell(r, anchor_col)
                    .and_then(|cell| cell.style.clone());
                for c in (anchor_col + 1)..=col_end_c {
                    let old_cell = workbook.active().get_cell(r, c);
                    let old_val = old_cell
                        .map(|x| x.value.clone())
                        .unwrap_or(CellValue::Empty);
                    let old_style = old_cell.and_then(|x| x.style.clone());
                    cmds.push(Box::new(SetCell {
                        sheet: sheet_idx,
                        row: r,
                        col: c,
                        old_value: old_val,
                        new_value: anchor_val.clone(),
                        old_style,
                        new_style: anchor_style.clone(),
                    }));
                }
            }
            if !cmds.is_empty() {
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: "fill right".to_string(),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, "Fill right".to_string());
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }

        // ── Auto-Fill Series Down ──
        AppAction::AutoFillDown {
            row_start,
            row_end,
            col_start,
            col_end,
        } => {
            let sheet_idx = workbook.active_sheet;
            let col_end_c = col_end.min(workbook.active().max_col());
            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for c in col_start..=col_end_c {
                // Collect seed: up to first 2 rows
                let seed_end = (row_start + 1).min(row_end);
                let seed: Vec<CellValue> = (row_start..=seed_end)
                    .map(|r| workbook.active().get_raw_value(r, c).clone())
                    .collect();
                let seed_len = seed.len() as u32;
                if row_end < row_start + seed_len {
                    // Selection is entirely within the seed — nothing to fill
                    continue;
                }
                let filler = auto_fill_series(&seed);
                for r in (row_start + seed_len)..=row_end {
                    let idx = (r - row_start - seed_len) as usize;
                    let new_value = filler(idx);
                    let old_cell = workbook.active().get_cell(r, c);
                    let old_val = old_cell
                        .map(|x| x.value.clone())
                        .unwrap_or(CellValue::Empty);
                    let old_style = old_cell.and_then(|x| x.style.clone());
                    cmds.push(Box::new(SetCell {
                        sheet: sheet_idx,
                        row: r,
                        col: c,
                        old_value: old_val,
                        new_value,
                        old_style,
                        new_style: None,
                    }));
                }
            }
            if !cmds.is_empty() {
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: "auto-fill series down".to_string(),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, "Auto-fill series down".to_string());
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }

        // ── Auto-Fill Series Right ──
        AppAction::AutoFillRight {
            row_start,
            row_end,
            col_start,
            col_end,
        } => {
            let sheet_idx = workbook.active_sheet;
            let row_end_c = row_end;
            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for r in row_start..=row_end_c {
                // Collect seed: up to first 2 columns
                let seed_end = (col_start + 1).min(col_end);
                let seed: Vec<CellValue> = (col_start..=seed_end)
                    .map(|c| workbook.active().get_raw_value(r, c).clone())
                    .collect();
                let seed_len = seed.len() as u32;
                if col_end < col_start + seed_len {
                    continue;
                }
                let filler = auto_fill_series(&seed);
                for c in (col_start + seed_len)..=col_end {
                    let idx = (c - col_start - seed_len) as usize;
                    let new_value = filler(idx);
                    let old_cell = workbook.active().get_cell(r, c);
                    let old_val = old_cell
                        .map(|x| x.value.clone())
                        .unwrap_or(CellValue::Empty);
                    let old_style = old_cell.and_then(|x| x.style.clone());
                    cmds.push(Box::new(SetCell {
                        sheet: sheet_idx,
                        row: r,
                        col: c,
                        old_value: old_val,
                        new_value,
                        old_style,
                        new_style: None,
                    }));
                }
            }
            if !cmds.is_empty() {
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: "auto-fill series right".to_string(),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, "Auto-fill series right".to_string());
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }

        // ── Smart Auto-Fill (direction from anchor→cursor) ──
        AppAction::AutoFill {
            anchor_row,
            anchor_col,
            cursor_row,
            cursor_col,
        } => {
            let sheet_idx = workbook.active_sheet;
            let row_min = anchor_row.min(cursor_row);
            let row_max = anchor_row.max(cursor_row);
            let col_min = anchor_col.min(cursor_col);
            let col_max = anchor_col.max(cursor_col);
            let row_span = row_max - row_min;
            let col_span = col_max - col_min;

            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();

            if row_span >= col_span {
                // Primary axis: vertical
                let fill_down = anchor_row <= cursor_row;
                for c in col_min..=col_max {
                    let seed: Vec<CellValue> = if fill_down {
                        // Seed from top rows
                        let seed_end = (row_min + 1).min(row_max);
                        (row_min..=seed_end)
                            .map(|r| workbook.active().get_raw_value(r, c).clone())
                            .collect()
                    } else {
                        // Seed from bottom rows (reversed so series goes upward)
                        let seed_start = row_max.saturating_sub(1).max(row_min);
                        (seed_start..=row_max)
                            .rev()
                            .map(|r| workbook.active().get_raw_value(r, c).clone())
                            .collect()
                    };
                    let seed_len = seed.len() as u32;
                    if row_span < seed_len {
                        continue;
                    }
                    let filler = auto_fill_series(&seed);
                    if fill_down {
                        for r in (row_min + seed_len)..=row_max {
                            let idx = (r - row_min - seed_len) as usize;
                            let new_value = filler(idx);
                            let old_cell = workbook.active().get_cell(r, c);
                            let old_val = old_cell
                                .map(|x| x.value.clone())
                                .unwrap_or(CellValue::Empty);
                            let old_style = old_cell.and_then(|x| x.style.clone());
                            cmds.push(Box::new(SetCell {
                                sheet: sheet_idx,
                                row: r,
                                col: c,
                                old_value: old_val,
                                new_value,
                                old_style,
                                new_style: None,
                            }));
                        }
                    } else {
                        for (i, r) in (row_min..row_max.saturating_sub(seed_len) + 1)
                            .rev()
                            .enumerate()
                        {
                            let new_value = filler(i);
                            let old_cell = workbook.active().get_cell(r, c);
                            let old_val = old_cell
                                .map(|x| x.value.clone())
                                .unwrap_or(CellValue::Empty);
                            let old_style = old_cell.and_then(|x| x.style.clone());
                            cmds.push(Box::new(SetCell {
                                sheet: sheet_idx,
                                row: r,
                                col: c,
                                old_value: old_val,
                                new_value,
                                old_style,
                                new_style: None,
                            }));
                        }
                    }
                }
            } else {
                // Primary axis: horizontal
                let fill_right = anchor_col <= cursor_col;
                for r in row_min..=row_max {
                    let seed: Vec<CellValue> = if fill_right {
                        let seed_end = (col_min + 1).min(col_max);
                        (col_min..=seed_end)
                            .map(|c| workbook.active().get_raw_value(r, c).clone())
                            .collect()
                    } else {
                        let seed_start = col_max.saturating_sub(1).max(col_min);
                        (seed_start..=col_max)
                            .rev()
                            .map(|c| workbook.active().get_raw_value(r, c).clone())
                            .collect()
                    };
                    let seed_len = seed.len() as u32;
                    if col_span < seed_len {
                        continue;
                    }
                    let filler = auto_fill_series(&seed);
                    if fill_right {
                        for c in (col_min + seed_len)..=col_max {
                            let idx = (c - col_min - seed_len) as usize;
                            let new_value = filler(idx);
                            let old_cell = workbook.active().get_cell(r, c);
                            let old_val = old_cell
                                .map(|x| x.value.clone())
                                .unwrap_or(CellValue::Empty);
                            let old_style = old_cell.and_then(|x| x.style.clone());
                            cmds.push(Box::new(SetCell {
                                sheet: sheet_idx,
                                row: r,
                                col: c,
                                old_value: old_val,
                                new_value,
                                old_style,
                                new_style: None,
                            }));
                        }
                    } else {
                        for (i, c) in (col_min..col_max.saturating_sub(seed_len) + 1)
                            .rev()
                            .enumerate()
                        {
                            let new_value = filler(i);
                            let old_cell = workbook.active().get_cell(r, c);
                            let old_val = old_cell
                                .map(|x| x.value.clone())
                                .unwrap_or(CellValue::Empty);
                            let old_style = old_cell.and_then(|x| x.style.clone());
                            cmds.push(Box::new(SetCell {
                                sheet: sheet_idx,
                                row: r,
                                col: c,
                                old_value: old_val,
                                new_value,
                                old_style,
                                new_style: None,
                            }));
                        }
                    }
                }
            }

            if !cmds.is_empty() {
                let dir = if row_span >= col_span {
                    if anchor_row <= cursor_row {
                        "down"
                    } else {
                        "up"
                    }
                } else if anchor_col <= cursor_col {
                    "right"
                } else {
                    "left"
                };
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: format!("auto-fill series {}", dir),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, format!("Auto-fill series {}", dir));
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }

        // ── Goto cell ──
        AppAction::GotoCell(addr) => {
            if let Some((r, c)) = parse_cell_address(&addr) {
                input.save_position(workbook.active_sheet);
                input.cursor.row = r;
                input.cursor.col = c;
                input.scroll_to_cursor(visible_rows, visible_cols);
                set_status(status, format!("→ {}", addr.to_uppercase()));
            } else {
                set_status(status, format!("Invalid cell address: {}", addr));
            }
        }

        // ── Cell note ──
        AppAction::SetNote { row, col, text } => {
            let sheet = workbook.active_mut();
            if text.is_empty() {
                sheet.notes.remove(&(row, col));
                set_status(status, "Note cleared".to_string());
            } else {
                sheet.notes.insert((row, col), text);
                set_status(status, "Note set".to_string());
            }
            workbook.dirty = true;
        }

        // ── Insert SUM formula ──
        AppAction::ToggleWrap => {
            let cur = workbook
                .active()
                .get_cell(input.cursor.row, input.cursor.col)
                .and_then(|c| c.style.as_ref())
                .map(|s| s.wrap)
                .unwrap_or(false);
            apply_style_sel(
                workbook,
                input,
                undo,
                status,
                &move |s| s.wrap = !cur,
                if cur { "Wrap off" } else { "Wrap on" },
            );
        }

        AppAction::MergeCells {
            row_start,
            col_start,
            row_end,
            col_end,
        } => {
            let sheet_idx = workbook.active_sheet;
            let cmd = Box::new(MergeCells::new(
                workbook, sheet_idx, row_start, col_start, row_end, col_end,
            ));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                input.cursor.row = row_start;
                input.cursor.col = col_start;
                input.scroll_to_cursor(visible_rows, visible_cols);
                set_status(
                    status,
                    format!(
                        "Merged {}:{} → {}:{}",
                        cell_address(row_start, col_start),
                        cell_address(row_end, col_end),
                        cell_address(row_start, col_start),
                        cell_address(row_end, col_end),
                    ),
                );
            }
        }

        AppAction::UnmergeCells { row, col } => {
            let sheet_idx = workbook.active_sheet;
            let cmd = Box::new(UnmergeCells::new(workbook, sheet_idx, row, col));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                set_status(status, format!("Unmerged {}", cell_address(row, col)));
            }
        }

        // ── Repeat last change (.) ──
        AppAction::RepeatLastChange => {
            let keys = input.last_edit_keys.clone();
            // Temporarily disable recording so replay doesn't overwrite last_edit_keys
            let was_recording = input.recording_edit;
            input.recording_edit = false;
            replay_macro(
                keys,
                workbook,
                input,
                undo,
                status,
                config,
                plugins,
                clipboard,
                visible_rows,
                visible_cols,
                edit_count,
            );
            input.recording_edit = was_recording;
        }

        // ── Change inner text object (ci"/ci'/ etc.) ──
        AppAction::ChangeInner { open, close } => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            let col = input.cursor.col;
            let cell_str = workbook.active().get_value(row, col).display();
            // Find innermost pair of open..close delimiters
            let open_pos = cell_str.find(open);
            let close_pos = cell_str.rfind(close);
            if let (Some(op), Some(cp)) = (open_pos, close_pos) {
                // Ensure the closing delimiter is after the opening one
                // (or at least not the same character position for symmetric delimiters)
                let inner_start = op + open.len_utf8();
                if cp > op || (open == close && cp >= op) {
                    let inner_end = cp;
                    let inner_end = if inner_end > inner_start {
                        inner_end
                    } else {
                        inner_start
                    };
                    let inner = cell_str[inner_start..inner_end].to_string();
                    let prefix = cell_str[..inner_start].to_string();
                    let suffix = cell_str[inner_end..].to_string();
                    // Clear the cell first (for undo)
                    let cmd = Box::new(SetCell::new(
                        workbook,
                        sheet_idx,
                        row,
                        col,
                        CellValue::Empty,
                    ));
                    if let Err(e) = cmd.execute(workbook) {
                        set_status(status, format!("Error: {}", e));
                        return ActionResult::Continue;
                    }
                    undo.push(cmd);
                    // Set up edit buffer with inner content; store prefix/suffix
                    input.edit_buffer = inner;
                    input.edit_cursor_pos = input.edit_buffer.len();
                    input.ci_prefix = prefix;
                    input.ci_suffix = suffix;
                    input.mode = Mode::Insert { replace: false };
                } else {
                    // No valid delimiters found — just enter insert mode normally
                    input.edit_buffer = cell_str;
                    input.edit_cursor_pos = input.edit_buffer.len();
                    input.mode = Mode::Insert { replace: false };
                }
            } else {
                // Delimiter not found — enter insert mode normally
                input.edit_buffer = cell_str;
                input.edit_cursor_pos = input.edit_buffer.len();
                input.mode = Mode::Insert { replace: false };
            }
        }

        AppAction::InsertSumBelow {
            row_start,
            col_start,
            row_end,
            col_end,
        } => {
            let sheet = workbook.active();
            let row_end_c = row_end.min(sheet.max_row());
            let col_end_c = col_end.min(sheet.max_col());
            let range_str = format!(
                "{}:{}",
                cell_address(row_start, col_start),
                cell_address(row_end_c, col_end_c)
            );
            let formula = format!("SUM({})", range_str);
            // Place =SUM one row below the selection, at the leftmost column of the selection
            let dest_row = row_end_c + 1;
            let dest_col = col_start;
            let sheet_idx = workbook.active_sheet;
            let cmd = Box::new(SetCell::new(
                workbook,
                sheet_idx,
                dest_row,
                dest_col,
                CellValue::Formula(formula.clone()),
            ));
            if let Err(e) = cmd.execute(workbook) {
                set_status(status, format!("Error: {}", e));
            } else {
                undo.push(cmd);
                input.cursor.row = dest_row;
                input.cursor.col = dest_col;
                input.scroll_to_cursor(visible_rows, visible_cols);
                set_status(
                    status,
                    format!("={}  →  {}", formula, cell_address(dest_row, dest_col)),
                );
            }
        }
    }

    ActionResult::Continue
}

fn do_paste(
    workbook: &mut Workbook,
    input: &mut InputState,
    undo: &mut UndoStack,
    status: &mut Option<(String, std::time::Instant)>,
    is_after: bool,
) {
    let reg = input.registers.get(None).clone();
    if reg.cells.is_empty() {
        set_status(status, "Nothing to paste".to_string());
        return;
    }
    let sheet_idx = workbook.active_sheet;
    let (start_row, start_col) = if reg.is_line {
        let row = if is_after {
            input.cursor.row + 1
        } else {
            input.cursor.row
        };
        (row, 0u32)
    } else {
        (input.cursor.row, input.cursor.col)
    };
    // Compute row/col delta from yank source to paste destination for formula adjustment
    let d_row = start_row as i64 - reg.source_row as i64;
    let d_col = start_col as i64 - reg.source_col as i64;
    let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
    for (dr, row_vals) in reg.cells.iter().enumerate() {
        for (dc, val) in row_vals.iter().enumerate() {
            let r = start_row + dr as u32;
            let c = start_col + dc as u32;
            // Adjust formula references for the paste offset
            let adjusted = if let CellValue::Formula(f) = val {
                CellValue::Formula(asat_formula::adjust_formula_refs(f, d_row, d_col))
            } else {
                val.clone()
            };
            let mut cmd = Box::new(SetCell::new(workbook, sheet_idx, r, c, adjusted));
            // Apply yanked style if available
            if let Some(sty) = reg.styles.get(dr).and_then(|row| row.get(dc)) {
                cmd.new_style = sty.clone();
            }
            cmds.push(cmd);
        }
    }
    let grouped = Box::new(asat_commands::GroupedCommand {
        description: "paste".to_string(),
        commands: cmds,
    });
    if let Err(e) = grouped.execute(workbook) {
        set_status(status, format!("Paste error: {}", e));
    } else {
        undo.push(grouped);
        set_status(status, "Pasted".to_string());
    }
}

pub(crate) fn replay_macro(
    keys: Vec<KeyEvent>,
    workbook: &mut Workbook,
    input: &mut InputState,
    undo: &mut UndoStack,
    status: &mut Option<(String, std::time::Instant)>,
    config: &mut Config,
    plugins: &mut PluginManager,
    clipboard: &mut Option<Clipboard>,
    visible_rows: u32,
    visible_cols: u32,
    edit_count: &mut u32,
) {
    for key in keys {
        let actions = input.handle_key(key, workbook);
        for action in actions {
            if matches!(
                action,
                AppAction::PlayMacro { .. }
                    | AppAction::StartRecording { .. }
                    | AppAction::StopRecording
            ) {
                continue;
            }
            match process_action(
                action,
                workbook,
                input,
                undo,
                status,
                config,
                plugins,
                clipboard,
                visible_rows,
                visible_cols,
                edit_count,
            ) {
                ActionResult::Continue | ActionResult::ClearTerminal => {}
                ActionResult::Quit | ActionResult::ForceQuit => return,
            }
        }
    }
}
