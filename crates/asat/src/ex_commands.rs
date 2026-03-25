use regex::Regex;

use asat_commands::{
    Command, DeleteCol, InsertCol, MergeCells, RemoveSheet, SetCell, UndoStack, UnmergeCells,
};
use asat_config::Config;
use asat_core::{cell_address, CellStyle, CellValue, Workbook};
use asat_input::{InputState, Mode};
use asat_plugins::PluginManager;

use crate::app::{
    apply_style_sel, compare_cell_values, filter_row_matches, parse_cell_address,
    parse_range_address, parse_range_address_cf, parse_color_arg, set_status, style_range,
    ActionResult,
};

pub(crate) fn handle_ex_command(
    cmd: &str,
    workbook: &mut Workbook,
    input: &mut InputState,
    undo: &mut UndoStack,
    status: &mut Option<(String, std::time::Instant)>,
    config: &mut Config,
    plugins: &mut PluginManager,
) -> ActionResult {
    let parts: Vec<&str> = cmd.trim().splitn(2, ' ').collect();
    let verb = parts[0];
    let arg = parts.get(1).copied().unwrap_or("").trim();

    match verb {
        "q" | "quit" => {
            if workbook.dirty {
                set_status(
                    status,
                    "Unsaved changes. Use :q! to force quit or :wq to save and quit".to_string(),
                );
            } else {
                return ActionResult::Quit;
            }
        }
        "q!" | "quit!" => {
            return ActionResult::ForceQuit;
        }
        "w" | "write" => {
            let path = if arg.is_empty() {
                workbook.file_path.clone()
            } else {
                Some(std::path::PathBuf::from(arg))
            };
            if let Some(path) = path {
                match asat_io::save(workbook, &path) {
                    Ok(_) => {
                        workbook.file_path = Some(path.clone());
                        workbook.dirty = false;
                        set_status(status, format!("\"{}\" written", path.display()));
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            } else {
                set_status(status, "No file name".to_string());
            }
        }
        "wq" | "x" => {
            let path = if arg.is_empty() {
                workbook.file_path.clone()
            } else {
                Some(std::path::PathBuf::from(arg))
            };
            if let Some(path) = path {
                match asat_io::save(workbook, &path) {
                    Ok(_) => {
                        workbook.dirty = false;
                        return ActionResult::Quit;
                    }
                    Err(e) => set_status(status, format!("Save error: {}", e)),
                }
            } else {
                return ActionResult::Quit;
            }
        }
        "e" | "edit" => {
            if !arg.is_empty() {
                let path = std::path::PathBuf::from(arg);
                match asat_io::load(&path) {
                    Ok(wb) => {
                        *workbook = wb;
                        input.cursor = asat_input::Cursor::new();
                        input.viewport = asat_input::Viewport::default();
                        set_status(status, format!("Opened \"{}\"", arg));
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }
        "tabnew" | "tabedit" => {
            let name = if arg.is_empty() {
                format!("Sheet{}", workbook.sheets.len() + 1)
            } else {
                arg.to_string()
            };
            workbook.active_sheet = workbook.add_sheet(name);
        }
        "tabclose" => {
            if workbook.sheets.len() > 1 {
                match RemoveSheet::new(workbook, workbook.active_sheet) {
                    Ok(cmd) => {
                        let name = workbook.sheets[workbook.active_sheet].name.clone();
                        if let Err(e) = cmd.execute(workbook) {
                            set_status(status, format!("Error: {}", e));
                        } else {
                            undo.push(Box::new(cmd));
                            set_status(status, format!("Closed sheet \"{}\"", name));
                        }
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            } else {
                set_status(status, "Cannot close last sheet".to_string());
            }
        }
        // ── Column operations ──
        "ic" | "insertcol" => {
            let sheet_idx = workbook.active_sheet;
            let col = input.cursor.col;
            let cmd = Box::new(InsertCol {
                sheet: sheet_idx,
                col,
            });
            match cmd.execute(workbook) {
                Ok(_) => {
                    undo.push(cmd);
                    set_status(status, "Column inserted".to_string());
                }
                Err(e) => set_status(status, format!("Error: {}", e)),
            }
        }
        "icr" | "insertcolright" => {
            let sheet_idx = workbook.active_sheet;
            let col = input.cursor.col + 1;
            let cmd = Box::new(InsertCol {
                sheet: sheet_idx,
                col,
            });
            match cmd.execute(workbook) {
                Ok(_) => {
                    undo.push(cmd);
                    set_status(status, "Column inserted".to_string());
                }
                Err(e) => set_status(status, format!("Error: {}", e)),
            }
        }
        "dc" | "deletecol" => {
            let sheet_idx = workbook.active_sheet;
            let col = input.cursor.col;
            let cmd = Box::new(DeleteCol::new(workbook, sheet_idx, col));
            match cmd.execute(workbook) {
                Ok(_) => {
                    undo.push(cmd);
                    set_status(status, "Column deleted".to_string());
                }
                Err(e) => set_status(status, format!("Error: {}", e)),
            }
        }
        // ── Row operations via ex-command ──
        "ir" | "insertrow" => {
            let sheet_idx = workbook.active_sheet;
            let row = if arg.is_empty() {
                input.cursor.row
            } else {
                arg.parse::<u32>()
                    .unwrap_or(input.cursor.row + 1)
                    .saturating_sub(1)
            };
            let cmd = Box::new(asat_commands::InsertRow {
                sheet: sheet_idx,
                row,
            });
            match cmd.execute(workbook) {
                Ok(_) => {
                    undo.push(cmd);
                    set_status(status, "Row inserted".to_string());
                }
                Err(e) => set_status(status, format!("Error: {}", e)),
            }
        }
        "dr" | "deleterow" => {
            let sheet_idx = workbook.active_sheet;
            let row = if arg.is_empty() {
                input.cursor.row
            } else {
                arg.parse::<u32>()
                    .unwrap_or(input.cursor.row + 1)
                    .saturating_sub(1)
            };
            let cmd = Box::new(asat_commands::DeleteRow::new(workbook, sheet_idx, row));
            match cmd.execute(workbook) {
                Ok(_) => {
                    undo.push(cmd);
                    set_status(status, "Row deleted".to_string());
                }
                Err(e) => set_status(status, format!("Error: {}", e)),
            }
        }
        // ── Cell styles ──
        "bold" | "b" => {
            let cur = workbook
                .active()
                .get_cell(input.cursor.row, input.cursor.col)
                .and_then(|c| c.style.as_ref())
                .map(|s| s.bold)
                .unwrap_or(false);
            apply_style_sel(
                workbook,
                input,
                undo,
                status,
                &move |s| s.bold = !cur,
                if cur { "Bold off" } else { "Bold on" },
            );
        }
        "italic" | "it" => {
            let cur = workbook
                .active()
                .get_cell(input.cursor.row, input.cursor.col)
                .and_then(|c| c.style.as_ref())
                .map(|s| s.italic)
                .unwrap_or(false);
            apply_style_sel(
                workbook,
                input,
                undo,
                status,
                &move |s| s.italic = !cur,
                if cur { "Italic off" } else { "Italic on" },
            );
        }
        "underline" | "ul" => {
            if arg == "full" {
                let cur = workbook
                    .active()
                    .get_cell(input.cursor.row, input.cursor.col)
                    .and_then(|c| c.style.as_ref())
                    .map(|s| s.underline_full)
                    .unwrap_or(false);
                apply_style_sel(
                    workbook,
                    input,
                    undo,
                    status,
                    &move |s| s.underline_full = !cur,
                    if cur {
                        "Underline full off"
                    } else {
                        "Underline full on"
                    },
                );
            } else {
                let cur = workbook
                    .active()
                    .get_cell(input.cursor.row, input.cursor.col)
                    .and_then(|c| c.style.as_ref())
                    .map(|s| s.underline)
                    .unwrap_or(false);
                apply_style_sel(
                    workbook,
                    input,
                    undo,
                    status,
                    &move |s| s.underline = !cur,
                    if cur { "Underline off" } else { "Underline on" },
                );
            }
        }
        "strike" | "strikethrough" | "st" => {
            let cur = workbook
                .active()
                .get_cell(input.cursor.row, input.cursor.col)
                .and_then(|c| c.style.as_ref())
                .map(|s| s.strikethrough)
                .unwrap_or(false);
            apply_style_sel(
                workbook,
                input,
                undo,
                status,
                &move |s| s.strikethrough = !cur,
                if cur {
                    "Strikethrough off"
                } else {
                    "Strikethrough on"
                },
            );
        }
        "wrap" | "ww" => {
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
        "fg" | "color" => {
            if arg.is_empty() {
                set_status(
                    status,
                    "Usage: :fg <#rrggbb or name>  e.g. :fg red  :fg #ff8800".to_string(),
                );
            } else if let Some(col) = parse_color_arg(arg) {
                apply_style_sel(
                    workbook,
                    input,
                    undo,
                    status,
                    &move |s| s.fg = Some(col),
                    &format!("Foreground: {}", arg),
                );
            } else {
                set_status(
                    status,
                    format!(
                        "Unknown color: \"{}\"  (use hex #rrggbb or a name like red, blue…)",
                        arg
                    ),
                );
            }
        }
        "bg" | "bgcolor" => {
            if arg.is_empty() {
                set_status(
                    status,
                    "Usage: :bg <#rrggbb or name>  e.g. :bg yellow  :bg #003366".to_string(),
                );
            } else if let Some(col) = parse_color_arg(arg) {
                apply_style_sel(
                    workbook,
                    input,
                    undo,
                    status,
                    &move |s| s.bg = Some(col),
                    &format!("Background: {}", arg),
                );
            } else {
                set_status(
                    status,
                    format!(
                        "Unknown color: \"{}\"  (use hex #rrggbb or a name like red, blue…)",
                        arg
                    ),
                );
            }
        }
        "hl" | "highlight" => {
            if arg.is_empty() {
                // No argument → clear highlight colours
                apply_style_sel(
                    workbook,
                    input,
                    undo,
                    status,
                    &|s| {
                        s.bg = None;
                        s.fg = None;
                    },
                    "Highlight cleared",
                );
            } else if let Some(bg) = parse_color_arg(arg) {
                // Auto-pick a contrasting text colour
                let fg = if bg.is_dark() {
                    asat_core::Color::rgb(240, 240, 240)
                } else {
                    asat_core::Color::rgb(20, 20, 20)
                };
                apply_style_sel(
                    workbook,
                    input,
                    undo,
                    status,
                    &move |s| {
                        s.bg = Some(bg);
                        s.fg = Some(fg);
                    },
                    &format!("Highlight: {}", arg),
                );
            } else {
                set_status(status, format!("Unknown color: \"{}\"", arg));
            }
        }
        "align" | "al" => {
            use asat_core::Alignment;
            let al = match arg.to_ascii_lowercase().as_str() {
                "l" | "left" => Alignment::Left,
                "c" | "center" | "centre" => Alignment::Center,
                "r" | "right" => Alignment::Right,
                "d" | "default" | "" => Alignment::Default,
                _ => {
                    set_status(
                        status,
                        "Usage: :align l/c/r  (left / center / right)".to_string(),
                    );
                    return ActionResult::Continue;
                }
            };
            apply_style_sel(
                workbook,
                input,
                undo,
                status,
                &move |s| s.align = al,
                &format!("Align: {:?}", al),
            );
        }
        "fmt" | "format" => {
            use asat_core::NumberFormat;
            let fmt: Option<NumberFormat> = match arg.trim().to_ascii_lowercase().as_str() {
                "none" | "clear" | "general" | "" => None, // clear format
                "int" | "integer" => Some(NumberFormat::Integer),
                "%" | "pct" | "percent" | "percentage" => Some(NumberFormat::Percentage(1)),
                "%0" => Some(NumberFormat::Percentage(0)),
                "%2" => Some(NumberFormat::Percentage(2)),
                "$" | "usd" => Some(NumberFormat::Currency("$".to_string())),
                "€" | "eur" => Some(NumberFormat::Currency("€".to_string())),
                "£" | "gbp" => Some(NumberFormat::Currency("£".to_string())),
                "¥" | "jpy" => Some(NumberFormat::Currency("¥".to_string())),
                "date" => Some(NumberFormat::Date(String::new())),
                "datetime" | "dt" | "now" => Some(NumberFormat::DateTime),
                "#,##0" | "thousands" | "t" => Some(NumberFormat::Thousands),
                "#,##0.00" | "t2" | "thousands2" => Some(NumberFormat::ThousandsDecimals(2)),
                other => {
                    // Try to detect decimal spec like "0.00" or "2"
                    if let Ok(d) = other.parse::<u8>() {
                        Some(NumberFormat::Decimal(d))
                    } else {
                        // Count decimal places from "0.000" pattern
                        if let Some(dot) = other.find('.') {
                            let decimals = other.len() - dot - 1;
                            Some(NumberFormat::Decimal(decimals as u8))
                        } else {
                            set_status(status, format!(
                                "Unknown format: \"{}\"  Try: %, $, int, 2, 0.00, #,##0, date, none", arg));
                            return ActionResult::Continue;
                        }
                    }
                }
            };
            let fmt_clone = fmt.clone();
            let label = fmt
                .as_ref()
                .map(|f| format!("{:?}", f))
                .unwrap_or_else(|| "General".to_string());
            apply_style_sel(
                workbook,
                input,
                undo,
                status,
                &move |s| s.format = fmt_clone.clone(),
                &format!("Format: {}", label),
            );
        }
        "copystyle" | "ys" => {
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
        "pastestyle" | "ps" => {
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
                set_status(
                    status,
                    "No style in clipboard  (use :copystyle or yS first)".to_string(),
                );
            }
        }
        "cs" | "clearstyle" => {
            // Works on visual selection too
            let (row_start, col_start, row_end, col_end) = style_range(workbook, input);
            let sheet_idx = workbook.active_sheet;
            if row_start == row_end && col_start == col_end {
                let r = row_start;
                let c = col_start;
                let old_cell = workbook.sheet(sheet_idx).and_then(|s| s.get_cell(r, c));
                let old_value = old_cell
                    .map(|c| c.value.clone())
                    .unwrap_or(CellValue::Empty);
                let old_style = old_cell.and_then(|c| c.style.clone());
                if old_style.is_none() {
                    set_status(status, "Cell has no style to clear".to_string());
                } else {
                    let cmd = Box::new(SetCell {
                        sheet: sheet_idx,
                        row: r,
                        col: c,
                        old_value: old_value.clone(),
                        new_value: old_value,
                        old_style,
                        new_style: None,
                    });
                    match cmd.execute(workbook) {
                        Ok(_) => {
                            undo.push(cmd);
                            set_status(status, "Style cleared".to_string());
                        }
                        Err(e) => set_status(status, format!("Error: {}", e)),
                    }
                }
            } else {
                // Range: clear styles by applying an empty CellStyle that looks like "no style"
                // Use apply_style_sel to handle the range, but we need to clear entirely —
                // set a flag and post-process to None
                apply_style_sel(
                    workbook,
                    input,
                    undo,
                    status,
                    &|s| *s = CellStyle::default(),
                    "Styles cleared",
                );
            }
        }
        // ── Column/row sizing via ex-command ──
        "cw" | "colwidth" => {
            if let Ok(w) = arg.parse::<u16>() {
                let col = input.cursor.col;
                workbook.active_mut().col_meta.entry(col).or_default().width = Some(w.max(3));
                set_status(status, format!("Column width: {}", w.max(3)));
            } else {
                set_status(status, "Usage: :cw <width>".to_string());
            }
        }
        "rh" | "rowheight" => {
            if let Ok(h) = arg.parse::<u16>() {
                let row = input.cursor.row;
                workbook
                    .active_mut()
                    .row_meta
                    .entry(row)
                    .or_default()
                    .height = Some(h.max(1));
                set_status(status, format!("Row height: {}", h.max(1)));
            } else {
                set_status(status, "Usage: :rh <height>".to_string());
            }
        }
        "theme" => {
            if arg.is_empty() {
                // Open theme manager
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
            } else {
                // Apply theme by name or id
                let themes = asat_config::builtin_themes();
                let needle = arg.to_ascii_lowercase();
                if let Some(preset) = themes.iter().find(|t| {
                    t.id.to_ascii_lowercase() == needle
                        || t.name.to_ascii_lowercase() == needle
                        || t.name.to_ascii_lowercase().replace(' ', "-") == needle
                }) {
                    config.theme_name = preset.id.to_string();
                    config.theme = preset.config.clone();
                    match config.save() {
                        Ok(_) => set_status(status, format!("Theme \"{}\" applied", preset.name)),
                        Err(e) => {
                            set_status(status, format!("Theme applied but couldn't save: {}", e))
                        }
                    }
                } else {
                    set_status(
                        status,
                        format!("Unknown theme: {}  (use :theme to open picker)", arg),
                    );
                }
            }
        }
        "set" => {
            // Basic :set support
            if arg.is_empty() {
                set_status(status, "Usage: :set <option>=<value>".to_string());
            } else {
                set_status(status, format!("Unknown option: {}", arg));
            }
        }
        // ── Help screen ──
        "help" | "h" => {
            input.help_tab = 0;
            input.help_scroll = 0;
            input.help_query.clear();
            input.mode = Mode::Help;
        }

        "home" => {
            input.mode = Mode::Welcome;
        }

        // ── Plugin manager panel ──
        "plugins" | "plug-ui" => {
            input.plugin_selected = 0;
            input.plugin_show_output = false;
            input.mode = Mode::PluginManager;
        }

        // ── Plugin management ──
        "plugin" | "plug" => match arg {
            "reload" | "r" => {
                plugins.reload();
                set_status(status, "Plugin init.py reloaded".to_string());
            }
            "list" | "l" | "" => {
                let info = plugins.info();
                set_status(status, info);
            }
            _ => {
                set_status(status, "Usage: :plugin [reload|list]".to_string());
            }
        },
        // ── Sort ──
        "sort" | "so" => {
            // :sort [<col>][!|desc]  — sort rows by a column letter or cursor column
            // Parse optional column letter prefix (A–Z, AA, AB, …)
            let arg = arg.trim();
            let (col, rest) = {
                let bytes = arg.as_bytes();
                let letter_end = bytes.iter().take_while(|b| b.is_ascii_uppercase()).count();
                if letter_end > 0 {
                    let letters = &arg[..letter_end];
                    let idx = letters
                        .chars()
                        .fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1))
                        - 1;
                    (idx, arg[letter_end..].trim())
                } else {
                    (input.cursor.col, arg)
                }
            };
            let descending = rest == "!" || rest.eq_ignore_ascii_case("desc");
            let sheet_idx = workbook.active_sheet;
            let sheet = workbook.active();

            if sheet.cells.is_empty() {
                set_status(status, "Nothing to sort".to_string());
                return ActionResult::Continue;
            }

            let max_row = sheet.max_row();
            let max_col = sheet.max_col();

            // Snapshot every row as an ordered Vec of optional cells
            let mut rows: Vec<Vec<Option<asat_core::Cell>>> = (0..=max_row)
                .map(|r| {
                    (0..=max_col)
                        .map(|c| sheet.cells.get(&(r, c)).cloned())
                        .collect()
                })
                .collect();

            // Sort rows by the key column value
            rows.sort_by(|a, b| {
                let av = a
                    .get(col as usize)
                    .and_then(|c| c.as_ref())
                    .map(|c| &c.value)
                    .unwrap_or(&CellValue::Empty);
                let bv = b
                    .get(col as usize)
                    .and_then(|c| c.as_ref())
                    .map(|c| &c.value)
                    .unwrap_or(&CellValue::Empty);
                let ord = compare_cell_values(av, bv);
                if descending {
                    ord.reverse()
                } else {
                    ord
                }
            });

            // Build one SetCell command per changed cell
            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for (new_row_idx, row_cells) in rows.iter().enumerate() {
                for (c_idx, new_cell) in row_cells.iter().enumerate() {
                    let r = new_row_idx as u32;
                    let c = c_idx as u32;
                    let old_cell = workbook.active().get_cell(r, c);
                    let old_value = old_cell
                        .map(|x| x.value.clone())
                        .unwrap_or(CellValue::Empty);
                    let old_style = old_cell.and_then(|x| x.style.clone());
                    let new_value = new_cell
                        .as_ref()
                        .map(|x| x.value.clone())
                        .unwrap_or(CellValue::Empty);
                    let new_style = new_cell.as_ref().and_then(|x| x.style.clone());
                    if old_value != new_value || old_style != new_style {
                        cmds.push(Box::new(SetCell {
                            sheet: sheet_idx,
                            row: r,
                            col: c,
                            old_value,
                            new_value,
                            old_style,
                            new_style,
                        }));
                    }
                }
            }

            if cmds.is_empty() {
                set_status(status, "Already sorted".to_string());
            } else {
                let col_label = asat_core::col_to_letter(col);
                let dir = if descending { "desc" } else { "asc" };
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: format!("sort by {} {}", col_label, dir),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, format!("Sorted by {} ({})", col_label, dir));
                    }
                    Err(e) => set_status(status, format!("Sort error: {}", e)),
                }
            }
        }
        // ── Search & replace ──
        "s" | "substitute" => {
            // :s/pattern/replacement/[g][i]
            if arg.is_empty() {
                set_status(status, "Usage: :s/pattern/replacement/[g][i]".to_string());
                return ActionResult::Continue;
            }
            let delim = arg.chars().next().unwrap_or('/');
            let rest = &arg[delim.len_utf8()..];
            let parts: Vec<&str> = rest.splitn(3, delim).collect();
            if parts.len() < 2 {
                set_status(status, "Usage: :s/pattern/replacement/[g][i]".to_string());
                return ActionResult::Continue;
            }
            let pattern = parts[0];
            let replacement = parts[1];
            let flags = parts.get(2).copied().unwrap_or("");
            let global = flags.contains('g');
            let icase = flags.contains('i');

            let re_pat = if icase {
                format!("(?i){}", pattern)
            } else {
                pattern.to_string()
            };
            let re = match Regex::new(&re_pat) {
                Ok(r) => r,
                Err(e) => {
                    set_status(status, format!("Invalid regex: {}", e));
                    return ActionResult::Continue;
                }
            };

            let sheet_idx = workbook.active_sheet;
            let sheet = workbook.active();

            // Restrict to visual selection when active, otherwise whole sheet
            let (row_start, col_start, row_end, col_end) =
                if let Some(anchor) = &input.visual_anchor {
                    let cur = input.cursor;
                    (
                        cur.row.min(anchor.row),
                        cur.col.min(anchor.col),
                        cur.row.max(anchor.row),
                        cur.col.max(anchor.col),
                    )
                } else {
                    (0, 0, sheet.max_row(), sheet.max_col())
                };

            let candidates: Vec<(u32, u32, String)> = sheet
                .cells
                .iter()
                .filter_map(|((r, c), cell)| {
                    if *r < row_start || *r > row_end || *c < col_start || *c > col_end {
                        return None;
                    }
                    if let CellValue::Text(s) = &cell.value {
                        if re.is_match(s) {
                            return Some((*r, *c, s.clone()));
                        }
                    }
                    None
                })
                .collect();

            if candidates.is_empty() {
                set_status(status, format!("Pattern not found: {}", pattern));
                return ActionResult::Continue;
            }

            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for (r, c, text) in &candidates {
                let new_text = if global {
                    re.replace_all(text, replacement).to_string()
                } else {
                    re.replace(text, replacement).to_string()
                };
                let old_cell = workbook.active().get_cell(*r, *c);
                let old_style = old_cell.and_then(|x| x.style.clone());
                cmds.push(Box::new(SetCell {
                    sheet: sheet_idx,
                    row: *r,
                    col: *c,
                    old_value: CellValue::Text(text.clone()),
                    new_value: CellValue::Text(new_text),
                    old_style: old_style.clone(),
                    new_style: old_style,
                }));
            }

            let n = cmds.len();
            let grouped = Box::new(asat_commands::GroupedCommand {
                description: format!("replace /{}/{}/", pattern, replacement),
                commands: cmds,
            });
            match grouped.execute(workbook) {
                Ok(_) => {
                    undo.push(grouped);
                    set_status(status, format!("{} cell(s) replaced", n));
                }
                Err(e) => set_status(status, format!("Replace error: {}", e)),
            }
        }
        // ── Goto cell ──
        "goto" | "go" => {
            if arg.is_empty() {
                set_status(status, "Usage: :goto <cell>  e.g. :goto B15".to_string());
            } else if let Some((r, c)) = parse_cell_address(arg) {
                input.save_position(workbook.active_sheet);
                input.cursor.row = r;
                input.cursor.col = c;
                set_status(status, format!("→ {}", arg.to_uppercase()));
            } else {
                set_status(status, format!("Invalid cell address: {}", arg));
            }
        }

        // ── Named ranges ──
        "name" => {
            // :name <name> <range>  e.g.  :name sales A1:C10
            let parts: Vec<&str> = arg.splitn(2, ' ').collect();
            if parts.len() < 2 {
                set_status(
                    status,
                    "Usage: :name <name> <range>  e.g. :name sales A1:C10".to_string(),
                );
            } else {
                let range_name = parts[0].to_uppercase();
                let range_str = parts[1].trim();
                if let Some(range) = parse_range_address(range_str, workbook.active_sheet) {
                    workbook.named_ranges.insert(range_name.clone(), range);
                    workbook.dirty = true;
                    set_status(
                        status,
                        format!(
                            "Named range \"{}\" = {}",
                            range_name,
                            range_str.to_uppercase()
                        ),
                    );
                } else {
                    set_status(status, format!("Invalid range: {}", range_str));
                }
            }
        }

        // ── Filter rows ──
        "filter" => {
            if arg.trim().eq_ignore_ascii_case("off") || arg.trim().eq_ignore_ascii_case("clear") {
                // Clear filter
                for rm in workbook.active_mut().row_meta.values_mut() {
                    rm.hidden = false;
                }
                workbook.dirty = true;
                set_status(status, "Filter cleared".to_string());
            } else {
                // :filter <col> <op> <value>  e.g. :filter A >100
                // Or :filter <col-letter> <op> <value>
                let fparts: Vec<&str> = arg.splitn(3, ' ').collect();
                if fparts.len() < 3 {
                    set_status(
                        status,
                        "Usage: :filter <col> <op> <val>  e.g. :filter A >100".to_string(),
                    );
                } else {
                    let col_str = fparts[0];
                    let op = fparts[1];
                    let val_str = fparts[2];
                    let col = if let Ok(n) = col_str.parse::<u32>() {
                        n.saturating_sub(1)
                    } else if let Some(c) = asat_core::letter_to_col(col_str) {
                        c
                    } else {
                        set_status(status, format!("Invalid column: {}", col_str));
                        return ActionResult::Continue;
                    };
                    let val_num = val_str.parse::<f64>().ok();
                    let sheet = workbook.active();
                    let max_row = sheet.max_row();
                    let total = max_row + 1;
                    // Collect visibility decisions
                    let decisions: Vec<(u32, bool)> = (0..=max_row)
                        .map(|r| {
                            let cv = sheet.get_value(r, col);
                            (r, filter_row_matches(cv, op, val_str, val_num))
                        })
                        .collect();
                    let visible_count = decisions.iter().filter(|(_, keep)| *keep).count();
                    for (r, keep) in decisions {
                        workbook.active_mut().row_meta.entry(r).or_default().hidden = !keep;
                    }
                    workbook.dirty = true;
                    set_status(
                        status,
                        format!("Filter: {} of {} rows visible", visible_count, total),
                    );
                }
            }
        }

        // ── Transpose selection ──
        "transpose" | "tp" => {
            let sheet_idx = workbook.active_sheet;
            let (row_start, col_start, row_end, col_end) = {
                if let Some(anchor) = &input.visual_anchor {
                    let cur = input.cursor;
                    (
                        cur.row.min(anchor.row),
                        cur.col.min(anchor.col),
                        cur.row.max(anchor.row),
                        cur.col.max(anchor.col),
                    )
                } else {
                    // Transpose whole data range from cursor cell
                    let s = workbook.active();
                    (input.cursor.row, input.cursor.col, s.max_row(), s.max_col())
                }
            };
            // Snapshot
            let mut snapshot: Vec<(u32, u32, CellValue)> = Vec::new();
            for r in row_start..=row_end {
                for c in col_start..=col_end {
                    let v = workbook.active().get_raw_value(r, c).clone();
                    snapshot.push((r, c, v));
                }
            }
            // Build transposed SetCell commands
            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for (r, c, v) in &snapshot {
                let new_r = col_start + (c - col_start);
                let new_c = row_start + (r - row_start);
                let old_val = workbook.active().get_raw_value(new_r, new_c).clone();
                cmds.push(Box::new(SetCell::new(
                    workbook,
                    sheet_idx,
                    new_r,
                    new_c,
                    v.clone(),
                )));
                let _ = old_val;
            }
            // Clear original non-transposed cells that are now outside the transposed block
            let grouped = Box::new(asat_commands::GroupedCommand {
                description: "transpose".to_string(),
                commands: cmds,
            });
            match grouped.execute(workbook) {
                Ok(_) => {
                    undo.push(grouped);
                    set_status(status, "Transposed".to_string());
                }
                Err(e) => set_status(status, format!("Transpose error: {}", e)),
            }
            input.visual_anchor = None;
            input.mode = asat_input::Mode::Normal;
        }

        // ── Remove duplicates ──
        "dedup" => {
            let sheet_idx = workbook.active_sheet;
            let col = if arg.is_empty() {
                input.cursor.col
            } else if let Some(c) = asat_core::letter_to_col(arg) {
                c
            } else {
                arg.parse::<u32>()
                    .unwrap_or(input.cursor.col + 1)
                    .saturating_sub(1)
            };
            let sheet = workbook.active();
            let max_row = sheet.max_row();
            let max_col = sheet.max_col();
            // Track seen key values (display string of the key column)
            let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
            let mut rows_to_delete: Vec<u32> = Vec::new();
            for r in 0..=max_row {
                let key = sheet.display_value(r, col);
                if !seen.insert(key) {
                    rows_to_delete.push(r);
                }
            }
            if rows_to_delete.is_empty() {
                set_status(status, "No duplicates found".to_string());
            } else {
                // Clear duplicate rows (delete from bottom to avoid row-shift issues)
                let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
                for &r in rows_to_delete.iter().rev() {
                    for c in 0..=max_col {
                        cmds.push(Box::new(SetCell::new(
                            workbook,
                            sheet_idx,
                            r,
                            c,
                            CellValue::Empty,
                        )));
                    }
                }
                let n = rows_to_delete.len();
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: "dedup".to_string(),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, format!("Removed {} duplicate row(s)", n));
                    }
                    Err(e) => set_status(status, format!("Dedup error: {}", e)),
                }
            }
        }

        // ── Cell note / comment ──
        "note" | "note!" | "comment" => {
            let row = input.cursor.row;
            let col = input.cursor.col;
            let sheet = workbook.active_mut();
            if verb == "note!" || arg == "-" || arg == "clear" {
                // :note! or :note clear or :note - — remove the note
                sheet.notes.remove(&(row, col));
                workbook.dirty = true;
                set_status(status, "Note cleared".to_string());
            } else if arg.is_empty() {
                // :note — open popup with full note text
                if let Some(note) = sheet.notes.get(&(row, col)) {
                    input.note_popup = Some(note.clone());
                } else {
                    set_status(
                        status,
                        "No note on this cell (use :note <text> to set)".to_string(),
                    );
                }
            } else {
                // :note <text> — set/replace the note
                sheet.notes.insert((row, col), arg.to_string());
                workbook.dirty = true;
                set_status(status, format!("Note set: {}", arg));
            }
        }

        // ── Conditional formatting ──
        "colfmt" | "cf" => {
            // :cf <range> <condition> <value> bg=<hex> [fg=<hex>]
            // :cf clear — remove all CF rules from active sheet
            // :cf list  — show CF rule count
            use asat_core::{CfCondition, ConditionalFormat};

            if arg.trim() == "clear" {
                workbook.active_mut().conditional_formats.clear();
                workbook.dirty = true;
                set_status(status, "Conditional formats cleared".to_string());
                return ActionResult::Continue;
            }
            if arg.trim() == "list" {
                let n = workbook.active().conditional_formats.len();
                set_status(status, format!("{} conditional format rule(s)", n));
                return ActionResult::Continue;
            }

            // Tokenise by whitespace, preserving bg=... fg=... tokens
            let tokens: Vec<&str> = arg.split_whitespace().collect();
            // Minimum: <range> <cond> [<value>] bg=<hex>
            if tokens.len() < 3 {
                set_status(
                    status,
                    "Usage: :cf <range> <cond> [val] bg=#rrggbb [fg=#rrggbb]".to_string(),
                );
                return ActionResult::Continue;
            }

            // Parse range (first token). Support whole-column like "A:A" by treating as
            // full-sheet rows.
            let range_str = tokens[0];
            let sheet_idx = workbook.active_sheet;
            let (row_start, col_start, row_end, col_end) =
                if let Some(cr) = parse_range_address_cf(range_str, sheet_idx, workbook) {
                    (cr.row_start, cr.col_start, cr.row_end, cr.col_end)
                } else {
                    set_status(status, format!("Invalid range: {}", range_str));
                    return ActionResult::Continue;
                };

            // Parse condition keyword + optional value
            let cond_str = tokens[1];
            let val_token = tokens.get(2).copied().unwrap_or("");

            // Collect bg= fg= tokens (may be at index 2 or 3 depending on whether
            // condition takes a value)
            let mut bg_hex: Option<String> = None;
            let mut fg_hex: Option<String> = None;
            for tok in &tokens {
                if let Some(rest) = tok.strip_prefix("bg=") {
                    bg_hex = Some(rest.to_string());
                } else if let Some(rest) = tok.strip_prefix("fg=") {
                    fg_hex = Some(rest.to_string());
                }
            }

            if bg_hex.is_none() {
                set_status(
                    status,
                    "Usage: :cf <range> <cond> [val] bg=#rrggbb [fg=#rrggbb]".to_string(),
                );
                return ActionResult::Continue;
            }

            // Parse value for numeric conditions; for "contains" it stays as string
            let condition: Option<CfCondition> = match cond_str {
                ">" => val_token.parse::<f64>().ok().map(CfCondition::Gt),
                "<" => val_token.parse::<f64>().ok().map(CfCondition::Lt),
                ">=" => val_token.parse::<f64>().ok().map(CfCondition::Gte),
                "<=" => val_token.parse::<f64>().ok().map(CfCondition::Lte),
                "=" | "==" => val_token.parse::<f64>().ok().map(CfCondition::Eq),
                "!=" | "<>" => val_token.parse::<f64>().ok().map(CfCondition::Ne),
                "contains" => Some(CfCondition::Contains(val_token.to_string())),
                "blank" | "isblank" => Some(CfCondition::IsBlank),
                "error" | "iserror" => Some(CfCondition::IsError),
                _ => None,
            };

            let condition = match condition {
                Some(c) => c,
                None => {
                    set_status(
                        status,
                        format!(
                            "Unknown condition: {}.  Use >, <, >=, <=, =, !=, contains, blank, error",
                            cond_str
                        ),
                    );
                    return ActionResult::Continue;
                }
            };

            let cf = ConditionalFormat {
                row_start,
                col_start,
                row_end,
                col_end,
                condition,
                bg: bg_hex,
                fg: fg_hex,
            };
            workbook.active_mut().conditional_formats.push(cf);
            workbook.dirty = true;
            let n = workbook.active().conditional_formats.len();
            set_status(
                status,
                format!("Conditional format rule added (total: {})", n),
            );
            input.visual_anchor = None;
            input.mode = asat_input::Mode::Normal;
        }

        // ── Fill down / fill right via ex-command ──
        "filldown" | "fd" => {
            let sheet_idx = workbook.active_sheet;
            let (row_start, col_start, row_end, col_end) = {
                if let Some(anchor) = &input.visual_anchor {
                    let cur = input.cursor;
                    let s = workbook.active();
                    (
                        cur.row.min(anchor.row),
                        cur.col.min(anchor.col),
                        cur.row.max(anchor.row).min(s.max_row()),
                        cur.col.max(anchor.col).min(s.max_col()),
                    )
                } else {
                    (
                        input.cursor.row,
                        input.cursor.col,
                        workbook.active().max_row(),
                        input.cursor.col,
                    )
                }
            };
            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for c in col_start..=col_end {
                let av = workbook.active().get_raw_value(row_start, c).clone();
                let asty = workbook
                    .active()
                    .get_cell(row_start, c)
                    .and_then(|x| x.style.clone());
                for r in (row_start + 1)..=row_end {
                    let old = workbook.active().get_cell(r, c);
                    let ov = old.map(|x| x.value.clone()).unwrap_or(CellValue::Empty);
                    let os = old.and_then(|x| x.style.clone());
                    cmds.push(Box::new(SetCell {
                        sheet: sheet_idx,
                        row: r,
                        col: c,
                        old_value: ov,
                        new_value: av.clone(),
                        old_style: os,
                        new_style: asty.clone(),
                    }));
                }
            }
            if cmds.is_empty() {
                set_status(status, "Nothing to fill down".to_string());
            } else {
                let n = cmds.len();
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: "fill down".to_string(),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, format!("Filled {} cell(s) down", n));
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }
        "fillright" | "fr" => {
            let sheet_idx = workbook.active_sheet;
            let (row_start, col_start, row_end, col_end) = {
                if let Some(anchor) = &input.visual_anchor {
                    let cur = input.cursor;
                    let s = workbook.active();
                    (
                        cur.row.min(anchor.row),
                        cur.col.min(anchor.col),
                        cur.row.max(anchor.row).min(s.max_row()),
                        cur.col.max(anchor.col).min(s.max_col()),
                    )
                } else {
                    (
                        input.cursor.row,
                        input.cursor.col,
                        input.cursor.row,
                        workbook.active().max_col(),
                    )
                }
            };
            let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
            for r in row_start..=row_end {
                let av = workbook.active().get_raw_value(r, col_start).clone();
                let asty = workbook
                    .active()
                    .get_cell(r, col_start)
                    .and_then(|x| x.style.clone());
                for c in (col_start + 1)..=col_end {
                    let old = workbook.active().get_cell(r, c);
                    let ov = old.map(|x| x.value.clone()).unwrap_or(CellValue::Empty);
                    let os = old.and_then(|x| x.style.clone());
                    cmds.push(Box::new(SetCell {
                        sheet: sheet_idx,
                        row: r,
                        col: c,
                        old_value: ov,
                        new_value: av.clone(),
                        old_style: os,
                        new_style: asty.clone(),
                    }));
                }
            }
            if cmds.is_empty() {
                set_status(status, "Nothing to fill right".to_string());
            } else {
                let n = cmds.len();
                let grouped = Box::new(asat_commands::GroupedCommand {
                    description: "fill right".to_string(),
                    commands: cmds,
                });
                match grouped.execute(workbook) {
                    Ok(_) => {
                        undo.push(grouped);
                        set_status(status, format!("Filled {} cell(s) right", n));
                    }
                    Err(e) => set_status(status, format!("Error: {}", e)),
                }
            }
        }

        "merge" => {
            let (row_start, col_start, row_end, col_end) = {
                if let Some(anchor) = &input.visual_anchor {
                    let cur = input.cursor;
                    let s = workbook.active();
                    (
                        cur.row.min(anchor.row),
                        cur.col.min(anchor.col),
                        cur.row.max(anchor.row).min(s.max_row()),
                        cur.col.max(anchor.col).min(s.max_col()),
                    )
                } else {
                    (
                        input.cursor.row,
                        input.cursor.col,
                        input.cursor.row,
                        input.cursor.col,
                    )
                }
            };
            let sheet_idx = workbook.active_sheet;
            let cmd = Box::new(MergeCells::new(
                workbook, sheet_idx, row_start, col_start, row_end, col_end,
            ));
            match cmd.execute(workbook) {
                Ok(_) => {
                    undo.push(cmd);
                    input.visual_anchor = None;
                    set_status(
                        status,
                        format!(
                            "Merged {}:{}",
                            cell_address(row_start, col_start),
                            cell_address(row_end, col_end),
                        ),
                    );
                }
                Err(e) => set_status(status, format!("Error: {}", e)),
            }
        }

        "unmerge" => {
            let sheet_idx = workbook.active_sheet;
            let row = input.cursor.row;
            let col = input.cursor.col;
            let cmd = Box::new(UnmergeCells::new(workbook, sheet_idx, row, col));
            match cmd.execute(workbook) {
                Ok(_) => {
                    undo.push(cmd);
                    set_status(status, format!("Unmerged {}", cell_address(row, col)));
                }
                Err(e) => set_status(status, format!("Error: {}", e)),
            }
        }

        // :freeze rows N  /  :freeze cols N  /  :freeze off
        "freeze" => {
            let sheet = workbook.sheets.get_mut(workbook.active_sheet).unwrap();
            let parts: Vec<&str> = arg.splitn(2, ' ').collect();
            match parts.as_slice() {
                ["rows", n] | ["row", n] => match n.parse::<u32>() {
                    Ok(v) => {
                        sheet.freeze_rows = v;
                        set_status(status, format!("Frozen {} row(s)", v));
                    }
                    Err(_) => set_status(status, "Usage: :freeze rows <N>".to_string()),
                },
                ["cols", n] | ["col", n] | ["columns", n] => match n.parse::<u32>() {
                    Ok(v) => {
                        sheet.freeze_cols = v;
                        set_status(status, format!("Frozen {} column(s)", v));
                    }
                    Err(_) => set_status(status, "Usage: :freeze cols <N>".to_string()),
                },
                ["off"] | ["none"] | ["0"] => {
                    sheet.freeze_rows = 0;
                    sheet.freeze_cols = 0;
                    set_status(status, "Freeze panes cleared".to_string());
                }
                _ => set_status(
                    status,
                    "Usage: :freeze rows <N>  |  :freeze cols <N>  |  :freeze off".to_string(),
                ),
            }
        }

        _ => {
            set_status(status, format!("Unknown command: {}", verb));
        }
    }
    ActionResult::Continue
}
