use std::io;
use std::path::PathBuf;
use std::time::Duration;

use anyhow::Result;
use arboard::Clipboard;
use crossterm::{
    event::{self, Event, KeyEvent},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{backend::CrosstermBackend, Terminal};
use regex::Regex;

use asat_commands::{Command, DeleteCol, InsertCol, MergeCells, SetCell, UndoStack, UnmergeCells};
use asat_config::Config;
use asat_core::{cell_address, CellStyle, CellValue, Workbook};
use asat_input::{AppAction, InputState, Mode};
use asat_plugins::{PluginEvent, PluginManager, PluginOutput};
use asat_tui::{render, RenderState};

fn main() -> Result<()> {
    // Parse CLI arguments
    let args: Vec<String> = std::env::args().collect();

    // Handle --version / -V before doing anything else
    if args
        .get(1)
        .map(|s| s == "--version" || s == "-V")
        .unwrap_or(false)
    {
        println!("asat {}", env!("CARGO_PKG_VERSION"));
        return Ok(());
    }
    if args
        .get(1)
        .map(|s| s == "--help" || s == "-h")
        .unwrap_or(false)
    {
        println!(
            "asat {} — terminal spreadsheet editor",
            env!("CARGO_PKG_VERSION")
        );
        println!();
        println!("USAGE:");
        println!("  asat [FILE]");
        println!();
        println!("ARGS:");
        println!("  [FILE]  File to open (csv, xlsx, ods, tsv, .asat). Creates new if absent.");
        println!();
        println!("OPTIONS:");
        println!("  -V, --version  Print version");
        println!("  -h, --help     Print this help");
        println!();
        println!("Website: https://okt4v.github.io/ASAT/");
        return Ok(());
    }

    let file_path: Option<PathBuf> = args.get(1).map(PathBuf::from);

    // Load config
    let config = Config::load().unwrap_or_default();

    // Load workbook
    let workbook = if let Some(ref path) = file_path {
        if path.exists() {
            match asat_io::load(path) {
                Ok(wb) => wb,
                Err(e) => {
                    eprintln!("Error loading {:?}: {}", path, e);
                    std::process::exit(1);
                }
            }
        } else {
            // New file
            let mut wb = Workbook::new();
            wb.file_path = Some(path.clone());
            wb
        }
    } else {
        Workbook::new()
    };

    // Swap file recovery: check before entering raw mode so we can prompt normally
    let workbook = if let Some(ref path) = file_path {
        let swp = swap_path(path);
        if swp.exists() {
            eprintln!(
                "Swap file found: {:?}\nThis may mean asat crashed with unsaved changes.",
                swp
            );
            eprint!("Recover? [Y/n] ");
            let mut answer = String::new();
            let _ = std::io::stdin().read_line(&mut answer);
            let recovered = if answer.trim().to_lowercase() != "n" {
                match asat_io::load(&swp) {
                    Ok(mut wb) => {
                        wb.file_path = Some(path.clone());
                        wb.dirty = true;
                        eprintln!("Recovered. Use :w to save.");
                        wb
                    }
                    Err(e) => {
                        eprintln!("Could not read swap file: {}", e);
                        workbook
                    }
                }
            } else {
                workbook
            };
            // Always delete the stale swap file — a new one will be created as edits happen
            let _ = std::fs::remove_file(&swp);
            recovered
        } else {
            workbook
        }
    } else {
        workbook
    };

    // Setup terminal
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;
    terminal.hide_cursor()?;

    // Run app
    let recent = load_recent_files(config.remember_recent as usize);
    let result = run_app(&mut terminal, workbook, config, file_path, recent);

    // Restore terminal
    disable_raw_mode()?;
    execute!(terminal.backend_mut(), LeaveAlternateScreen)?;
    terminal.show_cursor()?;

    if let Err(e) = result {
        eprintln!("Error: {}", e);
        std::process::exit(1);
    }

    Ok(())
}

fn swap_path(file_path: &std::path::Path) -> PathBuf {
    let name = file_path.file_name().unwrap_or_default().to_string_lossy();
    file_path
        .parent()
        .unwrap_or_else(|| std::path::Path::new("."))
        .join(format!(".{}.swp", name))
}

fn write_swap(workbook: &Workbook) {
    if let Some(ref p) = workbook.file_path {
        let _ = asat_io::save_swap(workbook, &swap_path(p));
    }
}

fn delete_swap(workbook: &Workbook) {
    if let Some(ref p) = workbook.file_path {
        let _ = std::fs::remove_file(swap_path(p));
    }
}

fn run_app(
    terminal: &mut Terminal<CrosstermBackend<io::Stdout>>,
    mut workbook: Workbook,
    mut config: Config,
    file_path: Option<PathBuf>,
    recent: Vec<String>,
) -> Result<()> {
    let mut input_state = InputState::new();
    input_state.scroll_padding = config.scroll_padding;
    input_state.recent_files = recent;

    // Show welcome screen only when no file was given
    if file_path.is_none() {
        input_state.mode = Mode::Welcome;
    }

    let mut undo_stack = UndoStack::with_limit(config.undo_limit);
    // Keep a single Clipboard instance alive for the whole session.
    // On Linux the clipboard is X11/Wayland-owned by this process; dropping the
    // handle clears the contents immediately, which is why clipboard managers
    // see 0ms content. Reusing one instance avoids that.
    let mut clipboard = Clipboard::new().ok();
    let mut plugins = PluginManager::new();
    plugins.load_init_script();

    // Fire Open event if a file was given at startup
    if file_path.is_some() {
        plugins.push_event(PluginEvent::Open {
            path: workbook.file_path.as_ref().map(|p| p.display().to_string()),
        });
    }

    let mut status_message: Option<(String, std::time::Instant)> = None;

    // Hint when init.py exists but Python support was not compiled in
    #[cfg(not(feature = "python"))]
    {
        let init_py = asat_config::config_dir().join("init.py");
        if init_py.exists() {
            set_status(
                &mut status_message,
                "init.py found but Python support not compiled in. \
                 Rebuild with --features asat-plugins/python"
                    .to_string(),
            );
        }
    }
    let status_timeout = if config.status_timeout == 0 {
        Duration::from_secs(u64::MAX / 2) // effectively never
    } else {
        Duration::from_secs(config.status_timeout as u64)
    };

    // Autosave: time-based — save every autosave_interval seconds when dirty
    let mut edit_count: u32 = 0;
    let mut last_autosave = std::time::Instant::now();
    let mut last_swap = std::time::Instant::now();

    loop {
        // Clear expired status messages
        if let Some((_, ts)) = &status_message {
            if ts.elapsed() > status_timeout {
                status_message = None;
            }
        }

        // Dispatch queued plugin events and process outputs from Python handlers
        for output in plugins.drain(&workbook) {
            match output {
                PluginOutput::Notify(msg) => {
                    set_status(&mut status_message, msg);
                }
                PluginOutput::Command(cmd) => {
                    handle_ex_command(
                        &cmd,
                        &mut workbook,
                        &mut input_state,
                        &mut undo_stack,
                        &mut status_message,
                        &mut config,
                        &mut plugins,
                    );
                }
                PluginOutput::SetCell {
                    sheet,
                    row,
                    col,
                    value,
                } => {
                    // sentinel usize::MAX means "active sheet"
                    let sheet_idx = if sheet == usize::MAX {
                        workbook.active_sheet
                    } else {
                        sheet
                    };
                    let cmd = Box::new(SetCell::new(&workbook, sheet_idx, row, col, value));
                    let _ = cmd.execute(&mut workbook);
                    // (plugin-driven cell sets bypass the undo stack intentionally)
                }
            }
        }

        // Autosave: save every autosave_interval seconds when the workbook is dirty
        if config.autosave_interval > 0
            && workbook.dirty
            && last_autosave.elapsed() >= Duration::from_secs(config.autosave_interval as u64)
        {
            last_autosave = std::time::Instant::now();
            if let Some(path) = workbook.file_path.clone() {
                if asat_io::save(&workbook, &path).is_ok() {
                    workbook.dirty = false;
                    set_status(
                        &mut status_message,
                        format!(
                            "Autosaved \"{}\"",
                            path.file_name().unwrap_or_default().to_string_lossy()
                        ),
                    );
                }
            }
        }

        // Swap file: write every 30 seconds when dirty (crash recovery)
        if workbook.dirty && last_swap.elapsed() >= Duration::from_secs(30) {
            last_swap = std::time::Instant::now();
            write_swap(&workbook);
        }

        // Recalculate all formula cells so the renderer always shows fresh values
        recalculate_all(&mut workbook);

        // Render
        terminal.draw(|frame| {
            let msg_ref = status_message.as_ref().map(|(m, _)| m.as_str());
            let formula_preview = if matches!(input_state.mode, Mode::Insert { .. })
                && input_state.edit_buffer.starts_with('=')
                && input_state.edit_buffer.len() > 1
            {
                let formula_str = &input_state.edit_buffer[1..];
                let sheet_idx = workbook.active_sheet;
                let row = input_state.cursor.row;
                let col = input_state.cursor.col;
                let val = asat_formula::evaluate(formula_str, &workbook, sheet_idx, row, col);
                Some(val.display())
            } else {
                None
            };
            // Collect cell references for formula highlight (Feature 3)
            let ref_cells: std::collections::HashSet<(u32, u32)> =
                if matches!(input_state.mode, Mode::Normal) {
                    let row = input_state.cursor.row;
                    let col = input_state.cursor.col;
                    match workbook.active().get_raw_value(row, col) {
                        asat_core::CellValue::Formula(f) => {
                            asat_formula::collect_cell_refs(f).into_iter().collect()
                        }
                        _ => std::collections::HashSet::new(),
                    }
                } else {
                    std::collections::HashSet::new()
                };

            let plugin_info = plugins.info();
            let plugin_custom_fns = asat_core::list_custom_fns();
            let state = RenderState {
                workbook: &workbook,
                input: &input_state,
                status_message: msg_ref,
                show_side_panel: false,
                config: &config,
                formula_preview,
                ref_cells,
                plugin_info,
                plugin_custom_fns,
            };
            render(frame, &state);
        })?;

        // Compute grid dimensions from actual terminal size for accurate scrolling
        let (visible_rows, visible_cols) = {
            let size = terminal.size()?;
            let show_cmd = matches!(input_state.mode, Mode::Command | Mode::Search { .. });
            // rows: total - formula(1) - tab(1) - status(1) - cmd(0/1) - col_header(1)
            let non_grid = if show_cmd { 5u16 } else { 4u16 };
            let grid_h = size.height.saturating_sub(non_grid);
            let grid_w = size.width;
            let vrows =
                visible_rows_in_height(grid_h, workbook.active(), input_state.viewport.top_row);
            let vcols =
                visible_cols_in_width(grid_w, workbook.active(), input_state.viewport.left_col);
            (vrows, vcols)
        };

        // Poll for input (16ms ≈ 60fps)
        if !event::poll(Duration::from_millis(16))? {
            continue;
        }

        let ev = event::read()?;
        if let Event::Key(key) = ev {
            let mode_before = input_state.mode.name().to_string();
            let actions = input_state.handle_key(key, &workbook);
            let mut should_quit = false;
            let mut force_quit = false;

            for action in actions {
                match process_action(
                    action,
                    &mut workbook,
                    &mut input_state,
                    &mut undo_stack,
                    &mut status_message,
                    &mut config,
                    &mut plugins,
                    &mut clipboard,
                    visible_rows,
                    visible_cols,
                    &mut edit_count,
                ) {
                    ActionResult::Continue => {}
                    ActionResult::Quit => should_quit = true,
                    ActionResult::ForceQuit => force_quit = true,
                    ActionResult::ClearTerminal => {
                        // An external process (e.g. editor) returned the terminal to us.
                        // Flush any stale input bytes and force a full redraw.
                        let _ = terminal.clear();
                    }
                }
            }

            // Fire ModeChange event when the mode has changed
            let mode_after = input_state.mode.name().to_string();
            if mode_after != mode_before {
                plugins.push_event(PluginEvent::ModeChange { mode: mode_after });
            }

            // Keep sub-command completions fresh after every key event
            update_subcmd_completions(&mut input_state);

            if force_quit {
                delete_swap(&workbook);
                return Ok(());
            }
            if should_quit {
                if workbook.dirty {
                    set_status(
                        &mut status_message,
                        "Unsaved changes. Use :w to save, :q! to force quit".to_string(),
                    );
                } else {
                    delete_swap(&workbook);
                    return Ok(());
                }
            }
        }
    }
}

enum ActionResult {
    Continue,
    Quit,
    ForceQuit,
    /// External process took over the terminal; force a full redraw on return.
    ClearTerminal,
}

/// Evaluate all formula cells in every sheet and store results in `sheet.computed`.
fn recalculate_all(workbook: &mut Workbook) {
    for idx in 0..workbook.sheets.len() {
        recalculate_sheet(workbook, idx);
    }
}

fn recalculate_sheet(workbook: &mut Workbook, sheet_idx: usize) {
    // Collect formula cells (immutable borrow ends at the semicolon)
    let formula_cells: Vec<(u32, u32, String)> = {
        let sheet = &workbook.sheets[sheet_idx];
        sheet
            .cells
            .iter()
            .filter_map(|((r, c), cell)| {
                if let CellValue::Formula(f) = &cell.value {
                    Some((*r, *c, f.clone()))
                } else {
                    None
                }
            })
            .collect()
    };

    if formula_cells.is_empty() {
        workbook.sheets[sheet_idx].computed.clear();
        return;
    }

    workbook.sheets[sheet_idx].computed.clear();

    // ── Circular reference detection ──────────────────────────────────────────
    // Build a dependency map: (row, col) → set of (row, col) it references on the same sheet.
    let formula_set: std::collections::HashSet<(u32, u32)> =
        formula_cells.iter().map(|(r, c, _)| (*r, *c)).collect();

    let deps: std::collections::HashMap<(u32, u32), Vec<(u32, u32)>> = formula_cells
        .iter()
        .map(|(r, c, f)| {
            let refs: Vec<(u32, u32)> = asat_formula::collect_cell_refs(f)
                .into_iter()
                .filter(|rc| formula_set.contains(rc))
                .collect();
            ((*r, *c), refs)
        })
        .collect();

    // DFS cycle detection: cells found in a cycle get #CIRC! instead of being evaluated.
    let mut cycle_cells = std::collections::HashSet::new();
    {
        #[derive(PartialEq)]
        enum Color {
            White,
            Gray,
            Black,
        }
        let mut color: std::collections::HashMap<(u32, u32), Color> = formula_cells
            .iter()
            .map(|(r, c, _)| ((*r, *c), Color::White))
            .collect();

        fn dfs(
            node: (u32, u32),
            deps: &std::collections::HashMap<(u32, u32), Vec<(u32, u32)>>,
            color: &mut std::collections::HashMap<(u32, u32), Color>,
            cycle_cells: &mut std::collections::HashSet<(u32, u32)>,
        ) {
            if color.get(&node) == Some(&Color::Black) {
                return;
            }
            if color.get(&node) == Some(&Color::Gray) {
                cycle_cells.insert(node);
                return;
            }
            color.insert(node, Color::Gray);
            if let Some(neighbors) = deps.get(&node) {
                for &neighbor in neighbors {
                    if color.get(&neighbor) == Some(&Color::Gray) {
                        // Back-edge: both ends are in a cycle
                        cycle_cells.insert(node);
                        cycle_cells.insert(neighbor);
                    } else {
                        dfs(neighbor, deps, color, cycle_cells);
                        if cycle_cells.contains(&neighbor) {
                            cycle_cells.insert(node);
                        }
                    }
                }
            }
            color.insert(node, Color::Black);
        }

        let keys: Vec<(u32, u32)> = deps.keys().copied().collect();
        for node in keys {
            dfs(node, &deps, &mut color, &mut cycle_cells);
        }
    }

    // Mark cycle cells immediately so passes skip them.
    let sheet = &mut workbook.sheets[sheet_idx];
    for &(r, c) in &cycle_cells {
        sheet.computed.insert(
            (r, c),
            asat_core::CellValue::Error(asat_core::CellError::CircularRef),
        );
    }

    // Up to 3 passes so formula→formula chains resolve correctly
    for _pass in 0..3 {
        let results: Vec<(u32, u32, CellValue)> = formula_cells
            .iter()
            .filter(|(r, c, _)| !cycle_cells.contains(&(*r, *c)))
            .map(|(r, c, f)| {
                let val = asat_formula::evaluate(f, workbook, sheet_idx, *r, *c);
                (*r, *c, val)
            })
            .collect();
        let sheet = &mut workbook.sheets[sheet_idx];
        for (r, c, val) in results {
            sheet.computed.insert((r, c), val);
        }
    }
}

/// Compute how many spreadsheet rows are fully visible starting from `top_row`.
fn visible_rows_in_height(grid_height: u16, sheet: &asat_core::Sheet, top_row: u32) -> u32 {
    let mut remaining = grid_height as u32;
    let mut count = 0u32;
    let mut r = top_row;
    loop {
        let h = sheet.row_height(r) as u32;
        if remaining < h {
            break;
        }
        remaining -= h;
        count += 1;
        r += 1;
        if r > 1_000_000 {
            break;
        }
    }
    count.max(1)
}

/// Compute how many columns are visible starting from `left_col` given `width` pixels.
/// Mirrors the gutter width and col_width logic in grid.rs.
fn visible_cols_in_width(grid_width: u16, sheet: &asat_core::Sheet, left_col: u32) -> u32 {
    const GUTTER: u16 = 5;
    const MIN_COL: u16 = 3;
    let mut avail = grid_width.saturating_sub(GUTTER);
    let mut count = 0u32;
    let mut c = left_col;
    loop {
        let w = sheet.col_width(c).max(MIN_COL);
        if avail < w {
            break;
        }
        avail -= w;
        count += 1;
        c += 1;
        if c > 10_000 {
            break;
        }
    }
    count.max(1)
}

/// Write `text` to the system clipboard using a persistent Clipboard instance.
fn copy_to_clipboard(cb: &mut Option<Clipboard>, text: &str) {
    if let Some(cb) = cb {
        let _ = cb.set_text(text.to_owned());
    }
}

/// Serialise a 2-D grid of CellValues as tab-separated values (rows separated by newlines).
fn cells_to_tsv(cells: &[Vec<CellValue>]) -> String {
    cells
        .iter()
        .map(|row| {
            row.iter()
                .map(|v| v.display())
                .collect::<Vec<_>>()
                .join("\t")
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// Compare two CellValues for sorting purposes.
/// Order: numbers < text < booleans < errors < empty
fn compare_cell_values(a: &CellValue, b: &CellValue) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    match (a, b) {
        (CellValue::Empty, CellValue::Empty) => Ordering::Equal,
        (CellValue::Empty, _) => Ordering::Greater, // empty sorts last
        (_, CellValue::Empty) => Ordering::Less,
        (CellValue::Number(x), CellValue::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (CellValue::Boolean(x), CellValue::Boolean(y)) => x.cmp(y),
        (CellValue::Text(x), CellValue::Text(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
        // Mixed types: numbers first, then text, booleans, errors
        (CellValue::Number(_), _) => Ordering::Less,
        (_, CellValue::Number(_)) => Ordering::Greater,
        (CellValue::Text(_), _) => Ordering::Less,
        (_, CellValue::Text(_)) => Ordering::Greater,
        _ => Ordering::Equal,
    }
}

#[allow(clippy::too_many_arguments)]
fn process_action(
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
                    Box::new(SetCell::new(workbook, sheet_idx, r, c, CellValue::Empty))
                        as Box<dyn asat_commands::Command>
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
            Ok(true) => set_status(status, "Undo".to_string()),
            Ok(false) => set_status(status, "Already at oldest change".to_string()),
            Err(e) => set_status(status, format!("Undo error: {}", e)),
        },
        AppAction::Redo => match undo.redo(workbook) {
            Ok(true) => set_status(status, "Redo".to_string()),
            Ok(false) => set_status(status, "Already at newest change".to_string()),
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
                .map(|c| sheet.get_value(row, c).clone())
                .collect()];
            let tsv = cells_to_tsv(&cells);
            copy_to_clipboard(clipboard, &tsv);
            input.registers.yank(None, cells, true);
            set_status(status, format!("Yanked row {} → clipboard", row + 1));
        }
        AppAction::YankCell => {
            let sheet = workbook.active();
            let (row, col) = (input.cursor.row, input.cursor.col);
            let val = sheet.get_value(row, col).clone();
            let text = val.display();
            copy_to_clipboard(clipboard, &text);
            let cells = vec![vec![val]];
            input.registers.yank(None, cells, false);
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
                    .map(|c| sheet.get_value(row, c).clone())
                    .collect()];
                let tsv = cells_to_tsv(&cells);
                copy_to_clipboard(clipboard, &tsv);
                input.registers.yank(None, cells, true);
                set_status(status, format!("Yanked row {} → clipboard", row + 1));
            }
        }
        AppAction::YankCol => {
            let sheet = workbook.active();
            let col = input.cursor.col;
            let max_row = sheet.max_row();
            let cells: Vec<Vec<CellValue>> = (0..=max_row)
                .map(|r| vec![sheet.get_value(r, col).clone()])
                .collect();
            let tsv = cells_to_tsv(&cells);
            copy_to_clipboard(clipboard, &tsv);
            input.registers.yank(None, cells, false);
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
                        .map(|c| sheet.get_value(r, c).clone())
                        .collect()
                })
                .collect();
            let rows = cells.len();
            let cols = cells.first().map(|r| r.len()).unwrap_or(0);
            let tsv = cells_to_tsv(&cells);
            copy_to_clipboard(clipboard, &tsv);
            input.registers.yank(None, cells, is_line);
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
                    input.search_matches = matches;
                    input.search_match_idx = idx;
                    set_status(status, format!("/{} [{}/{}]", pattern, idx + 1, total));
                }
            }
        }
        AppAction::ClearSearch => {
            input.search_matches.clear();
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
            *workbook = Workbook::new();
            input.mode = Mode::Normal;
            set_status(
                status,
                "New workbook — press i to start editing".to_string(),
            );
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
    let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
    for (dr, row_vals) in reg.cells.iter().enumerate() {
        for (dc, val) in row_vals.iter().enumerate() {
            let r = start_row + dr as u32;
            let c = start_col + dc as u32;
            cmds.push(Box::new(SetCell::new(
                workbook,
                sheet_idx,
                r,
                c,
                val.clone(),
            )));
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

#[allow(clippy::too_many_arguments)]
fn replay_macro(
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

// ── Color helpers ─────────────────────────────────────────────────────────────

/// Parse a color argument: either `#rrggbb` hex or a named color.
fn parse_color_arg(arg: &str) -> Option<asat_core::Color> {
    let s = arg.trim_start_matches('#');
    // Try 6-digit hex
    if s.len() == 6 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&s[0..2], 16),
            u8::from_str_radix(&s[2..4], 16),
            u8::from_str_radix(&s[4..6], 16),
        ) {
            return Some(asat_core::Color::rgb(r, g, b));
        }
    }
    // Named colors
    let (r, g, b): (u8, u8, u8) = match arg.to_ascii_lowercase().as_str() {
        "red" => (220, 50, 50),
        "darkred" | "maroon" => (140, 20, 20),
        "lightred" | "salmon" => (255, 120, 120),
        "orange" => (230, 130, 30),
        "gold" => (220, 180, 30),
        "yellow" => (220, 200, 50),
        "lightyellow" => (255, 255, 150),
        "lime" => (100, 220, 50),
        "green" => (50, 180, 80),
        "darkgreen" | "olive" => (50, 120, 30),
        "lightgreen" => (120, 220, 120),
        "teal" | "cyan" => (40, 180, 180),
        "lightcyan" => (100, 230, 230),
        "blue" => (50, 100, 220),
        "navy" | "darkblue" => (20, 50, 140),
        "lightblue" | "sky" => (100, 160, 255),
        "purple" | "violet" => (140, 80, 220),
        "indigo" => (80, 50, 160),
        "lightpurple" | "lavender" => (180, 140, 255),
        "pink" => (255, 150, 190),
        "hotpink" | "magenta" => (220, 60, 160),
        "brown" | "sienna" => (160, 90, 40),
        "white" => (245, 245, 245),
        "lightgray" | "lightgrey" => (200, 200, 200),
        "gray" | "grey" => (130, 130, 130),
        "darkgray" | "darkgrey" => (70, 70, 70),
        "black" => (15, 15, 15),
        _ => return None,
    };
    Some(asat_core::Color::rgb(r, g, b))
}

/// Returns true when the given (r,g,b) Color is perceptually dark.
fn color_is_dark(c: &asat_core::Color) -> bool {
    (c.r as u32 * 299 + c.g as u32 * 587 + c.b as u32 * 114) < 128_000
}

// ── Selection-aware style helpers ─────────────────────────────────────────────

/// Return the style range: the visual selection if active, otherwise the cursor cell.
fn style_range(workbook: &Workbook, input: &InputState) -> (u32, u32, u32, u32) {
    use asat_input::Mode;
    if matches!(input.mode, Mode::Visual { .. } | Mode::VisualLine) {
        let (rs, cs, re, ce) = input.visual_selection_bounds();
        let sh = workbook.active();
        (rs, cs, re.min(sh.max_row()), ce.min(sh.max_col()))
    } else if let Some((rs, cs, re, ce)) = input.visual_command_range {
        // Entered Command mode from Visual — apply to the saved selection
        let sh = workbook.active();
        (rs, cs, re.min(sh.max_row()), ce.min(sh.max_col()))
    } else {
        (
            input.cursor.row,
            input.cursor.col,
            input.cursor.row,
            input.cursor.col,
        )
    }
}

/// Apply a style mutation to the current cell or visual selection (with undo).
fn apply_style_sel(
    workbook: &mut Workbook,
    input: &InputState,
    undo: &mut UndoStack,
    status: &mut Option<(String, std::time::Instant)>,
    f: &dyn Fn(&mut CellStyle),
    msg: &str,
) {
    let sheet_idx = workbook.active_sheet;
    let (row_start, col_start, row_end, col_end) = style_range(workbook, input);

    // Single cell — use the existing single-cell helper
    if row_start == row_end && col_start == col_end {
        apply_style(
            workbook, undo, status, sheet_idx, row_start, col_start, f, msg,
        );
        return;
    }

    // Range — check size
    let count = (row_end - row_start + 1) as u64 * (col_end - col_start + 1) as u64;
    if count > 50_000 {
        set_status(status, "Selection too large (max 50 000 cells)".to_string());
        return;
    }

    let mut cmds: Vec<Box<dyn asat_commands::Command>> = Vec::new();
    for r in row_start..=row_end {
        for c in col_start..=col_end {
            let old_cell = workbook.sheet(sheet_idx).and_then(|s| s.get_cell(r, c));
            let old_value = old_cell
                .map(|c| c.value.clone())
                .unwrap_or(CellValue::Empty);
            let old_style = old_cell.and_then(|c| c.style.clone());
            let mut new_style = old_style.clone().unwrap_or_default();
            f(&mut new_style);
            cmds.push(Box::new(SetCell {
                sheet: sheet_idx,
                row: r,
                col: c,
                old_value: old_value.clone(),
                new_value: old_value,
                old_style,
                new_style: Some(new_style),
            }));
        }
    }
    let n_rows = row_end - row_start + 1;
    let n_cols = col_end - col_start + 1;
    let grouped = Box::new(asat_commands::GroupedCommand {
        description: msg.to_string(),
        commands: cmds,
    });
    match grouped.execute(workbook) {
        Ok(_) => {
            undo.push(grouped);
            set_status(status, format!("{} ({}×{})", msg, n_rows, n_cols));
        }
        Err(e) => set_status(status, format!("Error: {}", e)),
    }
}

/// Fetch the current CellStyle for a cell (or default), apply `f`, then push as a SetCell command.
#[allow(clippy::too_many_arguments)]
fn apply_style(
    workbook: &mut Workbook,
    undo: &mut UndoStack,
    status: &mut Option<(String, std::time::Instant)>,
    sheet_idx: usize,
    row: u32,
    col: u32,
    f: impl FnOnce(&mut CellStyle),
    msg: &str,
) {
    let old_cell = workbook.sheet(sheet_idx).and_then(|s| s.get_cell(row, col));
    let old_value = old_cell
        .map(|c| c.value.clone())
        .unwrap_or(CellValue::Empty);
    let old_style = old_cell.and_then(|c| c.style.clone());
    let mut new_style = old_style.clone().unwrap_or_default();
    f(&mut new_style);
    let cmd = Box::new(SetCell {
        sheet: sheet_idx,
        row,
        col,
        old_value: old_value.clone(),
        new_value: old_value,
        old_style,
        new_style: Some(new_style),
    });
    match cmd.execute(workbook) {
        Ok(_) => {
            undo.push(cmd);
            set_status(status, msg.to_string());
        }
        Err(e) => set_status(status, format!("Error: {}", e)),
    }
}

fn handle_ex_command(
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
                workbook.sheets.remove(workbook.active_sheet);
                if workbook.active_sheet >= workbook.sheets.len() {
                    workbook.active_sheet = workbook.sheets.len() - 1;
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
                let fg = if color_is_dark(&bg) {
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
                // :note — show existing note in status bar
                if let Some(note) = sheet.notes.get(&(row, col)) {
                    set_status(status, format!("Note: {}", note));
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

fn set_status(status: &mut Option<(String, std::time::Instant)>, msg: String) {
    *status = Some((msg, std::time::Instant::now()));
}

/// Cycle a month or weekday name by `delta` steps (+1 or -1).
/// Returns `Some(new_text)` preserving the original's format (case + long/short),
/// or `None` if the text is not a recognised sequence member.
fn cycle_text_sequence(text: &str, delta: i32) -> Option<String> {
    const MONTHS_LONG: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    const MONTHS_SHORT: [&str; 12] = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];
    const DAYS_LONG: [&str; 7] = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ];
    const DAYS_SHORT: [&str; 7] = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

    // Detect the case style of the input so we can reproduce it.
    let apply_case = |template: &str| -> String {
        let all_upper = text.chars().all(|c| !c.is_alphabetic() || c.is_uppercase());
        let all_lower = text.chars().all(|c| !c.is_alphabetic() || c.is_lowercase());
        if all_upper {
            template.to_uppercase()
        } else if all_lower {
            template.to_lowercase()
        } else {
            template.to_string() // title-case (canonical)
        }
    };

    let lower = text.to_lowercase();

    // Try long month names first (e.g. "january"), then short (e.g. "jan").
    // Long must be checked first so "may" (== both long and short for May) resolves to long.
    // However "may" (3 chars) == MONTHS_SHORT "May" too — either gives the same result.
    for lists in [
        (&MONTHS_LONG[..], 12usize),
        (&MONTHS_SHORT[..], 12),
        (&DAYS_LONG[..], 7),
        (&DAYS_SHORT[..], 7),
    ] {
        let (names, len) = lists;
        if let Some(idx) = names.iter().position(|&n| n.to_lowercase() == lower) {
            let next = ((idx as i32 + delta).rem_euclid(len as i32)) as usize;
            return Some(apply_case(names[next]));
        }
    }
    None
}

fn days_in_month(month: i32, year: i32) -> i32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}

/// Add `delta` days to (day, month, year), returning the new (day, month, year).
fn add_days_to_date(mut day: i32, mut month: i32, mut year: i32, delta: i32) -> (i32, i32, i32) {
    day += delta;
    loop {
        if day < 1 {
            month -= 1;
            if month < 1 {
                month = 12;
                year -= 1;
            }
            day += days_in_month(month, year);
        } else if day > days_in_month(month, year) {
            day -= days_in_month(month, year);
            month += 1;
            if month > 12 {
                month = 1;
                year += 1;
            }
        } else {
            break;
        }
    }
    (day, month, year)
}

/// Try to parse, increment/decrement, and reformat a date string.
/// Supported formats (auto-detected):
///   DD.MM.YYYY  DD/MM/YYYY  DD-MM-YYYY
///   DD.MM.YY    DD/MM/YY    DD-MM-YY
///   YYYY-MM-DD  YYYY.MM.DD  YYYY/MM/DD
///   MM/DD/YYYY  (detected when first part ≤ 12 AND second part > 12)
fn cycle_date(text: &str, delta: i32) -> Option<String> {
    // Detect separator: must be exactly one type present
    let sep = if text.contains('.') && !text.contains('/') && !text.contains('-') {
        '.'
    } else if text.contains('/') && !text.contains('.') && !text.contains('-') {
        '/'
    } else if text.contains('-') && !text.contains('.') && !text.contains('/') {
        '-'
    } else {
        return None;
    };

    let parts: Vec<&str> = text.splitn(4, sep).collect();
    if parts.len() != 3 {
        return None;
    }
    // All parts must be purely numeric
    if parts.iter().any(|p| !p.chars().all(|c| c.is_ascii_digit())) {
        return None;
    }

    let n: Vec<i32> = parts.iter().filter_map(|p| p.parse().ok()).collect();
    if n.len() != 3 {
        return None;
    }

    // Determine layout: (day_idx, month_idx, year_idx, year_is_first)
    #[derive(Clone, Copy)]
    #[allow(clippy::upper_case_acronyms)]
    enum Layout {
        DMY,
        MDY,
        YMD,
    }

    let (layout, two_digit_year) = if parts[0].len() == 4 {
        // YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD
        (Layout::YMD, false)
    } else if parts[2].len() == 4 {
        // n[0] and n[1] both ≤ 12 → ambiguous → assume DD/MM/YYYY (European)
        // n[0] > 12 → must be day, so DD/MM
        // n[1] > 12 → n[0] must be month, so MM/DD (US)
        if n[0] <= 12 && n[1] > 12 {
            (Layout::MDY, false)
        } else {
            (Layout::DMY, false)
        }
    } else if parts[2].len() == 2 {
        (Layout::DMY, true)
    } else {
        return None;
    };

    let (day, month, full_year) = match layout {
        Layout::DMY => {
            let y = if two_digit_year {
                if n[2] < 70 {
                    2000 + n[2]
                } else {
                    1900 + n[2]
                }
            } else {
                n[2]
            };
            (n[0], n[1], y)
        }
        Layout::MDY => (n[1], n[0], n[2]),
        Layout::YMD => (n[2], n[1], n[0]),
    };

    // Basic validation
    if !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }

    let (nd, nm, ny) = add_days_to_date(day, month, full_year, delta);

    // Preserve original field widths (zero-padded to original length)
    let fmt_field =
        |val: i32, orig: &str| -> String { format!("{:0>width$}", val, width = orig.len()) };

    Some(match layout {
        Layout::DMY => format!(
            "{}{}{}{}{}",
            fmt_field(nd, parts[0]),
            sep,
            fmt_field(nm, parts[1]),
            sep,
            fmt_field(if two_digit_year { ny % 100 } else { ny }, parts[2])
        ),
        Layout::MDY => format!(
            "{}{}{}{}{}",
            fmt_field(nm, parts[0]),
            sep,
            fmt_field(nd, parts[1]),
            sep,
            fmt_field(ny, parts[2])
        ),
        Layout::YMD => format!(
            "{}{}{}{}{}",
            fmt_field(ny, parts[0]),
            sep,
            fmt_field(nm, parts[1]),
            sep,
            fmt_field(nd, parts[2])
        ),
    })
}

/// Parse an Excel-style cell address like "B15" or "AA3" into (row, col) 0-indexed.
fn parse_cell_address(addr: &str) -> Option<(u32, u32)> {
    let addr = addr.trim().to_uppercase();
    let col_end = addr.find(|c: char| c.is_ascii_digit())?;
    if col_end == 0 {
        return None;
    }
    let col_str = &addr[..col_end];
    let row_str = &addr[col_end..];
    let col = asat_core::letter_to_col(col_str)?;
    let row: u32 = row_str.parse::<u32>().ok()?.checked_sub(1)?;
    Some((row, col))
}

/// Parse a range address like "A1:C10" or "B5" into a CellRange (0-indexed).
fn parse_range_address(s: &str, sheet: usize) -> Option<asat_core::CellRange> {
    let s = s.trim().to_uppercase();
    if let Some(colon) = s.find(':') {
        let (a, b) = (&s[..colon], &s[colon + 1..]);
        let (r1, c1) = parse_cell_address(a)?;
        let (r2, c2) = parse_cell_address(b)?;
        Some(asat_core::CellRange::new(sheet, r1, c1, r2, c2))
    } else {
        let (r, c) = parse_cell_address(&s)?;
        Some(asat_core::CellRange::single(sheet, r, c))
    }
}

/// Check if a cell value matches a filter condition (op, val_str, val_num).
/// op can be: >, <, >=, <=, =, ==, <>, !=, or a plain substring match.
fn filter_row_matches(val: &CellValue, op: &str, val_str: &str, val_num: Option<f64>) -> bool {
    // Numeric comparison
    if let (Some(vn), Some(cn)) = (
        val_num,
        match val {
            CellValue::Number(n) => Some(*n),
            _ => None,
        },
    ) {
        return match op {
            ">" => cn > vn,
            "<" => cn < vn,
            ">=" => cn >= vn,
            "<=" => cn <= vn,
            "=" | "==" => (cn - vn).abs() < 1e-10,
            "<>" | "!=" => (cn - vn).abs() >= 1e-10,
            _ => false,
        };
    }
    // Text match
    let cell_text = val.display().to_lowercase();
    let needle = val_str.to_lowercase();
    match op {
        "=" | "==" => cell_text == needle,
        "<>" | "!=" => cell_text != needle,
        _ => cell_text.contains(&needle),
    }
}

/// Parse a range for CF rules, also accepting whole-column ranges like "A:A" or "A:C".
/// Falls back to parse_range_address, then tries to interpret a bare column letter range.
fn parse_range_address_cf(
    s: &str,
    sheet: usize,
    workbook: &Workbook,
) -> Option<asat_core::CellRange> {
    let upper = s.trim().to_uppercase();
    // Handle "A:A" or "A:C" — letter-only ranges meaning entire column(s)
    if let Some(colon) = upper.find(':') {
        let left = &upper[..colon];
        let right = &upper[colon + 1..];
        // Both sides pure letters → whole-column range
        if left.chars().all(|c| c.is_ascii_alphabetic())
            && right.chars().all(|c| c.is_ascii_alphabetic())
        {
            let c1 = asat_core::letter_to_col(left)?;
            let c2 = asat_core::letter_to_col(right)?;
            let max_row = workbook.active().max_row().max(9999);
            return Some(asat_core::CellRange::new(
                sheet,
                0,
                c1.min(c2),
                max_row,
                c1.max(c2),
            ));
        }
    }
    parse_range_address(s, sheet)
}

// ── Auto-fill series detection ────────────────────────────────────────────────

const WEEKDAYS: &[&str] = &["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const MONTHS: &[&str] = &[
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

fn weekday_index(s: &str) -> Option<usize> {
    let lower = s.to_ascii_lowercase();
    WEEKDAYS
        .iter()
        .position(|w| w.to_ascii_lowercase() == lower)
}

fn month_index(s: &str) -> Option<usize> {
    let lower = s.to_ascii_lowercase();
    MONTHS.iter().position(|m| m.to_ascii_lowercase() == lower)
}

/// Given a seed slice of CellValues, return a closure that produces
/// the i-th fill value (0 = first cell *after* the seed).
fn auto_fill_series(seed: &[CellValue]) -> Box<dyn Fn(usize) -> CellValue> {
    if seed.is_empty() {
        return Box::new(|_| CellValue::Empty);
    }

    // Try numeric arithmetic
    let nums: Option<Vec<f64>> = seed
        .iter()
        .map(|v| {
            if let CellValue::Number(n) = v {
                Some(*n)
            } else {
                None
            }
        })
        .collect();
    if let Some(ns) = nums {
        let start = ns[0];
        let step = if ns.len() >= 2 { ns[1] - ns[0] } else { 1.0 };
        let seed_len = ns.len() as f64;
        return Box::new(move |i| CellValue::Number(start + (seed_len + i as f64) * step));
    }

    // Try weekday cycle
    let wdays: Option<Vec<usize>> = seed
        .iter()
        .map(|v| {
            if let CellValue::Text(s) = v {
                weekday_index(s)
            } else {
                None
            }
        })
        .collect();
    if let Some(wd) = wdays {
        let seed_len = wd.len();
        let first = wd[0];
        return Box::new(move |i| {
            let idx = (first + seed_len + i) % WEEKDAYS.len();
            CellValue::Text(WEEKDAYS[idx].to_string())
        });
    }

    // Try month cycle
    let mths: Option<Vec<usize>> = seed
        .iter()
        .map(|v| {
            if let CellValue::Text(s) = v {
                month_index(s)
            } else {
                None
            }
        })
        .collect();
    if let Some(ms) = mths {
        let seed_len = ms.len();
        let first = ms[0];
        return Box::new(move |i| {
            let idx = (first + seed_len + i) % MONTHS.len();
            CellValue::Text(MONTHS[idx].to_string())
        });
    }

    // Fallback: cycle through seed values
    let owned: Vec<CellValue> = seed.to_vec();
    let seed_len = owned.len();
    Box::new(move |i| owned[(seed_len + i) % seed_len].clone())
}

/// Keep `input.subcmd_completions` in sync with the current command buffer.
/// Called after every key event so Tab can cycle sub-command arguments.
fn update_subcmd_completions(input: &mut InputState) {
    if !matches!(input.mode, Mode::Command) {
        return;
    }
    let buf = input.command_buffer.clone();
    if let Some(space_idx) = buf.find(' ') {
        let verb = buf[..space_idx].to_ascii_lowercase();
        let arg_prefix = buf[space_idx + 1..].to_ascii_lowercase();

        match verb.as_str() {
            "theme" => {
                let themes = asat_config::builtin_themes();
                input.subcmd_completions = themes
                    .iter()
                    .filter(|t| {
                        let n = t.name.to_ascii_lowercase();
                        let id = t.id.to_ascii_lowercase();
                        arg_prefix.is_empty()
                            || n.starts_with(&arg_prefix)
                            || id.starts_with(&arg_prefix)
                            || n.contains(&arg_prefix)
                    })
                    .map(|t| t.name.to_string())
                    .collect();
            }
            _ => input.subcmd_completions.clear(),
        }
    } else {
        input.subcmd_completions.clear();
        input.subcmd_completion_idx = None;
    }
}

// ── Recent files I/O ─────────────────────────────────────────────────────────

fn recent_files_path() -> PathBuf {
    dirs::data_local_dir()
        .unwrap_or_else(|| {
            PathBuf::from(std::env::var("HOME").unwrap_or_default()).join(".local/share")
        })
        .join("asat")
        .join("recent")
}

fn load_recent_files(limit: usize) -> Vec<String> {
    let path = recent_files_path();
    std::fs::read_to_string(path)
        .unwrap_or_default()
        .lines()
        .filter(|l| !l.trim().is_empty())
        .map(|l| l.to_string())
        .take(limit.max(1))
        .collect()
}

fn save_recent_files(files: &[String]) {
    let path = recent_files_path();
    if let Some(dir) = path.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    let content = files.join("\n");
    let _ = std::fs::write(path, content);
}

/// Prepend `path` to the recent list (dedup + cap at 20).
fn push_recent(path: &str, recents: &mut Vec<String>) {
    recents.retain(|r| r != path);
    recents.insert(0, path.to_string());
    recents.truncate(20);
}

// ── File scanning (for fuzzy finder) ─────────────────────────────────────────

fn scan_files(root: &PathBuf) -> Vec<String> {
    let mut results = Vec::new();
    scan_dir(root, root, 5, &mut results);
    results.sort();
    results.truncate(2000);
    results
}

fn scan_dir(root: &PathBuf, dir: &PathBuf, depth: usize, results: &mut Vec<String>) {
    if depth == 0 {
        return;
    }
    let entries = match std::fs::read_dir(dir) {
        Ok(e) => e,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("");
        // Skip hidden files/dirs and common build directories
        if name.starts_with('.') {
            continue;
        }
        if matches!(
            name,
            "target" | "node_modules" | "__pycache__" | ".git" | "dist" | "build"
        ) {
            continue;
        }
        if path.is_dir() {
            scan_dir(root, &path, depth - 1, results);
        } else if let Ok(rel) = path.strip_prefix(root) {
            results.push(rel.to_string_lossy().to_string());
        }
    }
}

// ── Config editor ─────────────────────────────────────────────────────────────

fn open_config_in_editor() {
    use crossterm::execute;
    use crossterm::terminal::{disable_raw_mode, enable_raw_mode};

    let config_path = {
        let base = std::env::var("XDG_CONFIG_HOME")
            .map(PathBuf::from)
            .unwrap_or_else(|_| {
                dirs::home_dir()
                    .unwrap_or_else(|| PathBuf::from("."))
                    .join(".config")
            });
        base.join("asat").join("config.toml")
    };

    // Create default config if it doesn't exist yet
    if !config_path.exists() {
        let _ = asat_config::Config::write_default();
    }

    let editor = std::env::var("EDITOR")
        .or_else(|_| std::env::var("VISUAL"))
        .unwrap_or_else(|_| "nano".to_string());

    // Temporarily restore the terminal for the external editor
    let _ = disable_raw_mode();
    let _ = execute!(std::io::stdout(), crossterm::terminal::LeaveAlternateScreen);

    let _ = std::process::Command::new(&editor)
        .arg(&config_path)
        .status();

    // Restore the TUI
    let _ = enable_raw_mode();
    let _ = execute!(std::io::stdout(), crossterm::terminal::EnterAlternateScreen);
}

// ── Go-to definition helpers ──────────────────────────────────────────────────

/// Parse a formula string (without leading '=') and return the first cell reference:
/// `(sheet_name: Option<String>, row: u32, col: u32)`
fn find_first_cell_ref(formula: &str) -> Option<(Option<String>, u32, u32)> {
    let tokens = asat_formula::lexer::lex(formula).ok()?;
    let expr = asat_formula::parser::parse(&tokens).ok()?;
    find_first_ref_in_expr(&expr)
}

fn find_first_ref_in_expr(expr: &asat_formula::parser::Expr) -> Option<(Option<String>, u32, u32)> {
    use asat_formula::parser::Expr;
    match expr {
        Expr::CellRef { sheet, row, col } => Some((sheet.clone(), *row, *col)),
        Expr::RangeRef {
            sheet, row1, col1, ..
        } => Some((sheet.clone(), *row1, *col1)),
        Expr::UnaryMinus(e) | Expr::UnaryPlus(e) => find_first_ref_in_expr(e),
        Expr::Add(a, b)
        | Expr::Sub(a, b)
        | Expr::Mul(a, b)
        | Expr::Div(a, b)
        | Expr::Pow(a, b)
        | Expr::Concat(a, b)
        | Expr::Eq(a, b)
        | Expr::Neq(a, b)
        | Expr::Lt(a, b)
        | Expr::Lte(a, b)
        | Expr::Gt(a, b)
        | Expr::Gte(a, b) => find_first_ref_in_expr(a).or_else(|| find_first_ref_in_expr(b)),
        Expr::Call { args, .. } => args.iter().find_map(find_first_ref_in_expr),
        _ => None,
    }
}
