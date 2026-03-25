mod app;
mod ex_commands;
mod process_action;

use std::io;
use std::path::PathBuf;
use std::time::Duration;

use anyhow::Result;
use arboard::Clipboard;
use crossterm::{
    event::{self, Event},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{backend::CrosstermBackend, Terminal};

use asat_commands::{Command, SetCell, UndoStack};
use asat_config::Config;
use asat_core::Workbook;
use asat_input::{InputState, Mode};
use asat_plugins::{PluginEvent, PluginManager, PluginOutput};
use asat_tui::{render, RenderState};

use app::ActionResult;
use app::{
    delete_swap, load_recent_files, recalculate_all, set_status, swap_path,
    update_subcmd_completions, visible_cols_in_width, visible_rows_in_height, write_swap,
};
use ex_commands::handle_ex_command;
use process_action::process_action;

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
    // Parse-once cache: formula string → AST. Avoids re-lexing/parsing on every frame.
    let mut ast_cache: std::collections::HashMap<String, asat_formula::Expr> =
        std::collections::HashMap::new();
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
        recalculate_all(&mut workbook, &mut ast_cache);

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
            if key.kind != crossterm::event::KeyEventKind::Press {
                continue;
            }
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
