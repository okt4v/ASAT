pub mod command_line;
pub mod completion;
pub mod formula_bar;
pub mod formula_hint;
pub mod grid;
pub mod notification;
pub mod status_bar;
pub mod tab_bar;
pub mod theme_manager;
pub mod welcome;
pub mod whichkey;

use asat_config::Config;
use asat_core::Workbook;
use asat_input::{InputState, Mode};
use ratatui::{
    layout::{Constraint, Direction, Layout},
    Frame,
};

pub struct RenderState<'a> {
    pub workbook: &'a Workbook,
    pub input: &'a InputState,
    pub status_message: Option<&'a str>,
    pub show_side_panel: bool,
    pub config: &'a Config,
}

/// Parse a `#RRGGBB` hex string into a ratatui Color. Falls back to Reset on parse failure.
pub fn parse_hex_color(hex: &str) -> ratatui::style::Color {
    let h = hex.trim_start_matches('#');
    if h.len() == 6 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&h[0..2], 16),
            u8::from_str_radix(&h[2..4], 16),
            u8::from_str_radix(&h[4..6], 16),
        ) {
            return ratatui::style::Color::Rgb(r, g, b);
        }
    }
    ratatui::style::Color::Reset
}

/// Darken a ratatui Color by multiplying each RGB channel by `factor` (0.0–1.0).
pub fn darken(color: ratatui::style::Color, factor: f32) -> ratatui::style::Color {
    match color {
        ratatui::style::Color::Rgb(r, g, b) => ratatui::style::Color::Rgb(
            (r as f32 * factor) as u8,
            (g as f32 * factor) as u8,
            (b as f32 * factor) as u8,
        ),
        c => c,
    }
}

/// Returns true if the colour is perceptually dark (useful for choosing fg contrast).
pub fn is_dark_color(color: ratatui::style::Color) -> bool {
    match color {
        ratatui::style::Color::Rgb(r, g, b) => {
            (r as u32 * 299 + g as u32 * 587 + b as u32 * 114) < 128_000
        }
        _ => true,
    }
}

pub fn render(frame: &mut Frame, state: &RenderState<'_>) {
    let area = frame.area();
    let mode = &state.input.mode;
    let show_command = matches!(mode, Mode::Command | Mode::Search { .. });
    let is_special = matches!(
        mode,
        Mode::Welcome | Mode::FileFind | Mode::RecentFiles | Mode::ThemeManager
    );

    if is_special {
        // Special screens: just the content area + status bar
        let rows = Layout::default()
            .direction(Direction::Vertical)
            .constraints([Constraint::Min(3), Constraint::Length(1)])
            .split(area);

        match mode {
            Mode::Welcome => welcome::render_welcome(frame, rows[0], state),
            Mode::FileFind => {
                welcome::render_welcome(frame, rows[0], state);
                welcome::render_file_finder(frame, rows[0], state);
            }
            Mode::RecentFiles => {
                welcome::render_welcome(frame, rows[0], state);
                welcome::render_recent_files(frame, rows[0], state);
            }
            Mode::ThemeManager => theme_manager::render(frame, rows[0], state),
            _ => {}
        }

        status_bar::render(frame, rows[1], state);
        notification::render(frame, area, state);
        return;
    }

    // ── Normal spreadsheet layout ─────────────────────────────────────────
    let row_constraints = if show_command {
        vec![
            Constraint::Length(1),
            Constraint::Min(3),
            Constraint::Length(1),
            Constraint::Length(1),
            Constraint::Length(1),
        ]
    } else {
        vec![
            Constraint::Length(1),
            Constraint::Min(3),
            Constraint::Length(1),
            Constraint::Length(1),
        ]
    };

    let rows = Layout::default()
        .direction(Direction::Vertical)
        .constraints(row_constraints)
        .split(area);

    formula_bar::render(frame, rows[0], state);
    grid::render(frame, rows[1], state);
    tab_bar::render(frame, rows[2], state);
    status_bar::render(frame, rows[3], state);

    if show_command {
        command_line::render(frame, rows[4], state);
        completion::render(frame, area, state);
    }

    // Formula hint popup — shown when editing a formula cell.
    formula_hint::render(frame, rows[1], state);

    // Which-key overlay — rendered last so it floats above everything.
    whichkey::render(frame, rows[1], state);

    // Notification bubble — always on top, regardless of mode.
    notification::render(frame, area, state);
}
