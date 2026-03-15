pub mod themes;
pub use themes::{builtin_themes, ThemePreset};

use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum ConfigError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("TOML parse error: {0}")]
    Toml(#[from] toml::de::Error),
}

// ── Top-level config ──────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct Config {
    // ── Theme ─────────────────────────────────────────────────────────────────
    /// Name of a built-in theme preset (e.g. "nord", "dracula", "gruvbox-dark").
    /// When set, the [theme] section below is ignored.
    /// Leave empty or set to "custom" to use the [theme] colors directly.
    /// Run :theme inside ASAT to browse and apply themes interactively.
    pub theme_name: String,

    /// Custom colors — only active when theme_name is "" or "custom".
    /// Each field accepts a hex color string like "#268BD2".
    pub theme: ThemeConfig,

    // ── Display ───────────────────────────────────────────────────────────────
    /// Default width (in characters) for newly created columns.
    pub default_col_width: u16,

    /// Minimum column width; columns cannot be narrowed below this.
    pub min_col_width: u16,

    /// Maximum column width when using auto-fit (= key).
    pub max_col_width: u16,

    /// Number of rows kept visible above/below the cursor when scrolling.
    pub scroll_padding: u32,

    /// Show absolute row numbers in the gutter. (relative_line_numbers overrides this.)
    pub show_line_numbers: bool,

    /// Show Vim-style relative row numbers (distance from cursor row).
    pub relative_line_numbers: bool,

    /// Highlight the entire row the cursor is on with a subtle tint.
    pub highlight_cursor_row: bool,

    /// Highlight the entire column the cursor is on with a subtle tint.
    pub highlight_cursor_col: bool,

    /// Show the formula bar above the grid.
    pub show_formula_bar: bool,

    /// Show the sheet tab bar at the top.
    pub show_tab_bar: bool,

    /// Show the status bar at the bottom.
    pub show_status_bar: bool,

    /// Seconds before a status-bar message fades out (0 = never fade).
    pub status_timeout: u32,

    // ── Editing ───────────────────────────────────────────────────────────────
    /// Maximum number of undoable operations kept in history.
    pub undo_limit: usize,

    /// Seconds between auto-saves when the workbook has unsaved changes (0 = disabled).
    pub autosave_interval: u32,

    /// Create a .bak backup of the previous file before overwriting on save.
    pub backup_on_save: bool,

    /// Ask for confirmation before destructive operations like dd or :dr.
    pub confirm_delete: bool,

    /// Wrap cursor movement at the edges of the sheet (left↔right, up↔down).
    pub wrap_navigation: bool,

    // ── Number formatting ─────────────────────────────────────────────────────
    /// Maximum decimal places shown when displaying floating-point numbers
    /// that have no explicit number format applied (0–15).
    pub number_precision: u8,

    /// Default date display format for date-serial numbers.
    /// Tokens: YYYY MM DD HH mm ss  (e.g. "YYYY-MM-DD", "DD/MM/YYYY")
    pub date_format: String,

    // ── Files ─────────────────────────────────────────────────────────────────
    /// Default file format used when saving a file that has no extension.
    /// One of: "csv", "tsv", "xlsx", "ods", "asat"
    pub default_format: String,

    /// Field separator character for CSV files (single character).
    pub csv_delimiter: String,

    /// Number of recently opened files to remember and show on the welcome screen.
    pub remember_recent: u32,
}

impl Default for Config {
    fn default() -> Self {
        Config {
            // Theme
            theme_name: "solarized-dark".to_string(),
            theme: ThemeConfig::default(),
            // Display
            default_col_width: 10,
            min_col_width: 3,
            max_col_width: 60,
            scroll_padding: 3,
            show_line_numbers: false,
            relative_line_numbers: false,
            highlight_cursor_row: false,
            highlight_cursor_col: false,
            show_formula_bar: true,
            show_tab_bar: true,
            show_status_bar: true,
            status_timeout: 3,
            // Editing
            undo_limit: 1000,
            autosave_interval: 0,
            backup_on_save: false,
            confirm_delete: false,
            wrap_navigation: false,
            // Number formatting
            number_precision: 6,
            date_format: "YYYY-MM-DD".to_string(),
            // Files
            default_format: "csv".to_string(),
            csv_delimiter: ",".to_string(),
            remember_recent: 20,
        }
    }
}

// ── ThemeConfig ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct ThemeConfig {
    pub cursor_bg: String,
    pub cursor_fg: String,
    pub header_bg: String,
    pub header_fg: String,
    /// Main spreadsheet cell background colour
    pub cell_bg: String,
    pub selection_bg: String,
    /// Colour used for numeric/boolean cell values
    pub number_color: String,
    pub normal_mode_color: String,
    pub insert_mode_color: String,
    pub visual_mode_color: String,
    pub command_mode_color: String,
}

impl Default for ThemeConfig {
    fn default() -> Self {
        // Matches the solarized-dark preset
        ThemeConfig {
            cursor_bg: "#268BD2".to_string(),
            cursor_fg: "#000000".to_string(),
            header_bg: "#073642".to_string(),
            header_fg: "#93A1A1".to_string(),
            cell_bg: "#002B36".to_string(),
            selection_bg: "#2AA198".to_string(),
            number_color: "#2AA198".to_string(),
            normal_mode_color: "#859900".to_string(),
            insert_mode_color: "#268BD2".to_string(),
            visual_mode_color: "#6C71C4".to_string(),
            command_mode_color: "#CB4B16".to_string(),
        }
    }
}

// ── Load / Save ───────────────────────────────────────────────────────────────

impl Config {
    /// Load config from ~/.config/asat/config.toml, returning default if absent.
    /// If `theme_name` matches a built-in preset, that preset's colors are applied
    /// automatically — no need to keep hex values in the file.
    pub fn load() -> Result<Self, ConfigError> {
        let path = config_path();
        if !path.exists() {
            return Ok(Config::default());
        }
        let content = std::fs::read_to_string(&path)?;
        let mut config: Config = toml::from_str(&content)?;

        // Apply a named preset if theme_name is set
        config.apply_theme_preset();

        Ok(config)
    }

    /// If `theme_name` matches a built-in preset, overwrite `self.theme` with it.
    pub fn apply_theme_preset(&mut self) {
        if self.theme_name.is_empty() || self.theme_name == "custom" {
            return;
        }
        let themes = builtin_themes();
        if let Some(preset) = themes.iter().find(|t| {
            t.id.eq_ignore_ascii_case(&self.theme_name)
                || t.name.eq_ignore_ascii_case(&self.theme_name)
        }) {
            self.theme = preset.config.clone();
        }
    }

    /// Save the current config to disk, creating directories as needed.
    /// When a named preset is active the [theme] color block is omitted to keep
    /// the file clean — the preset name alone is enough to restore the colors.
    pub fn save(&self) -> Result<(), ConfigError> {
        let path = config_path();
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        let content = self.to_toml_string();
        std::fs::write(&path, content)?;
        Ok(())
    }

    /// Write the default config to disk.
    pub fn write_default() -> Result<(), ConfigError> {
        Config::default().save()
    }

    /// Produce a readable, well-commented TOML string.
    /// The [theme] block is omitted when a named preset is active.
    fn to_toml_string(&self) -> String {
        let using_preset = !self.theme_name.is_empty() && self.theme_name != "custom";
        let mut s = String::new();

        s.push_str("# ASAT configuration file\n");
        s.push_str("# Run :theme inside ASAT to browse themes interactively.\n");
        s.push_str(
            "# All fields are optional — missing ones use the default shown in comments.\n\n",
        );

        // Theme
        s.push_str("# ── Theme ────────────────────────────────────────────────────────────────\n");
        s.push_str("# Built-in themes: solarized-dark, solarized-light, nord, dracula,\n");
        s.push_str("#   gruvbox-dark, gruvbox-light, tokyo-night, catppuccin-mocha,\n");
        s.push_str("#   catppuccin-latte, one-dark, monokai, rose-pine, everforest-dark,\n");
        s.push_str("#   kanagawa, cyberpunk, amber-terminal, ice, github-dark\n");
        s.push_str("# Set to \"custom\" and edit the [theme] section below for full control.\n");
        s.push_str(&format!("theme_name = {:?}\n\n", self.theme_name));

        // Display
        s.push_str("# ── Display ──────────────────────────────────────────────────────────────\n");
        s.push_str(&format!(
            "default_col_width    = {}   # default width for new columns (chars)\n",
            self.default_col_width
        ));
        s.push_str(&format!(
            "min_col_width        = {}    # minimum column width (chars)\n",
            self.min_col_width
        ));
        s.push_str(&format!(
            "max_col_width        = {}   # maximum width when auto-fitting (= key)\n",
            self.max_col_width
        ));
        s.push_str(&format!(
            "scroll_padding       = {}    # rows kept visible above/below cursor\n",
            self.scroll_padding
        ));
        s.push_str(&format!(
            "show_line_numbers    = {}  # show row numbers in the gutter\n",
            self.show_line_numbers
        ));
        s.push_str(&format!(
            "relative_line_numbers = {}  # Vim-style relative row numbers\n",
            self.relative_line_numbers
        ));
        s.push_str(&format!(
            "highlight_cursor_row = {}  # tint the cursor row\n",
            self.highlight_cursor_row
        ));
        s.push_str(&format!(
            "highlight_cursor_col = {}  # tint the cursor column\n",
            self.highlight_cursor_col
        ));
        s.push_str(&format!(
            "show_formula_bar     = {}  # show the formula bar\n",
            self.show_formula_bar
        ));
        s.push_str(&format!(
            "show_tab_bar         = {}  # show the sheet tab bar\n",
            self.show_tab_bar
        ));
        s.push_str(&format!(
            "show_status_bar      = {}  # show the bottom status bar\n",
            self.show_status_bar
        ));
        s.push_str(&format!(
            "status_timeout       = {}    # seconds before status message fades (0 = never)\n\n",
            self.status_timeout
        ));

        // Editing
        s.push_str("# ── Editing ──────────────────────────────────────────────────────────────\n");
        s.push_str(&format!(
            "undo_limit           = {}  # maximum undo history depth\n",
            self.undo_limit
        ));
        s.push_str(&format!(
            "autosave_interval    = {}    # seconds between auto-saves (0 = disabled)\n",
            self.autosave_interval
        ));
        s.push_str(&format!(
            "backup_on_save       = {}  # write a .bak file before overwriting\n",
            self.backup_on_save
        ));
        s.push_str(&format!(
            "confirm_delete       = {}  # ask before dd / :dr / :dc\n",
            self.confirm_delete
        ));
        s.push_str(&format!(
            "wrap_navigation      = {}  # wrap cursor at sheet edges\n\n",
            self.wrap_navigation
        ));

        // Number formatting
        s.push_str("# ── Number formatting ────────────────────────────────────────────────────\n");
        s.push_str(&format!(
            "number_precision     = {}    # max decimal places for unformatted numbers (0–15)\n",
            self.number_precision
        ));
        s.push_str(&format!(
            "date_format          = {:?}  # date display format (YYYY MM DD HH mm ss)\n\n",
            self.date_format
        ));

        // Files
        s.push_str("# ── Files ────────────────────────────────────────────────────────────────\n");
        s.push_str(&format!(
            "default_format       = {:?}    # fallback format when saving without extension\n",
            self.default_format
        ));
        s.push_str(&format!(
            "csv_delimiter        = {:?}      # CSV field separator\n",
            self.csv_delimiter
        ));
        s.push_str(&format!(
            "remember_recent      = {}   # recent files shown on welcome screen\n\n",
            self.remember_recent
        ));

        // [theme] block — only when using custom colors
        if !using_preset {
            s.push_str(
                "# ── Custom colors (active when theme_name = \"custom\" or \"\") ────────────\n",
            );
            s.push_str("[theme]\n");
            s.push_str(&format!(
                "cursor_bg          = {:?}\n",
                self.theme.cursor_bg
            ));
            s.push_str(&format!(
                "cursor_fg          = {:?}\n",
                self.theme.cursor_fg
            ));
            s.push_str(&format!(
                "header_bg          = {:?}\n",
                self.theme.header_bg
            ));
            s.push_str(&format!(
                "header_fg          = {:?}\n",
                self.theme.header_fg
            ));
            s.push_str(&format!("cell_bg            = {:?}\n", self.theme.cell_bg));
            s.push_str(&format!(
                "selection_bg       = {:?}\n",
                self.theme.selection_bg
            ));
            s.push_str(&format!(
                "number_color       = {:?}\n",
                self.theme.number_color
            ));
            s.push_str(&format!(
                "normal_mode_color  = {:?}\n",
                self.theme.normal_mode_color
            ));
            s.push_str(&format!(
                "insert_mode_color  = {:?}\n",
                self.theme.insert_mode_color
            ));
            s.push_str(&format!(
                "visual_mode_color  = {:?}\n",
                self.theme.visual_mode_color
            ));
            s.push_str(&format!(
                "command_mode_color = {:?}\n",
                self.theme.command_mode_color
            ));
        }

        s
    }
}

// ── Filesystem helpers ────────────────────────────────────────────────────────

fn config_path() -> PathBuf {
    config_dir().join("asat/config.toml")
}

fn config_dir() -> PathBuf {
    std::env::var("XDG_CONFIG_HOME")
        .map(PathBuf::from)
        .unwrap_or_else(|_| {
            dirs::home_dir()
                .unwrap_or_else(|| PathBuf::from("."))
                .join(".config")
        })
}
