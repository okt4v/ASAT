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

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct Config {
    pub default_col_width: u16,
    pub autosave_interval: u32,   // edits between autosaves
    pub theme: ThemeConfig,
    pub scroll_padding: u32,
    pub show_line_numbers: bool,
}

impl Default for Config {
    fn default() -> Self {
        Config {
            default_col_width: 10,
            autosave_interval: 50,
            theme: ThemeConfig::default(),
            scroll_padding: 3,
            show_line_numbers: false,
        }
    }
}

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
        ThemeConfig {
            cursor_bg:         "#268BD2".to_string(),
            cursor_fg:         "#FFFFFF".to_string(),
            header_bg:         "#073642".to_string(),
            header_fg:         "#93A1A1".to_string(),
            cell_bg:           "#002B36".to_string(),
            selection_bg:      "#2AA198".to_string(),
            number_color:      "#2AA198".to_string(),
            normal_mode_color: "#859900".to_string(),
            insert_mode_color: "#268BD2".to_string(),
            visual_mode_color: "#6C71C4".to_string(),
            command_mode_color:"#CB4B16".to_string(),
        }
    }
}

impl Config {
    /// Load config from ~/.config/asat/config.toml, returning default if absent
    pub fn load() -> Result<Self, ConfigError> {
        let path = config_path();
        if !path.exists() {
            return Ok(Config::default());
        }
        let content = std::fs::read_to_string(&path)?;
        let config: Config = toml::from_str(&content)?;
        Ok(config)
    }

    /// Save the current config to disk (creates directories if needed).
    pub fn save(&self) -> Result<(), ConfigError> {
        let path = config_path();
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        let content = toml::to_string_pretty(self)
            .map_err(|e| ConfigError::Io(std::io::Error::new(std::io::ErrorKind::Other, e.to_string())))?;
        std::fs::write(&path, content)?;
        Ok(())
    }

    /// Write the default config to disk (creates directories if needed)
    pub fn write_default() -> Result<(), ConfigError> {
        let path = config_path();
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        let content = toml::to_string_pretty(&Config::default())
            .map_err(|e| ConfigError::Io(std::io::Error::new(std::io::ErrorKind::Other, e.to_string())))?;
        std::fs::write(&path, content)?;
        Ok(())
    }
}

fn config_path() -> PathBuf {
    dirs_next().join("asat/config.toml")
}

fn dirs_next() -> PathBuf {
    // Try XDG_CONFIG_HOME first, then ~/.config
    std::env::var("XDG_CONFIG_HOME")
        .map(PathBuf::from)
        .unwrap_or_else(|_| {
            dirs::home_dir()
                .unwrap_or_else(|| PathBuf::from("."))
                .join(".config")
        })
}
