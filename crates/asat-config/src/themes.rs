use crate::ThemeConfig;

pub struct ThemePreset {
    pub id:          &'static str,
    pub name:        &'static str,
    pub dark:        bool,
    pub description: &'static str,
    pub config:      ThemeConfig,
}

macro_rules! theme {
    (
        id: $id:expr, name: $name:expr, dark: $dark:expr, desc: $desc:expr,
        cursor_bg:    $cursor_bg:expr,
        header_bg:    $header_bg:expr,
        header_fg:    $header_fg:expr,
        cell_bg:      $cell_bg:expr,
        selection_bg: $sel_bg:expr,
        number_color: $num:expr,
        normal:       $normal:expr,
        insert:       $insert:expr,
        visual:       $visual:expr,
        command:      $command:expr,
    ) => {
        ThemePreset {
            id: $id, name: $name, dark: $dark, description: $desc,
            config: ThemeConfig {
                cursor_bg:          $cursor_bg.to_string(),
                cursor_fg:          "#000000".to_string(),
                header_bg:          $header_bg.to_string(),
                header_fg:          $header_fg.to_string(),
                cell_bg:            $cell_bg.to_string(),
                selection_bg:       $sel_bg.to_string(),
                number_color:       $num.to_string(),
                normal_mode_color:  $normal.to_string(),
                insert_mode_color:  $insert.to_string(),
                visual_mode_color:  $visual.to_string(),
                command_mode_color: $command.to_string(),
            },
        }
    };
}

pub fn builtin_themes() -> Vec<ThemePreset> {
    vec![
        theme! {
            id: "solarized-dark", name: "Solarized Dark", dark: true,
            desc: "The classic Solarized dark palette — precise, low-contrast, easy on the eyes.",
            cursor_bg:    "#268BD2", header_bg: "#073642", header_fg: "#93A1A1",
            cell_bg:      "#002B36", selection_bg: "#2AA198", number_color: "#2AA198",
            normal: "#859900", insert: "#268BD2", visual: "#6C71C4", command: "#CB4B16",
        },
        theme! {
            id: "solarized-light", name: "Solarized Light", dark: false,
            desc: "Solarized light — the same precise palette, inverted for bright environments.",
            cursor_bg:    "#268BD2", header_bg: "#EEE8D5", header_fg: "#839496",
            cell_bg:      "#FDF6E3", selection_bg: "#2AA198", number_color: "#268BD2",
            normal: "#859900", insert: "#268BD2", visual: "#6C71C4", command: "#CB4B16",
        },
        theme! {
            id: "nord", name: "Nord", dark: true,
            desc: "Arctic, north-bluish colour palette. Clean and professional.",
            cursor_bg:    "#88C0D0", header_bg: "#2E3440", header_fg: "#4C566A",
            cell_bg:      "#242933", selection_bg: "#434C5E", number_color: "#88C0D0",
            normal: "#A3BE8C", insert: "#88C0D0", visual: "#B48EAD", command: "#D08770",
        },
        theme! {
            id: "dracula", name: "Dracula", dark: true,
            desc: "A dark theme with vibrant, candy-like colours. Hugely popular.",
            cursor_bg:    "#BD93F9", header_bg: "#21222C", header_fg: "#6272A4",
            cell_bg:      "#282A36", selection_bg: "#44475A", number_color: "#8BE9FD",
            normal: "#50FA7B", insert: "#BD93F9", visual: "#FF79C6", command: "#FF5555",
        },
        theme! {
            id: "gruvbox-dark", name: "Gruvbox Dark", dark: true,
            desc: "Warm retro-groove colours. High contrast, comfortable for long sessions.",
            cursor_bg:    "#FE8019", header_bg: "#1D2021", header_fg: "#928374",
            cell_bg:      "#282828", selection_bg: "#504945", number_color: "#83A598",
            normal: "#B8BB26", insert: "#83A598", visual: "#D3869B", command: "#FB4934",
        },
        theme! {
            id: "gruvbox-light", name: "Gruvbox Light", dark: false,
            desc: "Gruvbox in light mode — warm cream tones with strong contrast.",
            cursor_bg:    "#D65D0E", header_bg: "#EBDBB2", header_fg: "#928374",
            cell_bg:      "#FBF1C7", selection_bg: "#D5C4A1", number_color: "#076678",
            normal: "#79740E", insert: "#076678", visual: "#8F3F71", command: "#CC241D",
        },
        theme! {
            id: "tokyo-night", name: "Tokyo Night", dark: true,
            desc: "Inspired by the neon lights of Tokyo. Deep blue with vivid accents.",
            cursor_bg:    "#7AA2F7", header_bg: "#16161E", header_fg: "#414868",
            cell_bg:      "#1A1B26", selection_bg: "#283457", number_color: "#73DACA",
            normal: "#9ECE6A", insert: "#7AA2F7", visual: "#BB9AF7", command: "#F7768E",
        },
        theme! {
            id: "catppuccin-mocha", name: "Catppuccin Mocha", dark: true,
            desc: "Soothing pastel colours on a rich dark background. Warm and modern.",
            cursor_bg:    "#89B4FA", header_bg: "#181825", header_fg: "#585B70",
            cell_bg:      "#1E1E2E", selection_bg: "#313244", number_color: "#94E2D5",
            normal: "#A6E3A1", insert: "#89B4FA", visual: "#CBA6F7", command: "#F38BA8",
        },
        theme! {
            id: "catppuccin-latte", name: "Catppuccin Latte", dark: false,
            desc: "Catppuccin's light variant — delicate pastels on a clean cream background.",
            cursor_bg:    "#1E66F5", header_bg: "#E6E9EF", header_fg: "#8C8FA1",
            cell_bg:      "#EFF1F5", selection_bg: "#DCE0E8", number_color: "#179299",
            normal: "#40A02B", insert: "#1E66F5", visual: "#8839EF", command: "#D20F39",
        },
        theme! {
            id: "one-dark", name: "One Dark", dark: true,
            desc: "Atom's iconic One Dark theme. Balanced contrast and rich syntax colours.",
            cursor_bg:    "#61AFEF", header_bg: "#21252B", header_fg: "#5C6370",
            cell_bg:      "#282C34", selection_bg: "#3E4451", number_color: "#56B6C2",
            normal: "#98C379", insert: "#61AFEF", visual: "#C678DD", command: "#E06C75",
        },
        theme! {
            id: "monokai", name: "Monokai", dark: true,
            desc: "The classic Monokai palette from Sublime Text. Timeless and punchy.",
            cursor_bg:    "#A6E22E", header_bg: "#1E1F1C", header_fg: "#75715E",
            cell_bg:      "#272822", selection_bg: "#49483E", number_color: "#66D9E8",
            normal: "#A6E22E", insert: "#66D9E8", visual: "#AE81FF", command: "#F92672",
        },
        theme! {
            id: "rose-pine", name: "Rosé Pine", dark: true,
            desc: "All natural pine, faux fur and new-age feel. Muted and romantic.",
            cursor_bg:    "#EBBCBA", header_bg: "#1F1D2E", header_fg: "#6E6A86",
            cell_bg:      "#191724", selection_bg: "#26233A", number_color: "#9CCFD8",
            normal: "#9CCFD8", insert: "#EBBCBA", visual: "#C4A7E7", command: "#EB6F92",
        },
        theme! {
            id: "everforest-dark", name: "Everforest Dark", dark: true,
            desc: "Green-tinted dark theme inspired by forests. Natural and restful.",
            cursor_bg:    "#A7C080", header_bg: "#272E33", header_fg: "#7A8478",
            cell_bg:      "#2D353B", selection_bg: "#3D484D", number_color: "#7FBBB3",
            normal: "#A7C080", insert: "#7FBBB3", visual: "#D699B6", command: "#E67E80",
        },
        theme! {
            id: "kanagawa", name: "Kanagawa Wave", dark: true,
            desc: "Inspired by the colours of feudal Japan. Deep ink blues and muted greens.",
            cursor_bg:    "#7E9CD8", header_bg: "#16161D", header_fg: "#54546D",
            cell_bg:      "#1F1F28", selection_bg: "#2D4F67", number_color: "#6A9589",
            normal: "#98BB6C", insert: "#7E9CD8", visual: "#957FB8", command: "#C34043",
        },
        theme! {
            id: "cyberpunk", name: "Cyberpunk", dark: true,
            desc: "Neon on black. Electric cyan, acid green, hot pink. The future is now.",
            cursor_bg:    "#00FFFF", header_bg: "#060610", header_fg: "#2A2A55",
            cell_bg:      "#0D0D1A", selection_bg: "#1A1A33", number_color: "#00FFFF",
            normal: "#00FF41", insert: "#00FFFF", visual: "#FF00FF", command: "#FF2A6D",
        },
        theme! {
            id: "amber-terminal", name: "Amber Terminal", dark: true,
            desc: "Classic phosphor amber monochrome, like an old IBM terminal. Pure nostalgia.",
            cursor_bg:    "#FF8C00", header_bg: "#050200", header_fg: "#3A2000",
            cell_bg:      "#0A0500", selection_bg: "#2A1500", number_color: "#FFA500",
            normal: "#FF8C00", insert: "#FFA500", visual: "#FFD700", command: "#FF4500",
        },
        theme! {
            id: "ice", name: "Ice", dark: true,
            desc: "Cool blues and pale cyans. Crisp and minimal like a frozen lake.",
            cursor_bg:    "#89DCEB", header_bg: "#151525", header_fg: "#3A4060",
            cell_bg:      "#1C1C2C", selection_bg: "#2A3050", number_color: "#89DCEB",
            normal: "#89DCEB", insert: "#89B4FA", visual: "#B5C1FF", command: "#F28FAD",
        },
        theme! {
            id: "github-dark", name: "GitHub Dark", dark: true,
            desc: "GitHub's official dark theme. Familiar, professional, well-tested.",
            cursor_bg:    "#58A6FF", header_bg: "#161B22", header_fg: "#484F58",
            cell_bg:      "#0D1117", selection_bg: "#1F3464", number_color: "#39C5CF",
            normal: "#3FB950", insert: "#58A6FF", visual: "#BC8CFF", command: "#F78166",
        },
    ]
}
