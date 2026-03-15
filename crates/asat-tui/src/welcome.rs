use crate::{darken, is_dark_color, parse_hex_color, RenderState};
use ratatui::{
    layout::{Alignment, Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, BorderType, Borders, Clear, List, ListItem, Paragraph},
    Frame,
};
use unicode_width::UnicodeWidthStr;

// ASAT logo in Doom-font style (box-drawing chars)
const LOGO: &[&str] = &[
    " █████╗  ███████╗ █████╗ ████████╗",
    "██╔══██╗ ██╔════╝██╔══██╗╚══██╔══╝",
    "███████║ ███████╗███████║   ██║   ",
    "██╔══██║ ╚════██║██╔══██║   ██║   ",
    "██║  ██║ ███████║██║  ██║   ██║   ",
    "╚═╝  ╚═╝ ╚══════╝╚═╝  ╚═╝   ╚═╝  ",
];

struct MenuItem {
    key: &'static str,
    icon: &'static str,
    name: &'static str,
    desc: &'static str,
}

const MENU: &[MenuItem] = &[
    MenuItem {
        key: "n",
        icon: "◆",
        name: "New File",
        desc: "Start with a blank workbook",
    },
    MenuItem {
        key: "f",
        icon: "◆",
        name: "Find File",
        desc: "Fuzzy-search files in current directory",
    },
    MenuItem {
        key: "r",
        icon: "◆",
        name: "Recent Files",
        desc: "Browse recently opened files",
    },
    MenuItem {
        key: "t",
        icon: "◆",
        name: "Themes",
        desc: "Browse and apply colour themes",
    },
    MenuItem {
        key: "c",
        icon: "◆",
        name: "Edit Config",
        desc: "Open config.toml in $EDITOR",
    },
    MenuItem {
        key: "q",
        icon: "◆",
        name: "Quit",
        desc: "Exit ASAT",
    },
];

pub fn render_welcome(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let theme = &state.config.theme;
    let amber = parse_hex_color(&theme.cursor_bg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let bg = parse_hex_color(&theme.cell_bg);
    let dim = darken(header_fg, 0.55);

    // Fill background
    frame.render_widget(Paragraph::new("").style(Style::default().bg(bg)), area);

    // ── Vertical layout ───────────────────────────────────────────────────
    // We centre a content block vertically
    let logo_h = LOGO.len() as u16;
    let content_h = logo_h + 2 /* subtitle */ + 1 /* spacer */ + 1 /* divider */
                  + MENU.len() as u16 + 2 /* spacer + version */;
    let top_pad = area.height.saturating_sub(content_h) / 2;

    let rows = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(top_pad),
            Constraint::Length(logo_h),
            Constraint::Length(1), // subtitle line 1
            Constraint::Length(1), // subtitle line 2
            Constraint::Length(1), // spacer
            Constraint::Length(1), // divider
            Constraint::Length(MENU.len() as u16),
            Constraint::Length(1), // spacer
            Constraint::Length(1), // version footer
            Constraint::Min(0),
        ])
        .split(area);

    // Use display width (columns), not byte length — box-drawing chars are multibyte but 1-wide
    let logo_width = LOGO
        .iter()
        .map(|l| UnicodeWidthStr::width(*l) as u16)
        .max()
        .unwrap_or(30);
    let center_x = area.x + area.width.saturating_sub(logo_width) / 2;

    // ── Logo ──────────────────────────────────────────────────────────────
    let logo_area = rows[1];
    for (i, line) in LOGO.iter().enumerate() {
        let y = logo_area.y + i as u16;
        if y >= logo_area.y + logo_area.height {
            break;
        }
        let rect = Rect {
            x: center_x,
            y,
            width: logo_width,
            height: 1,
        };
        frame.render_widget(
            Paragraph::new(Line::from(Span::styled(
                *line,
                Style::default().fg(amber).add_modifier(Modifier::BOLD),
            ))),
            rect,
        );
    }

    // ── Subtitle ──────────────────────────────────────────────────────────
    frame.render_widget(
        Paragraph::new(Line::from(vec![Span::styled(
            "Advanced Spreadsheet Alteration Tool",
            Style::default().fg(header_fg),
        )]))
        .alignment(Alignment::Center),
        rows[2],
    );
    frame.render_widget(
        Paragraph::new(Line::from(vec![Span::styled(
            "vim-powered  ·  terminal-native  ·  written in rust",
            Style::default().fg(dim),
        )]))
        .alignment(Alignment::Center),
        rows[3],
    );

    // ── Divider ───────────────────────────────────────────────────────────
    let div_w = 52u16.min(area.width);
    let div_x = area.x + area.width.saturating_sub(div_w) / 2;
    frame.render_widget(
        Paragraph::new(Line::from(Span::styled(
            "─".repeat(div_w as usize),
            Style::default().fg(dim),
        ))),
        Rect {
            x: div_x,
            y: rows[5].y,
            width: div_w,
            height: 1,
        },
    );

    // ── Menu items ────────────────────────────────────────────────────────
    let menu_w = 52u16.min(area.width);
    let menu_x = area.x + area.width.saturating_sub(menu_w) / 2;
    let menu_area = Rect {
        x: menu_x,
        y: rows[6].y,
        width: menu_w,
        height: rows[6].height,
    };

    for (i, item) in MENU.iter().enumerate() {
        let y = menu_area.y + i as u16;
        if y >= menu_area.y + menu_area.height {
            break;
        }
        let row = Rect {
            x: menu_x,
            y,
            width: menu_w,
            height: 1,
        };

        let is_quit = item.key == "q";
        let key_color = if is_quit {
            darken(header_fg, 0.75)
        } else {
            amber
        };
        let name_color = if is_quit {
            darken(header_fg, 0.65)
        } else {
            Color::White
        };

        frame.render_widget(
            Paragraph::new(Line::from(vec![
                Span::styled(format!(" {} ", item.icon), Style::default().fg(key_color)),
                Span::styled(
                    format!("{}", item.key),
                    Style::default().fg(key_color).add_modifier(Modifier::BOLD),
                ),
                Span::styled("  ", Style::default()),
                Span::styled(
                    format!("{:<18}", item.name),
                    Style::default().fg(name_color).add_modifier(Modifier::BOLD),
                ),
                Span::styled(item.desc, Style::default().fg(header_fg)),
            ])),
            row,
        );
    }

    // ── Version footer ────────────────────────────────────────────────────
    frame.render_widget(
        Paragraph::new(Line::from(vec![
            Span::styled("v0.1.0", Style::default().fg(dim)),
            Span::styled(
                "  ·  press ? for help",
                Style::default().fg(darken(header_fg, 0.4)),
            ),
        ]))
        .alignment(Alignment::Center),
        rows[8],
    );
}

// ── Fuzzy File Finder ─────────────────────────────────────────────────────────

pub fn render_file_finder(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let theme = &state.config.theme;
    let amber = parse_hex_color(&theme.cursor_bg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let sel_bg = parse_hex_color(&theme.selection_bg);
    let bg = parse_hex_color(&theme.cell_bg);

    // Draw background
    frame.render_widget(Paragraph::new("").style(Style::default().bg(bg)), area);

    // Panel dimensions: centred, reasonable size
    let panel_w = (area.width * 3 / 4)
        .max(50)
        .min(area.width.saturating_sub(4));
    let panel_h = (area.height * 3 / 4)
        .max(12)
        .min(area.height.saturating_sub(4));
    let panel_x = area.x + (area.width - panel_w) / 2;
    let panel_y = area.y + (area.height - panel_h) / 2;
    let panel = Rect {
        x: panel_x,
        y: panel_y,
        width: panel_w,
        height: panel_h,
    };

    let query = &state.input.finder_query;
    let files = state.input.filtered_finder_files();
    let selected = state.input.finder_selected;
    let total = state.input.finder_files.len();

    // Inner layout: search bar (3 rows) + results (rest)
    let inner = Layout::default()
        .direction(Direction::Vertical)
        .constraints([Constraint::Length(3), Constraint::Min(1)])
        .split(panel);

    // ── Search box ────────────────────────────────────────────────────────
    let search_block = Block::default()
        .title(format!(" Find File  ({} files) ", total))
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(amber))
        .title_style(Style::default().fg(amber).add_modifier(Modifier::BOLD))
        .style(Style::default().bg(header_bg));

    let search_line = Line::from(vec![
        Span::styled(
            " ❯ ",
            Style::default().fg(amber).add_modifier(Modifier::BOLD),
        ),
        Span::styled(query.as_str(), Style::default().fg(Color::White)),
        Span::styled("█", Style::default().fg(Color::White)),
    ]);

    frame.render_widget(Clear, inner[0]);
    frame.render_widget(Paragraph::new(search_line).block(search_block), inner[0]);

    // ── Results list ──────────────────────────────────────────────────────
    let results_block = Block::default()
        .borders(Borders::LEFT | Borders::RIGHT | Borders::BOTTOM)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(darken(header_fg, 0.55)))
        .style(Style::default().bg(bg));

    let visible_h = inner[1].height.saturating_sub(2) as usize;
    // Scroll the view so `selected` is always visible
    let scroll_top = if selected >= visible_h {
        selected - visible_h + 1
    } else {
        0
    };

    let items: Vec<ListItem> = files
        .iter()
        .enumerate()
        .skip(scroll_top)
        .take(visible_h)
        .map(|(i, path)| {
            let is_sel = i == selected;
            if is_sel {
                ListItem::new(Line::from(vec![
                    Span::styled(" ▶ ", Style::default().fg(amber)),
                    Span::styled(
                        path.as_str(),
                        Style::default()
                            .fg(Color::White)
                            .add_modifier(Modifier::BOLD),
                    ),
                ]))
                .style(Style::default().bg(sel_bg))
            } else {
                ListItem::new(Line::from(vec![
                    Span::styled("   ", Style::default()),
                    Span::styled(path.as_str(), Style::default().fg(header_fg)),
                ]))
            }
        })
        .collect();

    if files.is_empty() {
        let empty = Paragraph::new(Line::from(Span::styled(
            if query.is_empty() {
                " No files found in current directory"
            } else {
                " No files match your query"
            },
            Style::default().fg(darken(header_fg, 0.55)),
        )))
        .block(results_block);
        frame.render_widget(Clear, inner[1]);
        frame.render_widget(empty, inner[1]);
    } else {
        let hint = format!(" {}/{} ", files.len().min(selected + 1), files.len());
        let hint_block = results_block.title_bottom(Span::styled(
            hint,
            Style::default().fg(darken(header_fg, 0.55)),
        ));
        frame.render_widget(Clear, inner[1]);
        frame.render_widget(List::new(items).block(hint_block), inner[1]);
    }
}

// ── Recent Files ──────────────────────────────────────────────────────────────

pub fn render_recent_files(frame: &mut Frame, area: Rect, state: &RenderState<'_>) {
    let theme = &state.config.theme;
    let amber = parse_hex_color(&theme.cursor_bg);
    let header_bg = parse_hex_color(&theme.header_bg);
    let header_fg = parse_hex_color(&theme.header_fg);
    let sel_bg = parse_hex_color(&theme.selection_bg);
    let bg = parse_hex_color(&theme.cell_bg);
    let sel_fg = if is_dark_color(sel_bg) {
        Color::White
    } else {
        Color::Black
    };

    frame.render_widget(Paragraph::new("").style(Style::default().bg(bg)), area);

    let panel_w = (area.width * 2 / 3)
        .max(50)
        .min(area.width.saturating_sub(4));
    let files = &state.input.recent_files;
    let panel_h = ((files.len() as u16 + 4).max(6)).min(area.height.saturating_sub(4));
    let panel_x = area.x + (area.width - panel_w) / 2;
    let panel_y = area.y + (area.height - panel_h) / 2;
    let panel = Rect {
        x: panel_x,
        y: panel_y,
        width: panel_w,
        height: panel_h,
    };

    let selected = state.input.recent_selected;

    let block = Block::default()
        .title(" Recent Files ")
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(amber))
        .title_style(Style::default().fg(amber).add_modifier(Modifier::BOLD))
        .style(Style::default().bg(header_bg));

    if files.is_empty() {
        let para = Paragraph::new(vec![
            Line::from(""),
            Line::from(Span::styled(
                "  No recent files yet.",
                Style::default().fg(darken(header_fg, 0.6)),
            )),
            Line::from(Span::styled(
                "  Open a file with :e <path> to get started.",
                Style::default().fg(darken(header_fg, 0.45)),
            )),
        ])
        .block(block);
        frame.render_widget(Clear, panel);
        frame.render_widget(para, panel);
        return;
    }

    let items: Vec<ListItem> = files
        .iter()
        .enumerate()
        .map(|(i, path)| {
            let is_sel = i == selected;
            // Show just the filename bold + dim parent path
            let p = std::path::Path::new(path);
            let name = p.file_name().and_then(|n| n.to_str()).unwrap_or(path);
            let parent = p.parent().and_then(|p| p.to_str()).unwrap_or("");

            if is_sel {
                ListItem::new(Line::from(vec![
                    Span::styled(" ▶ ", Style::default().fg(amber)),
                    Span::styled(
                        name,
                        Style::default().fg(sel_fg).add_modifier(Modifier::BOLD),
                    ),
                    Span::styled(
                        format!("  {}", parent),
                        Style::default().fg(darken(sel_fg, 0.6)),
                    ),
                ]))
                .style(Style::default().bg(sel_bg))
            } else {
                ListItem::new(Line::from(vec![
                    Span::styled("   ", Style::default()),
                    Span::styled(name, Style::default().fg(header_fg)),
                    Span::styled(
                        format!("  {}", parent),
                        Style::default().fg(darken(header_fg, 0.55)),
                    ),
                ]))
            }
        })
        .collect();

    frame.render_widget(Clear, panel);
    frame.render_widget(List::new(items).block(block), panel);
}
