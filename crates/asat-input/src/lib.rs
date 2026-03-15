use asat_commands::RegisterMap;
use asat_core::{CellStyle, CellValue, Workbook};
use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};

// ── Ex-command completion list ────────────────────────────────────────────────

/// All supported ex-commands with short descriptions, used for completion.
pub const EX_COMMANDS: &[(&str, &str)] = &[
    ("q", "Quit (warns if unsaved)"),
    ("q!", "Force quit without saving"),
    ("w", "Save to current file"),
    ("w <file>", "Save to a new file"),
    ("wq", "Save and quit"),
    ("x", "Save and quit"),
    ("e <file>", "Open file"),
    ("tabnew", "New sheet"),
    ("tabedit", "New sheet (alias)"),
    ("tabclose", "Close current sheet"),
    ("ic", "Insert column left"),
    ("icr", "Insert column right"),
    ("dc", "Delete current column"),
    ("ir", "Insert row"),
    ("dr", "Delete current row"),
    ("cw <N>", "Set column width to N"),
    ("rh <N>", "Set row height to N"),
    ("bold", "Toggle bold on cell / selection"),
    ("italic", "Toggle italic on cell / selection"),
    ("underline", "Toggle underline on cell / selection"),
    ("strike", "Toggle strikethrough on cell / selection"),
    ("fg <color>", "Set foreground colour (hex #rrggbb or name)"),
    ("bg <color>", "Set background colour (hex #rrggbb or name)"),
    ("hl <color>", "Highlight: set bg + auto-contrast fg"),
    ("hl", "Clear highlight (remove bg/fg colours)"),
    ("align <l/c/r>", "Set alignment: left, center, or right"),
    ("fmt <spec>", "Number format: %, $, 0.00, int, date, none"),
    ("copystyle", "Copy current cell style to clipboard"),
    ("pastestyle", "Paste style clipboard to cell / selection"),
    ("cs", "Clear all styles from cell / selection"),
    ("theme", "Open theme picker (or :theme <name> to apply)"),
    ("set", "Set an option"),
    (
        "sort",
        "Sort rows by cursor column (:sort asc / :sort desc)",
    ),
    (
        "s /pat/repl/",
        "Find & replace in text cells (:s/pat/repl/g)",
    ),
    ("plugin", "Plugin engine: :plugin reload | :plugin list"),
    ("goto <cell>", "Jump to cell address (e.g. :goto B15)"),
    (
        "name <n> <r>",
        "Define named range (e.g. :name sales A1:C10)",
    ),
    (
        "filter <c> <op> <v>",
        "Filter rows by column value (e.g. :filter A >100)",
    ),
    ("filter off", "Clear row filter"),
    ("transpose", "Transpose current visual selection"),
    ("dedup", "Remove duplicate rows (by cursor column)"),
    ("note <text>", "Set a note/comment on the current cell"),
    (
        "colfmt <op> <v> <color>",
        "Conditional format (e.g. :colfmt >100 red)",
    ),
    ("filldown", "Fill selection down from top row"),
    ("fillright", "Fill selection right from leftmost column"),
];

/// All built-in formula function names, for Tab-completion in Insert mode.
pub const FN_NAMES: &[&str] = &[
    "SUM",
    "AVERAGE",
    "AVG",
    "COUNT",
    "COUNTA",
    "MIN",
    "MAX",
    "IF",
    "AND",
    "OR",
    "NOT",
    "ABS",
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "FLOOR",
    "CEILING",
    "MOD",
    "POWER",
    "SQRT",
    "LN",
    "LOG",
    "LOG10",
    "EXP",
    "INT",
    "TRUNC",
    "SIGN",
    "LEN",
    "LEFT",
    "RIGHT",
    "MID",
    "TRIM",
    "UPPER",
    "LOWER",
    "PROPER",
    "CONCATENATE",
    "CONCAT",
    "TEXT",
    "VALUE",
    "FIND",
    "SEARCH",
    "SUBSTITUTE",
    "REPLACE",
    "REPT",
    "ISNUMBER",
    "ISTEXT",
    "ISBLANK",
    "ISERROR",
    "IFERROR",
    "ISLOGICAL",
    "TRUE",
    "FALSE",
    "PI",
    "SUMIF",
    "COUNTIF",
    "SUMPRODUCT",
    "LARGE",
    "SMALL",
    "MEDIAN",
    "STDEV",
    "VAR",
    "AVERAGEIF",
    "MAXIFS",
    "MINIFS",
    "RANK",
    "PERCENTILE",
    "QUARTILE",
    "XLOOKUP",
    "CHOOSE",
    "PV",
    "FV",
    "PMT",
    "NPER",
    "RATE",
    "NPV",
    "IRR",
    "MIRR",
    "IPMT",
    "PPMT",
    "SLN",
    "DDB",
    "EFFECT",
    "NOMINAL",
    "CUMIPMT",
    "CUMPRINC",
];

/// Return completions whose command word starts with `prefix` (case-insensitive).
pub fn get_command_completions(prefix: &str) -> Vec<(&'static str, &'static str)> {
    let p = prefix.to_ascii_lowercase();
    EX_COMMANDS
        .iter()
        .filter(|(cmd, _)| {
            // Compare only up to the first space or '<' (the argument placeholder)
            let word = cmd
                .split(|c: char| c == ' ' || c == '<')
                .next()
                .unwrap_or(cmd);
            word.to_ascii_lowercase().starts_with(&p)
        })
        .copied()
        .collect()
}

// ── Mode ──────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq)]
pub enum Mode {
    Normal,
    Insert {
        replace: bool,
    },
    Visual {
        block: bool,
    }, // v = char visual, Ctrl+V = column/block
    VisualLine, // V = whole-row selection
    Command,
    Search {
        forward: bool,
    },
    Recording {
        register: char,
    },
    Welcome,      // home / start screen
    FileFind,     // fuzzy file finder
    RecentFiles,  // recent files list
    ThemeManager, // theme picker
    /// Navigate the grid with hjkl to pick a cell/range reference while writing a formula.
    /// `anchor` is set when the user pressed `:` to start a range.
    FormulaSelect {
        anchor: Option<(u32, u32)>,
    },
}

impl Mode {
    pub fn name(&self) -> &str {
        match self {
            Mode::Normal => "NORMAL",
            Mode::Insert { replace: false } => "INSERT",
            Mode::Insert { replace: true } => "REPLACE",
            Mode::Visual { block: false } => "VISUAL",
            Mode::Visual { block: true } => "V-COL",
            Mode::VisualLine => "V-ROW",
            Mode::Command => "COMMAND",
            Mode::Search { forward: true } => "SEARCH",
            Mode::Search { forward: false } => "SEARCH↑",
            Mode::Recording { register } => {
                Box::leak(format!("REC({})", register).into_boxed_str())
            }
            Mode::Welcome => "WELCOME",
            Mode::FileFind => "FIND FILE",
            Mode::RecentFiles => "RECENT",
            Mode::ThemeManager => "THEMES",
            Mode::FormulaSelect { anchor: None } => "F-REF",
            Mode::FormulaSelect { anchor: Some(_) } => "F-RANGE",
        }
    }
}

// ── Cursor & Viewport ─────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Cursor {
    pub row: u32,
    pub col: u32,
}

impl Cursor {
    pub fn new() -> Self {
        Cursor { row: 0, col: 0 }
    }
}

impl Default for Cursor {
    fn default() -> Self {
        Cursor::new()
    }
}

#[derive(Debug, Clone, Copy, Default)]
pub struct Viewport {
    pub top_row: u32,
    pub left_col: u32,
}

// ── Visual anchor ─────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy)]
pub struct VisualAnchor {
    pub row: u32,
    pub col: u32,
}

// ── App Actions ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum AppAction {
    // ── Navigation ──
    MoveCursor {
        row_delta: i32,
        col_delta: i32,
    },
    MoveCursorTo {
        row: u32,
        col: u32,
    },
    MoveToFirstRow,
    MoveToLastRow,
    MoveToFirstCol,
    MoveToLastCol,
    PageDown,
    PageUp,
    /// Jump to next/prev non-empty cell horizontally (w / b)
    MoveToNextNonEmptyH {
        forward: bool,
    },
    /// Jump to next/prev non-empty cell vertically (W / B)
    MoveToNextNonEmptyV {
        forward: bool,
    },
    /// Jump to next/prev "paragraph" (empty row boundary) (} / {)
    MoveToNextParagraph {
        forward: bool,
    },
    /// H / M / L — jump to top/middle/bottom of visible rows
    JumpHighRow,
    JumpMidRow,
    JumpLowRow,

    // ── Scroll without moving cursor ──
    ScrollCenter, // zz
    ScrollTop,    // zt
    ScrollBottom, // zb

    // ── Mode transitions ──
    EnterInsert {
        replace: bool,
    },
    EnterVisual {
        block: bool,
    },
    EnterVisualLine,
    EnterCommand,
    EnterSearch {
        forward: bool,
    },
    ExitMode,

    // ── Cell editing ──
    SetCell {
        sheet: usize,
        row: u32,
        col: u32,
        value: CellValue,
    },
    DeleteCellContent,
    DeleteCellRange {
        row_start: u32,
        col_start: u32,
        row_end: u32,
        col_end: u32,
    },
    ChangeCell, // c — clear cell, enter insert
    ToggleCase, // ~ — toggle uppercase/lowercase of text cell

    // ── Row / column operations ──
    InsertRowAbove,
    InsertRowBelow,
    OpenRowBelow, // o — insert row below + move there + insert mode
    OpenRowAbove, // O — insert row above + stay + insert mode
    DeleteCurrentRow,
    InsertColLeft,
    InsertColRight,
    DeleteCurrentCol,
    IncreaseColWidth {
        col: u32,
    }, // >>
    DecreaseColWidth {
        col: u32,
    }, // <<
    AutoFitCol {
        col: u32,
    }, // =
    IncreaseRowHeight {
        row: u32,
    }, // +
    DecreaseRowHeight {
        row: u32,
    }, // -
    AutoFitRow {
        row: u32,
    }, // _

    // ── Yank / paste ──
    YankRow,
    YankCell, // yc — yank single cell to register + system clipboard
    YankCellRange {
        row_start: u32,
        col_start: u32,
        row_end: u32,
        col_end: u32,
        is_line: bool,
    },
    PasteAfter,
    PasteBefore,

    // ── Cell arithmetic ──
    IncrementCell, // Ctrl+A — add 1 to number
    DecrementCell, // Ctrl+X — subtract 1 from number

    // ── Spreadsheet-specific ──
    JumpToCol {
        col: u32,
    }, // f{A-Z} — jump to column by letter
    JoinCellBelow,        // J — concat cell below into current, clear below
    VisualClearAndInsert, // c/s in visual — clear selection and enter insert
    InsertSumBelow {
        row_start: u32,
        col_start: u32,
        row_end: u32,
        col_end: u32,
    }, // S in visual

    // ── Undo / redo ──
    Undo,
    Redo,

    // ── Marks ──
    SetMark {
        ch: char,
    },
    JumpToMark {
        ch: char,
    },
    JumpToPrevPos, // ''

    // ── Search ──
    ExecuteSearch,
    FindNext,
    FindPrev,
    SearchCurrentCell, // * — search for content under cursor

    // ── Sheet navigation ──
    NextSheet,
    PrevSheet,

    // ── Macro recording/playback ──
    StartRecording {
        register: char,
    },
    StopRecording,
    PlayMacro {
        register: char,
    },

    // ── Welcome screen ──
    WelcomeNewFile,
    WelcomeEnterFileFind,
    WelcomeEnterRecent,
    WelcomeOpenConfig,
    WelcomeOpenThemes,

    // ── File finder ──
    FinderOpen, // open selected file
    FinderMoveUp,
    FinderMoveDown,
    FinderCancel,

    // ── Recent files ──
    RecentOpen,
    RecentMoveUp,
    RecentMoveDown,
    RecentCancel,

    // ── Theme manager ──
    ThemeApply(usize), // apply preset at this index and save
    ThemeManagerCancel,

    // ── Style copy/paste ──
    YankStyle,  // yS — copy current cell's style to clipboard
    PasteStyle, // pS — paste style clipboard to current cell / visual selection

    // ── Formula reference selection ──
    EnterFormulaSelect,      // enter cell-picking mode from Insert
    FormulaSelectConfirm,    // insert the selected cell/range ref into formula
    FormulaSelectStartRange, // mark current cell as range anchor (press ':')
    FormulaSelectCancel,     // return to Insert without inserting

    // ── Clipboard ──
    PasteFromClipboard, // Ctrl+V in Insert — paste system clipboard text into edit buffer

    // ── Fill ──
    /// Fill the anchor row/col values down / right across the selection
    FillDown {
        anchor_row: u32,
        col_start: u32,
        col_end: u32,
        row_end: u32,
    },
    FillRight {
        anchor_col: u32,
        row_start: u32,
        row_end: u32,
        col_end: u32,
    },

    // ── Goto ──
    GotoCell(String), // jump to a cell address like "B15"

    // ── Cell notes ──
    SetNote {
        row: u32,
        col: u32,
        text: String,
    },

    // ── Commands ──
    ExecuteCommand2(String),

    // ── App ──
    Quit,
    QuitForce,
    Save,
    SaveAs(String),
    OpenFile(String),
    NoOp,
}

// ── Input State ───────────────────────────────────────────────────────────────

pub struct InputState {
    pub mode: Mode,
    pub cursor: Cursor,
    pub viewport: Viewport,
    /// All cells matching the last search, in row-major order
    pub search_matches: Vec<(u32, u32)>,
    /// Index into search_matches of the currently-highlighted match
    pub search_match_idx: usize,
    pub visual_anchor: Option<VisualAnchor>,
    pub registers: RegisterMap,
    pub marks: std::collections::HashMap<char, (usize, Cursor)>,
    pub prev_position: Option<(usize, Cursor)>, // for ''

    pub edit_buffer: String,
    pub edit_cursor_pos: usize,
    pub command_buffer: String,
    pub search_buffer: String,
    pub last_search: Option<(String, bool)>,

    /// Key sequences captured while recording a macro
    pub recording_buffer: Vec<KeyEvent>,
    /// Named macro registers (a-z)
    pub macro_registers: std::collections::HashMap<char, Vec<KeyEvent>>,
    /// Last-used macro register (for `@@`)
    pub last_macro_register: Option<char>,

    /// Number of rows/cols to keep between cursor and viewport edge (scrolloff)
    pub scroll_padding: u32,

    /// Index into the currently filtered completion list (None = no active completion)
    pub completion_idx: Option<usize>,
    /// The command-buffer prefix that started the current completion session
    pub completion_prefix: String,

    // ── File finder state ──
    /// All scanned file paths (populated by main.rs when entering FileFind)
    pub finder_files: Vec<String>,
    /// Live-typed query for filtering finder_files
    pub finder_query: String,
    /// Currently highlighted index in the *filtered* results
    pub finder_selected: usize,

    // ── Recent files state ──
    pub recent_files: Vec<String>,
    pub recent_selected: usize,

    // ── Theme manager state ──
    pub theme_selected: usize,

    // ── Style clipboard ──
    pub style_clipboard: Option<CellStyle>,

    // ── Formula reference selection ──
    /// Cell being edited when FormulaSelect was entered (restored on cancel)
    pub formula_origin: Option<(u32, u32)>,

    // ── Sub-command completion (e.g. :theme <name>) ──
    /// Possible completions for the current command's argument (populated by main.rs)
    pub subcmd_completions: Vec<String>,
    /// Index into subcmd_completions while cycling with Tab
    pub subcmd_completion_idx: Option<usize>,

    // ── Formula function tab-completion ──
    pub fn_completion_prefix: String,
    pub fn_completion_candidates: Vec<String>,
    pub fn_completion_idx: Option<usize>,

    key_buffer: Vec<KeyEvent>,
    count_buffer: String,
}

impl InputState {
    pub fn new() -> Self {
        InputState {
            mode: Mode::Normal,
            cursor: Cursor::new(),
            viewport: Viewport::default(),
            visual_anchor: None,
            registers: RegisterMap::default(),
            marks: std::collections::HashMap::new(),
            prev_position: None,
            edit_buffer: String::new(),
            edit_cursor_pos: 0,
            command_buffer: String::new(),
            search_buffer: String::new(),
            last_search: None,
            search_matches: Vec::new(),
            search_match_idx: 0,
            recording_buffer: Vec::new(),
            macro_registers: std::collections::HashMap::new(),
            last_macro_register: None,
            scroll_padding: 3,
            completion_idx: None,
            completion_prefix: String::new(),
            finder_files: Vec::new(),
            finder_query: String::new(),
            finder_selected: 0,
            recent_files: Vec::new(),
            recent_selected: 0,
            theme_selected: 0,
            style_clipboard: None,
            formula_origin: None,
            subcmd_completions: Vec::new(),
            subcmd_completion_idx: None,
            fn_completion_prefix: String::new(),
            fn_completion_candidates: Vec::new(),
            fn_completion_idx: None,
            key_buffer: Vec::new(),
            count_buffer: String::new(),
        }
    }

    /// Returns the search highlight type for a cell: Some(true) = current match, Some(false) = other match
    pub fn search_highlight(&self, row: u32, col: u32) -> Option<bool> {
        let pos = self
            .search_matches
            .iter()
            .position(|&(r, c)| r == row && c == col)?;
        Some(pos == self.search_match_idx)
    }

    pub fn count(&self) -> u32 {
        self.count_buffer.parse().unwrap_or(1).max(1)
    }

    fn take_count(&mut self) -> u32 {
        let n = self.count();
        self.count_buffer.clear();
        n
    }

    /// Returns the current partial key sequence as a display string
    pub fn key_prefix(&self) -> String {
        self.key_buffer
            .iter()
            .map(|k| key_event_to_str(k))
            .collect()
    }

    /// Current visual selection bounds (row_start, col_start, row_end, col_end).
    /// V-ROW returns col_end=u32::MAX; V-COL returns row_end=u32::MAX.
    /// Callers must clamp to sheet bounds before iterating.
    pub fn visual_selection_bounds(&self) -> (u32, u32, u32, u32) {
        let (ar, ac) = self
            .visual_anchor
            .map(|a| (a.row, a.col))
            .unwrap_or((self.cursor.row, self.cursor.col));
        let (cr, cc) = (self.cursor.row, self.cursor.col);
        match &self.mode {
            Mode::VisualLine => (ar.min(cr), 0, ar.max(cr), u32::MAX),
            Mode::Visual { block: true } => (0, ac.min(cc), u32::MAX, ac.max(cc)),
            _ => (ar.min(cr), ac.min(cc), ar.max(cr), ac.max(cc)),
        }
    }

    pub fn scroll_to_cursor(&mut self, visible_rows: u32, visible_cols: u32) {
        // Clamp padding so it can never exceed 1/3 of the visible area
        let rpad = self.scroll_padding.min(visible_rows / 3);
        let cpad = self.scroll_padding.min(visible_cols / 3);

        let row_min = self.cursor.row.saturating_sub(rpad);
        let row_max = self.cursor.row.saturating_add(rpad);
        if row_min < self.viewport.top_row {
            self.viewport.top_row = row_min;
        } else if row_max >= self.viewport.top_row + visible_rows {
            self.viewport.top_row = row_max + 1 - visible_rows;
        }

        let col_min = self.cursor.col.saturating_sub(cpad);
        let col_max = self.cursor.col.saturating_add(cpad);
        if col_min < self.viewport.left_col {
            self.viewport.left_col = col_min;
        } else if col_max >= self.viewport.left_col + visible_cols {
            self.viewport.left_col = col_max + 1 - visible_cols;
        }
    }

    pub fn save_position(&mut self, sheet: usize) {
        self.prev_position = Some((sheet, self.cursor));
    }

    /// Cycle through function name completions when editing a formula.
    /// Finds the current identifier before the cursor and replaces it with
    /// the next/prev match from the function name list.
    fn cycle_fn_completion(&mut self, forward: bool) {
        // Extract the identifier at the cursor position (letters/digits/underscore)
        let buf = &self.edit_buffer[1..]; // skip leading '='
        let cursor_in_buf = self.edit_cursor_pos.saturating_sub(1);
        let cursor_in_buf = cursor_in_buf.min(buf.len());

        // Walk backward to find start of identifier
        let id_start = buf[..cursor_in_buf]
            .rfind(|c: char| !c.is_alphanumeric() && c != '_')
            .map(|p| p + 1)
            .unwrap_or(0);
        let prefix = buf[id_start..cursor_in_buf].to_uppercase();

        if prefix.is_empty() {
            return;
        }

        // Rebuild candidate list if prefix changed
        if prefix != self.fn_completion_prefix {
            self.fn_completion_prefix = prefix.clone();
            self.fn_completion_candidates = FN_NAMES
                .iter()
                .filter(|n| n.starts_with(&prefix))
                .map(|n| n.to_string())
                .collect();
            self.fn_completion_idx = None;
        }

        if self.fn_completion_candidates.is_empty() {
            return;
        }

        let len = self.fn_completion_candidates.len();
        let next = match self.fn_completion_idx {
            None => {
                if forward {
                    0
                } else {
                    len - 1
                }
            }
            Some(i) => {
                if forward {
                    (i + 1) % len
                } else {
                    if i == 0 {
                        len - 1
                    } else {
                        i - 1
                    }
                }
            }
        };
        self.fn_completion_idx = Some(next);

        let chosen = &self.fn_completion_candidates[next];
        // Replace the prefix in the buffer with chosen + '('
        let full_start = 1 + id_start; // +1 for the '='
        let full_end = self.edit_cursor_pos;
        self.edit_buffer.drain(full_start..full_end);
        let replacement = format!("{}(", chosen);
        self.edit_buffer.insert_str(full_start, &replacement);
        self.edit_cursor_pos = full_start + replacement.len();
    }

    pub fn handle_key(&mut self, key: KeyEvent, workbook: &Workbook) -> Vec<AppAction> {
        // Capture key while recording (before dispatch, so the stop-key handler
        // can pop it back off if it decides not to record it).
        if matches!(self.mode, Mode::Recording { .. }) {
            self.recording_buffer.push(key);
        }

        match self.mode.clone() {
            Mode::Normal => self.handle_normal(key, workbook),
            Mode::Insert { replace } => self.handle_insert(key, replace),
            Mode::Visual { .. } => self.handle_visual(key, false),
            Mode::VisualLine => self.handle_visual(key, true),
            Mode::Command => self.handle_command(key),
            Mode::Search { forward } => self.handle_search(key, forward),
            Mode::Recording { .. } => self.handle_normal(key, workbook),
            Mode::Welcome => self.handle_welcome(key),
            Mode::FileFind => self.handle_file_find(key),
            Mode::RecentFiles => self.handle_recent_files(key),
            Mode::ThemeManager => self.handle_theme_manager(key),
            Mode::FormulaSelect { .. } => self.handle_formula_select(key),
        }
    }

    // ── Normal mode ───────────────────────────────────────────────────────────

    fn handle_normal(&mut self, key: KeyEvent, workbook: &Workbook) -> Vec<AppAction> {
        let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);

        // ── Stop recording when 'q' pressed in Recording mode ─────────────────
        if matches!(self.mode, Mode::Recording { .. }) && key.code == KeyCode::Char('q') && !ctrl {
            self.recording_buffer.pop(); // remove the 'q' itself from the buffer
            return vec![AppAction::StopRecording];
        }

        // ── Pending multi-key sequence ────────────────────────────────────────
        if !self.key_buffer.is_empty() {
            let first = self.key_buffer[0].code;
            let first_ctrl = self.key_buffer[0].modifiers.contains(KeyModifiers::CONTROL);
            self.key_buffer.clear();
            let n = self.take_count();

            return match (first, key.code) {
                // gg / gt / gT
                (KeyCode::Char('g'), KeyCode::Char('g')) => {
                    vec![AppAction::MoveToFirstRow]
                }
                (KeyCode::Char('g'), KeyCode::Char('t')) => vec![AppAction::NextSheet],
                (KeyCode::Char('g'), KeyCode::Char('T')) => vec![AppAction::PrevSheet],

                // g{A-Z} — goto column (single letter address shortcut)
                (KeyCode::Char('g'), KeyCode::Char(c)) if c.is_ascii_uppercase() => {
                    // Treat as goto column if no row digit follows — just jump to col
                    let col = (c as u32).saturating_sub('A' as u32);
                    vec![AppAction::MoveCursorTo { row: 0, col }]
                }

                // dd / yy / yc / yr / dC
                (KeyCode::Char('d'), KeyCode::Char('d')) => {
                    (0..n).map(|_| AppAction::DeleteCurrentRow).collect()
                }
                (KeyCode::Char('d'), KeyCode::Char('C')) => {
                    (0..n).map(|_| AppAction::DeleteCurrentCol).collect()
                }
                (KeyCode::Char('y'), KeyCode::Char('y')) => {
                    (0..n).map(|_| AppAction::YankRow).collect()
                }
                (KeyCode::Char('y'), KeyCode::Char('r')) => {
                    // yr — explicit row yank (same as yy)
                    (0..n).map(|_| AppAction::YankRow).collect()
                }
                (KeyCode::Char('y'), KeyCode::Char('c')) => {
                    // yc — yank current cell only
                    vec![AppAction::YankCell]
                }

                // f{A-Za-z} — jump to column by letter (A=0, B=1, …, Z=25)
                (KeyCode::Char('f'), KeyCode::Char(c)) if c.is_ascii_alphabetic() => {
                    let upper = c.to_ascii_uppercase();
                    let col = (upper as u32).saturating_sub('A' as u32);
                    vec![AppAction::JumpToCol { col }]
                }

                // zz / zt / zb
                (KeyCode::Char('z'), KeyCode::Char('z')) => vec![AppAction::ScrollCenter],
                (KeyCode::Char('z'), KeyCode::Char('t')) => vec![AppAction::ScrollTop],
                (KeyCode::Char('z'), KeyCode::Char('b')) => vec![AppAction::ScrollBottom],

                // >> / <<
                (KeyCode::Char('>'), KeyCode::Char('>')) => {
                    let col = workbook.active().max_col().min(self.cursor.col);
                    (0..n)
                        .map(|_| AppAction::IncreaseColWidth {
                            col: self.cursor.col,
                        })
                        .collect()
                }
                (KeyCode::Char('<'), KeyCode::Char('<')) => (0..n)
                    .map(|_| AppAction::DecreaseColWidth {
                        col: self.cursor.col,
                    })
                    .collect(),

                // m{char} — set mark
                (KeyCode::Char('m'), KeyCode::Char(c)) if c.is_ascii_alphabetic() => {
                    vec![AppAction::SetMark { ch: c }]
                }

                // '{char} — jump to mark, '' — jump to prev position
                (KeyCode::Char('\''), KeyCode::Char('\'')) => vec![AppAction::JumpToPrevPos],
                (KeyCode::Char('\''), KeyCode::Char(c)) if c.is_ascii_alphabetic() => {
                    vec![AppAction::JumpToMark { ch: c }]
                }

                // c{c} — change cell
                (KeyCode::Char('c'), KeyCode::Char('c')) => {
                    self.edit_buffer.clear();
                    self.edit_cursor_pos = 0;
                    vec![AppAction::ChangeCell]
                }

                // yS — yank (copy) the current cell's style
                (KeyCode::Char('y'), KeyCode::Char('S')) => vec![AppAction::YankStyle],

                // pS — paste style to current cell / visual selection
                (KeyCode::Char('p'), KeyCode::Char('S')) => vec![AppAction::PasteStyle],

                // q{char} — start recording macro to register
                (KeyCode::Char('q'), KeyCode::Char(c)) if c.is_ascii_alphabetic() => {
                    vec![AppAction::StartRecording { register: c }]
                }

                // @{char} — play macro from register; @@ — replay last
                (KeyCode::Char('@'), KeyCode::Char('@')) => {
                    let reg = self.last_macro_register.unwrap_or('a');
                    (0..n)
                        .map(|_| AppAction::PlayMacro { register: reg })
                        .collect()
                }
                (KeyCode::Char('@'), KeyCode::Char(c)) if c.is_ascii_alphabetic() => (0..n)
                    .map(|_| AppAction::PlayMacro { register: c })
                    .collect(),

                _ => vec![AppAction::NoOp],
            };
        }

        // ── Count digit accumulation ──────────────────────────────────────────
        if let KeyCode::Char(c) = key.code {
            if c.is_ascii_digit() && (c != '0' || !self.count_buffer.is_empty()) && !ctrl {
                self.count_buffer.push(c);
                return vec![AppAction::NoOp];
            }
        }

        let n = self.take_count();

        // ── Single-key bindings ───────────────────────────────────────────────
        match key.code {
            // ── Movement ──────────────────────────────────────────────────────
            KeyCode::Char('h') | KeyCode::Left => (0..n)
                .map(|_| AppAction::MoveCursor {
                    row_delta: 0,
                    col_delta: -1,
                })
                .collect(),
            KeyCode::Char('j') | KeyCode::Down => (0..n)
                .map(|_| AppAction::MoveCursor {
                    row_delta: 1,
                    col_delta: 0,
                })
                .collect(),
            KeyCode::Char('k') | KeyCode::Up => (0..n)
                .map(|_| AppAction::MoveCursor {
                    row_delta: -1,
                    col_delta: 0,
                })
                .collect(),
            KeyCode::Char('l') | KeyCode::Right => (0..n)
                .map(|_| AppAction::MoveCursor {
                    row_delta: 0,
                    col_delta: 1,
                })
                .collect(),

            // Word/block jumps
            KeyCode::Char('w') => (0..n)
                .map(|_| AppAction::MoveToNextNonEmptyH { forward: true })
                .collect(),
            KeyCode::Char('b') => (0..n)
                .map(|_| AppAction::MoveToNextNonEmptyH { forward: false })
                .collect(),
            KeyCode::Char('W') => (0..n)
                .map(|_| AppAction::MoveToNextNonEmptyV { forward: true })
                .collect(),
            KeyCode::Char('B') => (0..n)
                .map(|_| AppAction::MoveToNextNonEmptyV { forward: false })
                .collect(),
            KeyCode::Char('e') => (0..n)
                .map(|_| AppAction::MoveToNextNonEmptyH { forward: true })
                .collect(),

            // Paragraph jumps
            KeyCode::Char('}') => (0..n)
                .map(|_| AppAction::MoveToNextParagraph { forward: true })
                .collect(),
            KeyCode::Char('{') => (0..n)
                .map(|_| AppAction::MoveToNextParagraph { forward: false })
                .collect(),

            // Row extremes
            KeyCode::Char('0') | KeyCode::Home => vec![AppAction::MoveToFirstCol],
            KeyCode::Char('$') | KeyCode::End => vec![AppAction::MoveToLastCol],

            // File extremes (multi-key: buffer 'g')
            KeyCode::Char('g') => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }
            KeyCode::Char('G') => vec![AppAction::MoveToLastRow],

            // Screen jumps
            KeyCode::Char('H') => vec![AppAction::JumpHighRow],
            KeyCode::Char('M') => vec![AppAction::JumpMidRow],
            KeyCode::Char('L') => vec![AppAction::JumpLowRow],

            // Cell arithmetic
            KeyCode::Char('a') if ctrl => (0..n).map(|_| AppAction::IncrementCell).collect(),
            KeyCode::Char('x') if ctrl => (0..n).map(|_| AppAction::DecrementCell).collect(),

            // Page scroll
            KeyCode::Char('d') if ctrl => (0..n).map(|_| AppAction::PageDown).collect(),
            KeyCode::Char('u') if ctrl => (0..n).map(|_| AppAction::PageUp).collect(),
            KeyCode::Char('f') if ctrl => (0..n).map(|_| AppAction::PageDown).collect(),
            KeyCode::Char('b') if ctrl => (0..n).map(|_| AppAction::PageUp).collect(),
            KeyCode::PageDown => (0..n).map(|_| AppAction::PageDown).collect(),
            KeyCode::PageUp => (0..n).map(|_| AppAction::PageUp).collect(),

            // z prefix — scroll
            KeyCode::Char('z') => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }

            // ── Tab / sheet ───────────────────────────────────────────────────
            KeyCode::Tab => (0..n)
                .map(|_| AppAction::MoveCursor {
                    row_delta: 0,
                    col_delta: 1,
                })
                .collect(),
            KeyCode::BackTab => (0..n)
                .map(|_| AppAction::MoveCursor {
                    row_delta: 0,
                    col_delta: -1,
                })
                .collect(),
            KeyCode::Char('t') if ctrl => vec![AppAction::NextSheet],
            KeyCode::Char('T') if ctrl => vec![AppAction::PrevSheet],

            // ── Editing ───────────────────────────────────────────────────────
            KeyCode::Enter | KeyCode::F(2) => {
                let val = workbook
                    .active()
                    .get_value(self.cursor.row, self.cursor.col);
                self.edit_buffer = val.formula_bar_display();
                self.edit_cursor_pos = self.edit_buffer.len();
                vec![AppAction::EnterInsert { replace: false }]
            }
            KeyCode::Char('i') => {
                let val = workbook
                    .active()
                    .get_value(self.cursor.row, self.cursor.col);
                self.edit_buffer = val.formula_bar_display();
                self.edit_cursor_pos = self.edit_buffer.len();
                vec![AppAction::EnterInsert { replace: false }]
            }
            KeyCode::Char('a') => {
                let val = workbook
                    .active()
                    .get_value(self.cursor.row, self.cursor.col);
                self.edit_buffer = val.formula_bar_display();
                self.edit_cursor_pos = self.edit_buffer.len();
                vec![AppAction::EnterInsert { replace: false }]
            }
            // c — change: buffer for cc, but single c acts on cell too
            KeyCode::Char('c') if !ctrl => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }
            // s — substitute (clear + insert, no need for 2nd key)
            KeyCode::Char('s') => {
                self.edit_buffer.clear();
                self.edit_cursor_pos = 0;
                vec![AppAction::ChangeCell]
            }
            // r — replace mode
            KeyCode::Char('r') if !ctrl => {
                self.edit_buffer.clear();
                self.edit_cursor_pos = 0;
                vec![AppAction::EnterInsert { replace: true }]
            }
            // D — delete cell content
            KeyCode::Char('D') => vec![AppAction::DeleteCellContent],
            // x / Del — delete cell content
            KeyCode::Char('x') | KeyCode::Delete => vec![AppAction::DeleteCellContent],

            // ~ — toggle case
            KeyCode::Char('~') => vec![AppAction::ToggleCase],

            // d prefix — dd or bail
            KeyCode::Char('d') if !ctrl => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }

            // o / O — open row
            KeyCode::Char('o') => vec![AppAction::OpenRowBelow],
            KeyCode::Char('O') => vec![AppAction::OpenRowAbove],

            // = — auto-fit column
            KeyCode::Char('=') => vec![AppAction::AutoFitCol {
                col: self.cursor.col,
            }],

            // + / - — row height; _ — auto-fit (reset) row
            KeyCode::Char('+') => (0..n)
                .map(|_| AppAction::IncreaseRowHeight {
                    row: self.cursor.row,
                })
                .collect(),
            KeyCode::Char('-') => (0..n)
                .map(|_| AppAction::DecreaseRowHeight {
                    row: self.cursor.row,
                })
                .collect(),
            KeyCode::Char('_') => vec![AppAction::AutoFitRow {
                row: self.cursor.row,
            }],

            // f prefix — column jump by letter
            KeyCode::Char('f') if !ctrl => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }

            // J — join: concat cell below into current cell, then clear below
            KeyCode::Char('J') => vec![AppAction::JoinCellBelow],

            // > / < prefix — column width
            KeyCode::Char('>') => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }
            KeyCode::Char('<') => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }

            // ── Yank / paste ──────────────────────────────────────────────────
            KeyCode::Char('y') => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }
            KeyCode::Char('p') => (0..n).map(|_| AppAction::PasteAfter).collect(),
            KeyCode::Char('P') => (0..n).map(|_| AppAction::PasteBefore).collect(),

            // ── Visual modes ──────────────────────────────────────────────────
            KeyCode::Char('v') if !ctrl => {
                self.visual_anchor = Some(VisualAnchor {
                    row: self.cursor.row,
                    col: self.cursor.col,
                });
                vec![AppAction::EnterVisual { block: false }]
            }
            KeyCode::Char('v') if ctrl => {
                self.visual_anchor = Some(VisualAnchor {
                    row: self.cursor.row,
                    col: self.cursor.col,
                });
                vec![AppAction::EnterVisual { block: true }]
            }
            KeyCode::Char('V') => {
                self.visual_anchor = Some(VisualAnchor {
                    row: self.cursor.row,
                    col: self.cursor.col,
                });
                vec![AppAction::EnterVisualLine]
            }

            // ── Undo / redo ───────────────────────────────────────────────────
            KeyCode::Char('u') if !ctrl => (0..n).map(|_| AppAction::Undo).collect(),
            KeyCode::Char('r') if ctrl => (0..n).map(|_| AppAction::Redo).collect(),

            // ── Search ────────────────────────────────────────────────────────
            KeyCode::Char(':') => {
                self.command_buffer.clear();
                vec![AppAction::EnterCommand]
            }
            KeyCode::Char('/') => {
                self.search_buffer.clear();
                vec![AppAction::EnterSearch { forward: true }]
            }
            KeyCode::Char('?') => {
                self.search_buffer.clear();
                vec![AppAction::EnterSearch { forward: false }]
            }
            KeyCode::Char('n') => (0..n).map(|_| AppAction::FindNext).collect(),
            KeyCode::Char('N') => (0..n).map(|_| AppAction::FindPrev).collect(),
            KeyCode::Char('*') => vec![AppAction::SearchCurrentCell],

            // ── Marks ─────────────────────────────────────────────────────────
            KeyCode::Char('m') => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }
            KeyCode::Char('\'') => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }

            // ── Macro recording / playback ─────────────────────────────────────
            // q alone buffers; q{char} starts recording (handled in multi-key block above)
            KeyCode::Char('q') if !ctrl => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }
            // @ alone buffers; @{char} plays macro (handled in multi-key block above)
            KeyCode::Char('@') if !ctrl => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }

            KeyCode::Esc => {
                self.count_buffer.clear();
                vec![AppAction::NoOp]
            }
            _ => vec![AppAction::NoOp],
        }
    }

    // ── Welcome screen ────────────────────────────────────────────────────────

    fn handle_welcome(&mut self, key: KeyEvent) -> Vec<AppAction> {
        match key.code {
            KeyCode::Char('n') | KeyCode::Char('N') => vec![AppAction::WelcomeNewFile],
            KeyCode::Char('f') | KeyCode::Char('F') => vec![AppAction::WelcomeEnterFileFind],
            KeyCode::Char('r') | KeyCode::Char('R') => vec![AppAction::WelcomeEnterRecent],
            KeyCode::Char('t') | KeyCode::Char('T') => vec![AppAction::WelcomeOpenThemes],
            KeyCode::Char('c') | KeyCode::Char('C') => vec![AppAction::WelcomeOpenConfig],
            KeyCode::Char('q') | KeyCode::Char('Q') | KeyCode::Esc => vec![AppAction::Quit],
            _ => vec![AppAction::NoOp],
        }
    }

    // ── Theme manager ─────────────────────────────────────────────────────────

    fn handle_theme_manager(&mut self, key: KeyEvent) -> Vec<AppAction> {
        let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
        match key.code {
            KeyCode::Esc | KeyCode::Char('q') => vec![AppAction::ThemeManagerCancel],
            KeyCode::Enter => vec![AppAction::ThemeApply(self.theme_selected)],
            KeyCode::Up | KeyCode::Char('k') if !ctrl => {
                if self.theme_selected > 0 {
                    self.theme_selected -= 1;
                }
                vec![AppAction::NoOp]
            }
            KeyCode::Down | KeyCode::Char('j') if !ctrl => {
                self.theme_selected += 1; // clamped in process_action after theme count is known
                vec![AppAction::NoOp]
            }
            _ => vec![AppAction::NoOp],
        }
    }

    // ── Fuzzy file finder ─────────────────────────────────────────────────────

    /// Returns files from finder_files filtered by finder_query (case-insensitive substring).
    pub fn filtered_finder_files(&self) -> Vec<&String> {
        let q = self.finder_query.to_ascii_lowercase();
        self.finder_files
            .iter()
            .filter(|f| q.is_empty() || f.to_ascii_lowercase().contains(&q))
            .collect()
    }

    fn handle_file_find(&mut self, key: KeyEvent) -> Vec<AppAction> {
        let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
        match key.code {
            KeyCode::Esc => {
                self.finder_query.clear();
                self.finder_selected = 0;
                self.mode = Mode::Welcome;
                vec![AppAction::NoOp]
            }
            KeyCode::Enter => vec![AppAction::FinderOpen],
            // Navigation: arrow keys only — letters must type into the query
            KeyCode::Up => vec![AppAction::FinderMoveUp],
            KeyCode::Down => vec![AppAction::FinderMoveDown],
            // Ctrl+k / Ctrl+j as alternative nav that doesn't conflict with typing
            KeyCode::Char('k') if ctrl => vec![AppAction::FinderMoveUp],
            KeyCode::Char('j') if ctrl => vec![AppAction::FinderMoveDown],
            KeyCode::Backspace => {
                self.finder_query.pop();
                self.finder_selected = 0;
                vec![AppAction::NoOp]
            }
            // Every other printable character types into the query
            KeyCode::Char(c) if !ctrl => {
                self.finder_query.push(c);
                self.finder_selected = 0;
                vec![AppAction::NoOp]
            }
            _ => vec![AppAction::NoOp],
        }
    }

    // ── Recent files ──────────────────────────────────────────────────────────

    fn handle_recent_files(&mut self, key: KeyEvent) -> Vec<AppAction> {
        let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
        match key.code {
            KeyCode::Esc => {
                self.recent_selected = 0;
                self.mode = Mode::Welcome;
                vec![AppAction::NoOp]
            }
            KeyCode::Enter => vec![AppAction::RecentOpen],
            KeyCode::Up | KeyCode::Char('k') if !ctrl => vec![AppAction::RecentMoveUp],
            KeyCode::Down | KeyCode::Char('j') if !ctrl => vec![AppAction::RecentMoveDown],
            _ => vec![AppAction::NoOp],
        }
    }

    // ── Insert mode ───────────────────────────────────────────────────────────

    // ── Formula reference selection ───────────────────────────────────────────

    fn handle_formula_select(&mut self, key: KeyEvent) -> Vec<AppAction> {
        let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
        match key.code {
            // Confirm — insert the ref and return to Insert
            KeyCode::Enter => vec![AppAction::FormulaSelectConfirm],

            // Cancel — return to Insert without inserting
            KeyCode::Esc => vec![AppAction::FormulaSelectCancel],

            // Start range selection
            KeyCode::Char(':') | KeyCode::Char('.') => vec![AppAction::FormulaSelectStartRange],

            // Navigate with hjkl and arrow keys (grid cursor moves)
            KeyCode::Char('h') | KeyCode::Left => vec![AppAction::MoveCursor {
                row_delta: 0,
                col_delta: -1,
            }],
            KeyCode::Char('j') | KeyCode::Down => vec![AppAction::MoveCursor {
                row_delta: 1,
                col_delta: 0,
            }],
            KeyCode::Char('k') | KeyCode::Up => vec![AppAction::MoveCursor {
                row_delta: -1,
                col_delta: 0,
            }],
            KeyCode::Char('l') | KeyCode::Right => vec![AppAction::MoveCursor {
                row_delta: 0,
                col_delta: 1,
            }],

            // Page navigation
            KeyCode::Char('d') if ctrl => vec![AppAction::PageDown],
            KeyCode::Char('u') if ctrl => vec![AppAction::PageUp],

            // Jump to first/last (gg / G)
            KeyCode::Char('g') if !ctrl => {
                self.key_buffer.push(key);
                vec![AppAction::NoOp]
            }
            KeyCode::Char('G') => vec![AppAction::MoveToLastRow],

            _ => vec![AppAction::NoOp],
        }
    }

    // ── Insert mode ───────────────────────────────────────────────────────────

    fn handle_insert(&mut self, key: KeyEvent, _replace: bool) -> Vec<AppAction> {
        let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
        match key.code {
            // Ctrl+R while editing a formula → enter cell-reference picking mode
            KeyCode::Char('r') if ctrl && self.edit_buffer.starts_with('=') => {
                vec![AppAction::EnterFormulaSelect]
            }

            // Ctrl+V — paste from system clipboard into edit buffer
            KeyCode::Char('v') if ctrl => {
                vec![AppAction::PasteFromClipboard]
            }

            // Tab — when editing a formula (starts with '='), cycle function name completions
            KeyCode::Tab if self.edit_buffer.starts_with('=') => {
                self.cycle_fn_completion(true);
                vec![AppAction::NoOp]
            }
            KeyCode::BackTab if self.edit_buffer.starts_with('=') => {
                self.cycle_fn_completion(false);
                vec![AppAction::NoOp]
            }

            // Ctrl+A — beginning of edit buffer
            KeyCode::Char('a') if ctrl => {
                self.edit_cursor_pos = 0;
                vec![AppAction::NoOp]
            }
            // Ctrl+E — end of edit buffer
            KeyCode::Char('e') if ctrl => {
                self.edit_cursor_pos = self.edit_buffer.len();
                vec![AppAction::NoOp]
            }
            // Ctrl+W — delete word backward
            KeyCode::Char('w') if ctrl => {
                if self.edit_cursor_pos > 0 {
                    let mut pos = self.edit_cursor_pos;
                    // Skip trailing whitespace
                    while pos > 0 {
                        let prev = prev_char_boundary(&self.edit_buffer, pos);
                        if self.edit_buffer[prev..pos]
                            .chars()
                            .next()
                            .map(|c| c.is_whitespace())
                            .unwrap_or(false)
                        {
                            pos = prev;
                        } else {
                            break;
                        }
                    }
                    // Delete word characters
                    while pos > 0 {
                        let prev = prev_char_boundary(&self.edit_buffer, pos);
                        if self.edit_buffer[prev..pos]
                            .chars()
                            .next()
                            .map(|c| c.is_whitespace())
                            .unwrap_or(false)
                        {
                            break;
                        }
                        pos = prev;
                    }
                    self.edit_buffer.drain(pos..self.edit_cursor_pos);
                    self.edit_cursor_pos = pos;
                }
                vec![AppAction::NoOp]
            }
            // Ctrl+U — delete to start of buffer
            KeyCode::Char('u') if ctrl => {
                self.edit_buffer.drain(..self.edit_cursor_pos);
                self.edit_cursor_pos = 0;
                vec![AppAction::NoOp]
            }
            // Ctrl+K — delete to end of buffer
            KeyCode::Char('k') if ctrl => {
                self.edit_buffer.truncate(self.edit_cursor_pos);
                vec![AppAction::NoOp]
            }

            KeyCode::Esc if !ctrl => {
                let value = parse_cell_value(&self.edit_buffer);
                let row = self.cursor.row;
                let col = self.cursor.col;
                self.edit_buffer.clear();
                self.edit_cursor_pos = 0;
                vec![
                    AppAction::SetCell {
                        sheet: 0,
                        row,
                        col,
                        value,
                    },
                    AppAction::ExitMode,
                ]
            }
            KeyCode::Enter => {
                let value = parse_cell_value(&self.edit_buffer);
                let row = self.cursor.row;
                let col = self.cursor.col;
                self.edit_buffer.clear();
                self.edit_cursor_pos = 0;
                vec![
                    AppAction::SetCell {
                        sheet: 0,
                        row,
                        col,
                        value,
                    },
                    AppAction::ExitMode,
                    AppAction::MoveCursor {
                        row_delta: 1,
                        col_delta: 0,
                    },
                ]
            }
            KeyCode::Tab => {
                let value = parse_cell_value(&self.edit_buffer);
                let row = self.cursor.row;
                let col = self.cursor.col;
                self.edit_buffer.clear();
                self.edit_cursor_pos = 0;
                vec![
                    AppAction::SetCell {
                        sheet: 0,
                        row,
                        col,
                        value,
                    },
                    AppAction::ExitMode,
                    AppAction::MoveCursor {
                        row_delta: 0,
                        col_delta: 1,
                    },
                ]
            }
            KeyCode::Backspace => {
                if self.edit_cursor_pos > 0 {
                    let pos = prev_char_boundary(&self.edit_buffer, self.edit_cursor_pos);
                    self.edit_buffer.drain(pos..self.edit_cursor_pos);
                    self.edit_cursor_pos = pos;
                }
                vec![AppAction::NoOp]
            }
            KeyCode::Delete => {
                if self.edit_cursor_pos < self.edit_buffer.len() {
                    let end = next_char_boundary(&self.edit_buffer, self.edit_cursor_pos);
                    self.edit_buffer.drain(self.edit_cursor_pos..end);
                }
                vec![AppAction::NoOp]
            }
            KeyCode::Left => {
                if self.edit_cursor_pos > 0 {
                    self.edit_cursor_pos =
                        prev_char_boundary(&self.edit_buffer, self.edit_cursor_pos);
                }
                vec![AppAction::NoOp]
            }
            KeyCode::Right => {
                if self.edit_cursor_pos < self.edit_buffer.len() {
                    self.edit_cursor_pos =
                        next_char_boundary(&self.edit_buffer, self.edit_cursor_pos);
                }
                vec![AppAction::NoOp]
            }
            KeyCode::Home => {
                self.edit_cursor_pos = 0;
                vec![AppAction::NoOp]
            }
            KeyCode::End => {
                self.edit_cursor_pos = self.edit_buffer.len();
                vec![AppAction::NoOp]
            }
            KeyCode::Char(c) => {
                // Reset formula completion on any non-tab char
                self.fn_completion_idx = None;
                self.fn_completion_prefix.clear();
                self.fn_completion_candidates.clear();
                self.edit_buffer.insert(self.edit_cursor_pos, c);
                self.edit_cursor_pos += c.len_utf8();
                vec![AppAction::NoOp]
            }
            _ => vec![AppAction::NoOp],
        }
    }

    // ── Visual mode (shared for Visual and VisualLine) ────────────────────────

    fn handle_visual(&mut self, key: KeyEvent, is_line: bool) -> Vec<AppAction> {
        let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);

        // Ctrl-key overrides must come before the plain char matches below
        if ctrl {
            match key.code {
                // Fill down — Ctrl+D
                KeyCode::Char('d') => {
                    let (row_start, col_start, row_end, col_end) = self.visual_selection_bounds();
                    let col_end_c = col_end.min(1000);
                    self.visual_anchor = None;
                    return vec![
                        AppAction::FillDown {
                            anchor_row: row_start,
                            col_start,
                            col_end: col_end_c,
                            row_end,
                        },
                        AppAction::ExitMode,
                    ];
                }
                // Fill right — Ctrl+R
                KeyCode::Char('r') => {
                    let (row_start, col_start, row_end, col_end) = self.visual_selection_bounds();
                    let row_end_c = row_end.min(100_000);
                    self.visual_anchor = None;
                    return vec![
                        AppAction::FillRight {
                            anchor_col: col_start,
                            row_start,
                            row_end: row_end_c,
                            col_end,
                        },
                        AppAction::ExitMode,
                    ];
                }
                _ => {}
            }
        }

        match key.code {
            KeyCode::Esc => {
                self.visual_anchor = None;
                vec![AppAction::ExitMode]
            }

            // Toggle modes
            KeyCode::Char('v') if !is_line => {
                self.visual_anchor = None;
                vec![AppAction::ExitMode]
            }
            KeyCode::Char('V') if is_line => {
                self.visual_anchor = None;
                vec![AppAction::ExitMode]
            }
            KeyCode::Char('V') if !is_line => vec![AppAction::EnterVisualLine],
            KeyCode::Char('v') if is_line => vec![AppAction::EnterVisual { block: false }],

            // Movement
            KeyCode::Char('h') | KeyCode::Left => vec![AppAction::MoveCursor {
                row_delta: 0,
                col_delta: -1,
            }],
            KeyCode::Char('j') | KeyCode::Down => vec![AppAction::MoveCursor {
                row_delta: 1,
                col_delta: 0,
            }],
            KeyCode::Char('k') | KeyCode::Up => vec![AppAction::MoveCursor {
                row_delta: -1,
                col_delta: 0,
            }],
            KeyCode::Char('l') | KeyCode::Right => vec![AppAction::MoveCursor {
                row_delta: 0,
                col_delta: 1,
            }],
            KeyCode::Char('w') => vec![AppAction::MoveToNextNonEmptyH { forward: true }],
            KeyCode::Char('b') => vec![AppAction::MoveToNextNonEmptyH { forward: false }],
            KeyCode::Char('W') => vec![AppAction::MoveToNextNonEmptyV { forward: true }],
            KeyCode::Char('B') => vec![AppAction::MoveToNextNonEmptyV { forward: false }],
            KeyCode::Char('}') => vec![AppAction::MoveToNextParagraph { forward: true }],
            KeyCode::Char('{') => vec![AppAction::MoveToNextParagraph { forward: false }],
            KeyCode::Char('0') | KeyCode::Home => vec![AppAction::MoveToFirstCol],
            KeyCode::Char('$') | KeyCode::End => vec![AppAction::MoveToLastCol],
            KeyCode::Char('g') => vec![AppAction::MoveToFirstRow],
            KeyCode::Char('G') => vec![AppAction::MoveToLastRow],

            // Delete
            KeyCode::Char('d') | KeyCode::Char('x') | KeyCode::Delete => {
                let (row_start, col_start, row_end, col_end) = self.visual_selection_bounds();
                self.visual_anchor = None;
                vec![
                    AppAction::DeleteCellRange {
                        row_start,
                        col_start,
                        row_end,
                        col_end,
                    },
                    AppAction::ExitMode,
                ]
            }

            // Change (clear + insert) — ExitMode must come last so Insert wins
            KeyCode::Char('c') | KeyCode::Char('s') => {
                let (row_start, col_start, row_end, col_end) = self.visual_selection_bounds();
                self.visual_anchor = None;
                vec![
                    AppAction::DeleteCellRange {
                        row_start,
                        col_start,
                        row_end,
                        col_end,
                    },
                    AppAction::VisualClearAndInsert,
                ]
            }

            // S — insert =SUM(range) into cell below / right of selection
            KeyCode::Char('S') => {
                let (row_start, col_start, row_end, col_end) = self.visual_selection_bounds();
                self.visual_anchor = None;
                vec![
                    AppAction::InsertSumBelow {
                        row_start,
                        col_start,
                        row_end,
                        col_end,
                    },
                    AppAction::ExitMode,
                ]
            }

            // >> / << — widen/narrow all columns in selection
            KeyCode::Char('>') => {
                let (_, col_start, _, col_end) = self.visual_selection_bounds();
                self.visual_anchor = None;
                // Clamp col_end to a sane maximum (u32::MAX can appear in V-COL mode)
                let col_end = col_end.min(1000);
                let mut actions: Vec<AppAction> = (col_start..=col_end)
                    .map(|c| AppAction::IncreaseColWidth { col: c })
                    .collect();
                actions.push(AppAction::ExitMode);
                actions
            }
            KeyCode::Char('<') => {
                let (_, col_start, _, col_end) = self.visual_selection_bounds();
                self.visual_anchor = None;
                let col_end = col_end.min(1000);
                let mut actions: Vec<AppAction> = (col_start..=col_end)
                    .map(|c| AppAction::DecreaseColWidth { col: c })
                    .collect();
                actions.push(AppAction::ExitMode);
                actions
            }

            // Yank
            KeyCode::Char('y') => {
                let (row_start, col_start, row_end, col_end) = self.visual_selection_bounds();
                self.visual_anchor = None;
                vec![
                    AppAction::YankCellRange {
                        row_start,
                        col_start,
                        row_end,
                        col_end,
                        is_line,
                    },
                    AppAction::ExitMode,
                ]
            }

            _ => vec![AppAction::NoOp],
        }
    }

    // ── Command mode ──────────────────────────────────────────────────────────

    fn handle_command(&mut self, key: KeyEvent) -> Vec<AppAction> {
        match key.code {
            KeyCode::Esc => {
                self.command_buffer.clear();
                self.completion_idx = None;
                self.completion_prefix.clear();
                vec![AppAction::ExitMode]
            }
            KeyCode::Enter => {
                let cmd = self.command_buffer.clone();
                self.command_buffer.clear();
                self.completion_idx = None;
                self.completion_prefix.clear();
                vec![AppAction::ExecuteCommand2(cmd), AppAction::ExitMode]
            }
            KeyCode::Backspace => {
                self.completion_idx = None;
                self.completion_prefix.clear();
                self.subcmd_completion_idx = None;
                if self.command_buffer.is_empty() {
                    vec![AppAction::ExitMode]
                } else {
                    self.command_buffer.pop();
                    vec![AppAction::NoOp]
                }
            }
            // Tab / Shift-Tab: cycle through completions
            KeyCode::Tab | KeyCode::BackTab => {
                let forward = key.code == KeyCode::Tab;
                let buf = self.command_buffer.clone();

                // ── Sub-command completion (buffer has a space → verb done) ───
                if let Some(space_idx) = buf.find(' ') {
                    let completions = self.subcmd_completions.clone();
                    if completions.is_empty() {
                        return vec![AppAction::NoOp];
                    }
                    let len = completions.len();
                    let next = match self.subcmd_completion_idx {
                        None => {
                            if forward {
                                0
                            } else {
                                len - 1
                            }
                        }
                        Some(i) => {
                            if forward {
                                (i + 1) % len
                            } else if i == 0 {
                                len - 1
                            } else {
                                i - 1
                            }
                        }
                    };
                    self.subcmd_completion_idx = Some(next);
                    // Replace just the argument part
                    let verb_part = &buf[..=space_idx]; // "theme "
                    self.command_buffer = format!("{}{}", verb_part, completions[next]);
                    return vec![AppAction::NoOp];
                }

                // ── Verb-level completion ─────────────────────────────────────
                // First Tab press in a session: capture the current prefix
                if self.completion_idx.is_none() {
                    self.completion_prefix = self
                        .command_buffer
                        .split_whitespace()
                        .next()
                        .unwrap_or("")
                        .to_string();
                }

                let matches = get_command_completions(&self.completion_prefix);
                if matches.is_empty() {
                    return vec![AppAction::NoOp];
                }

                let len = matches.len();
                let next = match self.completion_idx {
                    None => {
                        if forward {
                            0
                        } else {
                            len - 1
                        }
                    }
                    Some(i) => {
                        if forward {
                            (i + 1) % len
                        } else if i == 0 {
                            len - 1
                        } else {
                            i - 1
                        }
                    }
                };
                self.completion_idx = Some(next);

                // Fill command buffer with the selected command's name (word part only)
                let chosen = matches[next].0;
                let word = chosen
                    .split(|c: char| c == ' ' || c == '<')
                    .next()
                    .unwrap_or(chosen);
                self.command_buffer = word.to_string();

                vec![AppAction::NoOp]
            }
            KeyCode::Char(c) => {
                // Any typed char resets both completion sessions
                self.completion_idx = None;
                self.completion_prefix.clear();
                self.subcmd_completion_idx = None;
                self.command_buffer.push(c);
                vec![AppAction::NoOp]
            }
            _ => vec![AppAction::NoOp],
        }
    }

    // ── Search mode ───────────────────────────────────────────────────────────

    fn handle_search(&mut self, key: KeyEvent, forward: bool) -> Vec<AppAction> {
        match key.code {
            KeyCode::Esc => {
                self.search_buffer.clear();
                vec![AppAction::ExitMode]
            }
            KeyCode::Enter => {
                let pattern = self.search_buffer.clone();
                self.last_search = Some((pattern, forward));
                self.search_buffer.clear();
                vec![AppAction::ExecuteSearch, AppAction::ExitMode]
            }
            KeyCode::Backspace => {
                self.search_buffer.pop();
                vec![AppAction::NoOp]
            }
            KeyCode::Char(c) => {
                self.search_buffer.push(c);
                vec![AppAction::NoOp]
            }
            _ => vec![AppAction::NoOp],
        }
    }
}

impl Default for InputState {
    fn default() -> Self {
        InputState::new()
    }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

pub fn parse_cell_value(s: &str) -> CellValue {
    if s.is_empty() {
        return CellValue::Empty;
    }
    if let Some(f) = s.strip_prefix('=') {
        return CellValue::Formula(f.to_string());
    }
    if let Ok(n) = s.parse::<f64>() {
        return CellValue::Number(n);
    }
    match s.to_uppercase().as_str() {
        "TRUE" => return CellValue::Boolean(true),
        "FALSE" => return CellValue::Boolean(false),
        _ => {}
    }
    CellValue::Text(s.to_string())
}

fn key_event_to_str(k: &KeyEvent) -> String {
    let ctrl = k.modifiers.contains(KeyModifiers::CONTROL);
    match k.code {
        KeyCode::Char(c) if ctrl => format!("^{}", c),
        KeyCode::Char(c) => c.to_string(),
        KeyCode::Esc => "<Esc>".into(),
        KeyCode::Enter => "<CR>".into(),
        KeyCode::Tab => "<Tab>".into(),
        KeyCode::BackTab => "<S-Tab>".into(),
        KeyCode::Left => "←".into(),
        KeyCode::Right => "→".into(),
        KeyCode::Up => "↑".into(),
        KeyCode::Down => "↓".into(),
        _ => "?".into(),
    }
}

fn prev_char_boundary(s: &str, pos: usize) -> usize {
    let mut p = pos.saturating_sub(1);
    while p > 0 && !s.is_char_boundary(p) {
        p -= 1;
    }
    p
}

fn next_char_boundary(s: &str, pos: usize) -> usize {
    let mut p = pos + 1;
    while p < s.len() && !s.is_char_boundary(p) {
        p += 1;
    }
    p.min(s.len())
}
