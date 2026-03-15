# ASAT — A Spreadsheet And Terminal

> Terminal spreadsheet editor for Vim users. Modal editing (Normal/Insert/Visual/Command), 30+ live formulas, multi-sheet workbooks, CSV · XLSX · ODS support, full undo stack, system clipboard, macros, and marks. Written in Rust with ratatui.

[![License: GPL-3.0](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Built with Rust](https://img.shields.io/badge/built%20with-Rust-orange.svg)](https://www.rust-lang.org)

**Website:** https://okt4v.github.io/ASAT/

---

## Features

- **Modal editing** — Normal, Insert, Visual (char/line/block), Command, Search, and Macro Recording modes, exactly like Vim
- **Live formula engine** — 30+ built-in functions across math, text, logic, and lookup. Formulas re-evaluate after every edit
- **Multi-sheet workbooks** — tab bar, `:tabnew`, `:tabclose`, `gt` / `gT` to switch sheets
- **File format support** — read and write CSV, TSV, XLSX, and ODS; native `.asat` format with bincode + zstd compression
- **System clipboard** — all yank operations (`yy`, `yc`, visual `y`) copy to your OS clipboard as tab-separated text
- **Full undo stack** — 1000-deep undo/redo covering cell edits, row/column operations, pastes, and style changes
- **Macros** — record key sequences to named registers (`qa` … `qz`), replay with `@a`, chain with `{N}@a`
- **Marks** — set named positions (`ma`), jump back (`'a`), and swap with `''`
- **Cell styling** — bold, italic, underline, strikethrough, foreground/background colour, alignment, and number formats
- **Themes** — built-in theme picker (`:theme`) with multiple colour presets, saved to config
- **Formula reference picker** — press `Ctrl+R` inside a formula to navigate the grid and insert cell/range references interactively

---

## Installation

### Prerequisites

- [Rust](https://rustup.rs/) 1.75 or later

### Build from source

```bash
git clone https://github.com/okt4v/ASAT.git
cd ASAT
cargo build --release
```

The binary is at `target/release/asat`. Copy it somewhere on your `$PATH`:

```bash
cp target/release/asat ~/.local/bin/
```

### Run

```bash
asat                   # open welcome screen
asat file.csv          # open a CSV file
asat budget.xlsx       # open an Excel workbook
asat report.ods        # open an ODS file
asat new_file.csv      # create a new file at this path
```

---

## Keybind Reference

### Normal Mode — Navigation

| Key | Action |
|-----|--------|
| `h` / `←` | Move left |
| `j` / `↓` | Move down |
| `k` / `↑` | Move up |
| `l` / `→` | Move right |
| `w` / `b` | Jump to next / previous non-empty cell (horizontal) |
| `W` / `B` | Jump to next / previous non-empty cell (vertical) |
| `}` / `{` | Jump to next / previous paragraph boundary (empty row) |
| `0` / `Home` | Jump to first column |
| `$` / `End` | Jump to last column with data |
| `gg` | Jump to first row |
| `G` | Jump to last row with data |
| `H` / `M` / `L` | Jump to top / middle / bottom of visible area |
| `f{A-Z}` | Jump to column by letter (`fC` → column C) |
| `Ctrl+d` / `Ctrl+f` / `PgDn` | Page down |
| `Ctrl+u` / `Ctrl+b` / `PgUp` | Page up |
| `zz` / `zt` / `zb` | Scroll cursor to centre / top / bottom of screen |
| `{N}` + motion | Repeat motion N times (e.g. `5j` moves down 5 rows) |

### Normal Mode — Editing

| Key | Action |
|-----|--------|
| `i` / `Enter` / `F2` | Edit current cell |
| `a` | Edit current cell (append) |
| `s` / `cc` | Clear cell and enter insert mode |
| `r` | Enter replace mode |
| `x` / `Del` / `D` | Delete cell content |
| `~` | Toggle case of text cell |
| `J` | Join cell below into current cell (space-separated), clear below |
| `Ctrl+a` | Increment number in cell by 1 |
| `Ctrl+x` | Decrement number in cell by 1 |
| `o` / `O` | Insert row below / above and enter insert mode |
| `u` | Undo |
| `Ctrl+r` | Redo |

### Normal Mode — Yank & Paste

| Key | Action |
|-----|--------|
| `yy` / `yr` | Yank current row → register + system clipboard |
| `yc` | Yank current cell → register + system clipboard |
| `yS` | Copy current cell's style to style clipboard |
| `p` / `P` | Paste after / before cursor |
| `pS` | Paste style clipboard onto current cell |

### Normal Mode — Rows & Columns

| Key | Action |
|-----|--------|
| `dd` | Delete current row |
| `dC` | Delete current column |
| `>>` / `<<` | Increase / decrease current column width |
| `=` | Auto-fit current column to widest content |
| `+` / `-` | Increase / decrease current row height |
| `_` | Reset row height to auto |

### Insert Mode (while editing a cell)

| Key | Action |
|-----|--------|
| `Esc` | Confirm edit, return to Normal |
| `Enter` | Confirm edit, move down |
| `Tab` | Confirm edit, move right |
| `Ctrl+a` | Move edit cursor to start |
| `Ctrl+e` | Move edit cursor to end |
| `Ctrl+w` | Delete word backward |
| `Ctrl+u` | Delete to start of buffer |
| `Ctrl+k` | Delete to end of buffer |
| `Ctrl+r` | Enter F-REF mode to pick a cell reference (formulas only) |

### Visual Mode

| Key | Action |
|-----|--------|
| `v` | Character/cell visual mode |
| `V` | Line (full-row) visual mode |
| `Ctrl+v` | Column block visual mode |
| `d` / `x` / `Del` | Delete selection |
| `c` / `s` | Clear selection and enter insert mode |
| `y` | Yank selection → register + system clipboard (TSV) |
| `S` | Insert `=SUM(range)` below the selection |
| `>` / `<` | Widen / narrow all columns in selection |

### Marks & Macros

| Key | Action |
|-----|--------|
| `m{a-z}` | Set named mark |
| `'{a-z}` | Jump to named mark |
| `''` | Jump back to position before last jump |
| `q{a-z}` | Start recording macro to register |
| `q` | Stop recording |
| `@{a-z}` | Play macro from register |
| `@@` | Replay last macro |
| `{N}@{reg}` | Play macro N times |

### Search

| Key | Action |
|-----|--------|
| `/` | Forward search (regex supported) |
| `?` | Backward search |
| `n` / `N` | Next / previous match |
| `*` | Search for content of current cell |

### Sheets

| Key | Action |
|-----|--------|
| `Tab` / `gt` / `Ctrl+t` | Next sheet |
| `Shift+Tab` / `gT` / `Ctrl+T` | Previous sheet |

---

## Ex Commands

Enter command mode with `:`.

| Command | Action |
|---------|--------|
| `:w [file]` | Save (optionally to a new path) |
| `:q` / `:q!` | Quit / force quit |
| `:wq` / `:x` | Save and quit |
| `:e <file>` | Open file |
| `:tabnew [name]` | New sheet |
| `:tabclose` | Close current sheet |
| `:ic` / `:icr` | Insert column left / right |
| `:dc` | Delete current column |
| `:ir [N]` | Insert row at cursor or line N |
| `:dr [N]` | Delete row at cursor or line N |
| `:cw <N>` | Set column width to N characters |
| `:rh <N>` | Set row height to N lines |
| `:bold` / `:italic` / `:underline` / `:strike` | Toggle text style |
| `:fg <color>` / `:bg <color>` | Set foreground / background colour (hex or named) |
| `:hl <color>` / `:hl` | Highlight cell / clear highlight |
| `:align <l/c/r>` | Set alignment: left, center, or right |
| `:fmt <spec>` | Number format: `%`, `$`, `0.00`, `int`, `date`, `none` |
| `:cs` | Clear all styles |
| `:theme [name]` | Open theme picker or apply a theme by name |
| `:sort [asc\|desc]` | Sort all rows by the current cursor column (undoable) |
| `:s/pat/repl/[g][i]` | Find & replace in text cells — `g` = all occurrences, `i` = case-insensitive (undoable) |

Tab-completion works in command mode — press `Tab` to cycle through matching commands.

---

## Formula Engine

Start any cell with `=` to write a formula. Formulas re-evaluate automatically after every edit.

```
=SUM(A1:A10)
=IF(B2>100, "Over budget", "OK")
=AVERAGE(C1:C20) * 1.1
=CONCATENATE(A1, " ", B1)
```

**Supported functions:**

| Category | Functions |
|----------|-----------|
| Math | `SUM`, `AVERAGE`, `MIN`, `MAX`, `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `FLOOR`, `CEILING`, `MOD`, `POWER`, `SQRT`, `LN`, `LOG`, `LOG10`, `EXP`, `INT`, `TRUNC`, `SIGN` |
| Text | `LEN`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `UPPER`, `LOWER`, `PROPER`, `CONCATENATE`, `TEXT`, `VALUE`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `REPT` |
| Logic | `IF`, `AND`, `OR`, `NOT`, `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISERROR`, `IFERROR`, `ISLOGICAL` |
| Lookup | `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `OFFSET` |
| Date | `NOW`, `TODAY`, `DATE`, `YEAR`, `MONTH`, `DAY` |
| Constants | `TRUE`, `FALSE`, `PI()` |

**Reference syntax:**

| Syntax | Meaning |
|--------|---------|
| `A1` | Relative cell reference |
| `$A$1` | Absolute cell reference |
| `A1:B10` | Range |
| `Sheet2.C4` | Cross-sheet reference |

**Interactive reference picking:** While editing a formula, press `Ctrl+R` to enter F-REF mode. Navigate with `hjkl`, press `:` to start a range, and `Enter` to insert the reference back into your formula.

---

## Configuration

ASAT reads a config file on startup. The default location is:

- Linux/macOS: `~/.config/asat/config.toml`

Open the config from the welcome screen with `c`, or edit it directly. Changes to the theme take effect immediately; other options require a restart.

---

## File Formats

| Format | Read | Write | Notes |
|--------|------|-------|-------|
| `.csv` / `.tsv` | ✅ | ✅ | Active sheet only |
| `.xlsx` | ✅ | ✅ | Multi-sheet, via calamine + rust_xlsxwriter |
| `.xls` / `.xlsm` | ✅ | — | Read via calamine |
| `.ods` | ✅ | ✅ | Multi-sheet OpenDocument Spreadsheet |
| `.asat` | ✅ | ✅ | Native format — bincode + zstd, multi-sheet |

---

## Crate Architecture

ASAT is structured as a Cargo workspace with focused, single-responsibility crates:

```
crates/
  asat-core/      — Workbook, Sheet, Cell, CellValue, CellStyle (no internal deps)
  asat-formula/   — Lexer, parser (AST), evaluator, 30+ built-in functions
  asat-io/        — CSV, XLSX, ODS, .asat drivers (calamine read / rust_xlsxwriter write)
  asat-tui/       — ratatui widgets: grid, formula bar, status bar, tab bar, command line
  asat-input/     — Modal state machine, InputState, AppAction enum
  asat-commands/  — Command trait, UndoStack, SetCell, InsertRow/Col, DeleteRow/Col
  asat-plugins/   — Plugin manager stub + PluginEvent (PyO3 in a future release)
  asat-config/    — Config struct, config.toml parsing, ThemeConfig
  asat/           — Binary: main loop, AppAction dispatch, ex-command handler
```

---

## Roadmap

- [x] MVP — CSV, navigation, insert, undo, `:w` / `:q`
- [x] Full Vim feel — `dd`/`yy`/`p`, marks, registers, visual mode, macros
- [x] Multi-sheet + XLSX / ODS
- [x] Formula engine — 30+ functions, live evaluation, F-REF picker
- [x] Styles + formatting — bold, italic, colour, alignment, number formats
- [ ] Plugin system — PyO3, `init.py`, `asat.function` / `asat.on` / `asat.bind`
- [ ] Polish — sort/filter, autosave, ODS formula round-trip, column freeze

---

## License

GNU General Public License v3.0 — see [LICENSE](LICENSE) for details.

Any fork or modified version distributed publicly must remain open source, credit the original authors, and use the same license.
