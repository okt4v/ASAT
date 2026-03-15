# ASAT — A Spreadsheet And Terminal

> Terminal spreadsheet editor for Vim users. Modal editing (Normal/Insert/Visual/Command), 40+ live formulas, multi-sheet workbooks, CSV · XLSX · ODS support, full undo stack, system clipboard, named ranges, filter/freeze panes, cell notes, conditional formatting, macros, and marks. Written in Rust with ratatui.

[![License: GPL-3.0](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Built with Rust](https://img.shields.io/badge/built%20with-Rust-orange.svg)](https://www.rust-lang.org)

**Website:** https://okt4v.github.io/ASAT/

---

## Features

- **Modal editing** — Normal, Insert, Visual (char/line/block), Command, Search, and Macro Recording modes, exactly like Vim
- **Live formula engine** — 40+ built-in functions across math, text, logic, lookup, statistical, and finance. Includes `AVERAGEIF`, `MAXIFS`, `MINIFS`, `RANK`, `PERCENTILE`, `QUARTILE`, `XLOOKUP`, and `CHOOSE`. Formulas re-evaluate after every edit (lazy dirty-cell tracking)
- **Named ranges** — `:name SALES A1:C10` defines a named range usable in formulas as `=SUM(SALES)`
- **Multi-sheet workbooks** — tab bar, `:tabnew`, `:tabclose`, `gt` / `gT` to switch sheets
- **File format support** — read and write CSV, TSV, XLSX, and ODS (OpenDocument Spreadsheet); native `.asat` format with bincode + zstd compression; ODS formula round-trip with cached computed values
- **System clipboard** — yank (`yy`, `yc`, visual `y`) copies as TSV; `Ctrl+V` in Insert mode pastes from system clipboard
- **Full undo stack** — 1000-deep undo/redo covering cell edits, row/column operations, pastes, and style changes
- **Sort & find/replace** — `:sort asc/desc` sorts rows by the cursor column; `:s/pat/repl/g` does regex find & replace across cells; both undoable
- **Filter rows** — `:filter <col> <op> <val>` hides non-matching rows (supports `=`, `!=`, `>`, `<`, `>=`, `<=`, `~`); `:filter off` restores all rows
- **Freeze panes** — frozen rows and columns render as sticky headers with a visual separator; set via `:freeze rows N` / `:freeze cols N`
- **Fill down / right** — `Ctrl+D` / `Ctrl+R` in Visual mode; also `:filldown` / `:fillright` ex-commands
- **Goto cell** — `g<letter>` jumps to a column; `:goto B15` jumps to any cell address
- **Transpose** — `:transpose` swaps rows and columns in the visual selection
- **Remove duplicates** — `:dedup` removes duplicate rows by the current cursor column
- **Cell notes** — `:note <text>` attaches a comment to the current cell; cells with notes show a `▸` corner marker; `:note` with no argument shows the current note; `:note!` clears it
- **Conditional formatting** — `:colfmt <op> <val> <color>` applies a background colour to all cells in the column matching the condition (undoable)
- **Thousands separator** — `:fmt thousands` formats numbers with comma separators (`#,##0`); `:fmt t2` adds two decimal places
- **Formula tab-completion** — press `Tab` while typing a formula (`=SU…`) to cycle through matching function names
- **Time-based autosave** — configurable autosave interval in seconds (edit `autosave_interval` in `config.toml`; 0 = disabled)
- **Macros** — record key sequences to named registers (`qa` … `qz`), replay with `@a`, chain with `{N}@a`
- **Marks** — set named positions (`ma`), jump back (`'a`), and swap with `''`
- **Cell styling** — bold, italic, underline, strikethrough, foreground/background colour, alignment, and number formats
- **Themes** — built-in theme picker (`:theme`) with multiple colour presets, saved to config
- **Formula reference picker** — press `Ctrl+R` inside a formula to navigate the grid and insert cell/range references interactively
- **Plugin system** — extend ASAT with Python via PyO3; hook into cell changes, mode transitions, and file events; register custom formula functions from `~/.config/asat/init.py`

---

## Installation

### Pre-built binaries (GitHub Releases)

Download a binary for your platform from the [v0.1.3 release](https://github.com/okt4v/ASAT/releases/tag/v0.1.3):

| Platform | Link |
|----------|------|
| Linux x86_64 (glibc) | [asat-x86_64-unknown-linux-gnu.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.3/asat-x86_64-unknown-linux-gnu.tar.gz) |
| Linux x86_64 (musl)  | [asat-x86_64-unknown-linux-musl.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.3/asat-x86_64-unknown-linux-musl.tar.gz) |
| Linux aarch64        | [asat-aarch64-unknown-linux-gnu.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.3/asat-aarch64-unknown-linux-gnu.tar.gz) |
| macOS arm64          | [asat-aarch64-apple-darwin.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.3/asat-aarch64-apple-darwin.tar.gz) |
| macOS x86_64         | [asat-x86_64-apple-darwin.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.3/asat-x86_64-apple-darwin.tar.gz) |
| Windows x86_64       | [asat-x86_64-pc-windows-msvc.zip](https://github.com/okt4v/ASAT/releases/download/v0.1.3/asat-x86_64-pc-windows-msvc.zip) |

Extract the archive and place the `asat` binary somewhere on your `$PATH` (e.g. `~/.local/bin/`).

### Arch Linux (AUR)

```bash
yay -S asat-bin
# or
paru -S asat-bin
```

### Homebrew (macOS / Linux)

```bash
brew tap okt4v/tap
brew install asat
```

### Debian / Ubuntu (.deb)

```bash
wget https://github.com/okt4v/ASAT/releases/download/v0.1.3/asat_0.1.0-1_amd64.deb
sudo dpkg -i asat_0.1.0-1_amd64.deb
```

### Build from source

```bash
git clone https://github.com/okt4v/ASAT.git
cd ASAT
bash install.sh
```

The install script builds in release mode and copies the binary to `~/.local/bin/asat`. It will warn you if that directory is not on your `$PATH`.

To build manually instead:

```bash
cargo build --release
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
| `Ctrl+r` | Enter F-REF mode to pick a cell/range reference (formulas only) |
| `Ctrl+v` | Paste from system clipboard into edit buffer |
| `Tab` / `Shift+Tab` | Cycle through formula function name completions (when buffer starts with `=`) |

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
| `Ctrl+d` | Fill down — copy anchor row to all selected rows |
| `Ctrl+r` | Fill right — copy anchor column to all selected columns |

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
| `:goto <addr>` / `:go <addr>` | Jump to cell address (e.g. `:goto B15`) |
| `:name <NAME> <range>` | Define a named range (e.g. `:name SALES A1:C10`) |
| `:filter <col> <op> <val>` | Hide rows where column does not match (ops: `=` `!=` `>` `<` `>=` `<=` `~`) |
| `:filter off` | Unhide all filtered rows |
| `:transpose` / `:tp` | Transpose the visual selection (swap rows and columns) |
| `:dedup` | Remove duplicate rows by the current cursor column |
| `:note [text]` | Set a note on current cell; `:note` shows it; `:note!` clears it |
| `:colfmt <op> <val> <color>` / `:cf` | Conditional format — apply background colour to matching cells in column |
| `:filldown` / `:fd` | Fill the cursor cell value down to the selection end |
| `:fillright` / `:fr` | Fill the cursor cell value right to the selection end |
| `:fmt thousands` / `:fmt t2` | Thousands-separator number format (`#,##0` or `#,##0.00`) |
| `:freeze rows <N>` | Freeze top N rows as sticky header |
| `:freeze cols <N>` | Freeze left N columns as sticky header |
| `:freeze off` | Clear all frozen panes |

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
| Lookup | `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `CHOOSE` |
| Date | `NOW`, `TODAY`, `DATE`, `YEAR`, `MONTH`, `DAY` |
| Statistical | `COUNT`, `COUNTA`, `SUMIF`, `COUNTIF`, `AVERAGEIF`, `MAXIFS`, `MINIFS`, `SUMPRODUCT`, `MEDIAN`, `STDEV`, `VAR`, `LARGE`, `SMALL`, `RANK`, `PERCENTILE`, `QUARTILE` |
| Finance | `PV`, `FV`, `PMT`, `NPER`, `RATE`, `NPV`, `IRR`, `MIRR`, `IPMT`, `PPMT`, `SLN`, `DDB`, `EFFECT`, `NOMINAL`, `CUMIPMT`, `CUMPRINC` |
| Constants | `TRUE`, `FALSE`, `PI()` |

**Reference syntax:**

| Syntax | Meaning |
|--------|---------|
| `A1` | Relative cell reference |
| `$A$1` | Absolute cell reference |
| `A1:B10` | Range |
| `Sheet2.C4` | Cross-sheet reference |

**Interactive reference picking:** While editing a formula, press `Ctrl+R` to enter F-REF mode. Navigate with `hjkl`/arrows. Press `v` (or `:`) to anchor a range start, then move to the end cell and press `Enter` to insert the range reference (e.g. `A1:C5`). Press `Enter` without anchoring to insert a single cell reference. Press `Esc` to cancel. The formula bar continues to show your in-progress formula throughout.

### Financial Functions

| Function | Signature | Description |
|----------|-----------|-------------|
| `PV` | `PV(rate, nper, pmt, [fv], [type])` | Present value of an investment or annuity |
| `FV` | `FV(rate, nper, pmt, [pv], [type])` | Future value of an investment or annuity |
| `PMT` | `PMT(rate, nper, pv, [fv], [type])` | Periodic payment for a loan or annuity |
| `NPER` | `NPER(rate, pmt, pv, [fv], [type])` | Number of periods to pay off a loan |
| `RATE` | `RATE(nper, pmt, pv, [fv], [type], [guess])` | Interest rate per period (iterative) |
| `NPV` | `NPV(rate, cf1, cf2, …)` | Net present value of a series of cashflows |
| `IRR` | `IRR(values, [guess])` | Internal rate of return (iterative) |
| `MIRR` | `MIRR(values, finance_rate, reinvest_rate)` | Modified internal rate of return |
| `IPMT` | `IPMT(rate, per, nper, pv, [fv], [type])` | Interest portion of a payment |
| `PPMT` | `PPMT(rate, per, nper, pv, [fv], [type])` | Principal portion of a payment |
| `CUMIPMT` | `CUMIPMT(rate, nper, pv, start, end, type)` | Cumulative interest paid over a range of periods |
| `CUMPRINC` | `CUMPRINC(rate, nper, pv, start, end, type)` | Cumulative principal paid over a range of periods |
| `SLN` | `SLN(cost, salvage, life)` | Straight-line depreciation per period |
| `DDB` | `DDB(cost, salvage, life, period, [factor])` | Double-declining balance depreciation |
| `EFFECT` | `EFFECT(nominal_rate, npery)` | Effective annual interest rate |
| `NOMINAL` | `NOMINAL(effect_rate, npery)` | Nominal annual interest rate |

**Convention:** `pv`/`pmt` follow the Excel cash-flow sign convention (money paid out is negative). `type` = 0 means end-of-period payments (default), 1 means beginning-of-period.

**Examples:**
```
=PMT(5%/12, 360, 200000)         → monthly payment on a 30yr £200k mortgage at 5%
=PV(8%/12, 60, -500)             → present value of 60 monthly payments of £500
=FV(6%/12, 120, -200, -5000)     → future value after 10yrs of £200/mo + £5k lump sum
=NPV(10%, -10000, 3000, 4000, 5000, 6000)  → NPV of a project at 10% discount rate
=IRR({-10000, 3000, 4000, 5000}) → IRR of a cashflow series
=IPMT(5%/12, 1, 360, 200000)     → interest paid in month 1 of the mortgage above
=SLN(50000, 5000, 5)             → £9000/yr straight-line depreciation
=EFFECT(5%, 12)                  → effective annual rate for 5% nominal compounded monthly
```

---

## Configuration

ASAT reads a config file on startup at `~/.config/asat/config.toml` (respects `$XDG_CONFIG_HOME`).

Press `c` on the welcome screen to open it in your `$EDITOR`. The file is created automatically with all options commented on first open.

```toml
# Select a built-in theme by name — no hex values needed
theme_name = "nord"

# Display
default_col_width    = 10
min_col_width        = 3
max_col_width        = 60
scroll_padding       = 3
show_line_numbers    = false
relative_line_numbers = false
highlight_cursor_row = false
highlight_cursor_col = false
show_formula_bar     = true
show_tab_bar         = true
show_status_bar      = true
status_timeout       = 3      # seconds before status message fades (0 = never)

# Editing
undo_limit           = 1000
autosave_interval    = 0      # seconds between auto-saves (0 = disabled)
backup_on_save       = false  # write a .bak before overwriting
confirm_delete       = false  # ask before dd / :dr / :dc
wrap_navigation      = false  # wrap cursor at sheet edges

# Number formatting
number_precision     = 6      # max decimal places for unformatted numbers
date_format          = "YYYY-MM-DD"

# Files
default_format       = "csv"  # fallback format when saving without extension
csv_delimiter        = ","
remember_recent      = 20     # files shown on welcome screen

# Custom colors — only used when theme_name = "custom" or ""
# [theme]
# cursor_bg = "#268BD2"
# cell_bg   = "#002B36"
# ...
```

Built-in themes: `solarized-dark`, `solarized-light`, `nord`, `dracula`, `gruvbox-dark`, `gruvbox-light`, `tokyo-night`, `catppuccin-mocha`, `catppuccin-latte`, `one-dark`, `monokai`, `rose-pine`, `everforest-dark`, `kanagawa`, `cyberpunk`, `amber-terminal`, `ice`, `github-dark`

Browse and apply themes interactively with `:theme`.

---

## Plugin System

ASAT supports Python plugins via PyO3. Build with the feature enabled:

```bash
cargo build --release --features asat-plugins/python
```

Place your script at `~/.config/asat/init.py`. It is loaded on startup and can be reloaded live with `:plugin reload`.

```python
import asat

# React to events
@asat.on("cell_change")
def on_change(sheet, row, col, old, new):
    if isinstance(new, (int, float)) and new > 1_000_000:
        asat.notify("Very large value!")

@asat.on("mode_change")
def on_mode(mode):
    if mode == "INSERT":
        asat.notify("Editing...")

# Register custom formula functions (callable as =DOUBLE(A1))
@asat.function("DOUBLE")
def double(x):
    return x * 2

# Read/write cells by address
@asat.on("open")
def on_open(path):
    asat.write("A1", "Loaded: " + str(path))
    val = asat.read("B2")
```

**Available events:** `open`, `pre_save`, `post_save`, `cell_change`, `mode_change`, `sheet_change`

**Plugin API:** `asat.notify(msg)`, `asat.command(cmd)`, `asat.read("A1")`, `asat.write("A1", val)`, `asat.get_cell(row, col)`, `asat.set_cell(row, col, val)`

**Commands:** `:plugin list` — show loaded handlers and functions; `:plugin reload` — hot-reload `init.py`

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
  asat-plugins/   — Plugin manager: PyO3 Python backend (opt-in with --features asat-plugins/python)
  asat-config/    — Config struct, config.toml parsing, ThemeConfig
  asat/           — Binary: main loop, AppAction dispatch, ex-command handler
```

---

## Roadmap

- [x] MVP — CSV, navigation, insert, undo, `:w` / `:q`
- [x] Full Vim feel — `dd`/`yy`/`p`, marks, registers, visual mode, macros
- [x] Multi-sheet + XLSX / ODS (read & write)
- [x] Formula engine — 30+ functions, live evaluation, F-REF picker
- [x] Styles + formatting — bold, italic, colour, alignment, number formats
- [x] Sort & find/replace — `:sort`, `:s/pat/repl/g`
- [x] Plugin system — PyO3 backend, `init.py`, `@asat.on`, `@asat.function`, live reload with `:plugin reload` (build with `--features asat-plugins/python`)
- [x] Polish — filter rows, freeze panes, fill down/right, goto, transpose, dedup, named ranges, cell notes, conditional formatting, thousands separator, formula tab-completion, time-based autosave, ODS formula round-trip, XLOOKUP/AVERAGEIF/RANK/PERCENTILE/CHOOSE and more

---

## License

GNU General Public License v3.0 — see [LICENSE](LICENSE) for details.

Any fork or modified version distributed publicly must remain open source, credit the original authors, and use the same license.
