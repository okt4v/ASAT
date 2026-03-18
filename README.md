# ASAT — A Spreadsheet And Terminal

> Terminal spreadsheet editor for Vim users. Modal editing (Normal/Insert/Visual/Command), 50+ live formulas, multi-sheet workbooks, CSV · XLSX · ODS support, full undo stack, system clipboard, named ranges, filter/freeze panes, cell notes, live conditional formatting, auto-fill series, macros, marks, `.` repeat, Python plugins, searchable help screen, and circular reference detection. Written in Rust with ratatui.

[![License: GPL-3.0](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Built with Rust](https://img.shields.io/badge/built%20with-Rust-orange.svg)](https://www.rust-lang.org)

**Website:** <https://okt4v.github.io/ASAT/>

---

## Features

- **Modal editing** — Normal, Insert, Visual (char/line/block), Command, Search, and Macro Recording modes, exactly like Vim
- **Live formula engine** — 50+ built-in functions across math, text, logic, lookup, statistical, finance, and date. Includes volatile functions (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`) that recalculate every frame. `=NOW()` auto-displays as `YYYY-MM-DD HH:MM`; use `:fmt datetime` for other date-serial cells. Formulas re-evaluate after every edit (lazy dirty-cell tracking)
- **Circular reference detection** — cells that form dependency cycles display `#CIRC!` instead of crashing or hanging
- **Named ranges** — `:name SALES A1:C10` defines a named range usable in formulas as `=SUM(SALES)`
- **Multi-sheet workbooks** — tab bar, `:tabnew`, `:tabclose`, `gt` / `gT` to switch sheets
- **File format support** — read and write CSV, TSV, XLSX, and ODS (OpenDocument Spreadsheet); native `.asat` format with bincode + zstd compression; ODS formula round-trip with cached computed values
- **System clipboard** — yank (`yy`, `yc`, visual `y`) copies values and styles as TSV; `Ctrl+V` in Insert mode pastes from system clipboard; delete (`dd`, `x`, visual `d`) also yanks to register + clipboard before clearing
- **Formula-aware paste** — pasting a formula adjusts relative cell references (`A1` becomes `B2` when pasted one row down and one column right); absolute references (`$A$1`) stay fixed; mixed references (`$A1`, `A$1`) lock only the `$`-prefixed axis
- **Full undo stack** — 1000-deep undo/redo covering cell edits, row/column operations, pastes, and style changes; undo/redo repositions the cursor to the affected cell
- **Sort & find/replace** — `:sort [COL] [!]` sorts rows by any column (e.g. `:sort B!` = column B descending); `:s/pat/repl/g` does regex find & replace across cells; both undoable
- **Filter rows** — `:filter <col> <op> <val>` hides non-matching rows (supports `=`, `!=`, `>`, `<`, `>=`, `<=`, `~`); `:filter off` restores all rows
- **Freeze panes** — frozen rows and columns render as sticky headers with a visual separator; set via `:freeze rows N` / `:freeze cols N`
- **Fill down / right** — `Ctrl+D` / `Ctrl+R` in Visual mode copies anchor cell across selection; also `:filldown` / `:fillright` ex-commands
- **Auto-fill series** — `Ctrl+F` / `Ctrl+E` in Visual mode detects and extends arithmetic sequences, weekday names, and month names across the selection
- **Live conditional formatting** — `:cf <range> <cond> <val> bg=#hex` applies colour rules that re-evaluate every frame; supports `>`, `<`, `>=`, `<=`, `=`, `!=`, `contains`, `blank`, `error`; `:cf clear` removes all rules
- **Formula color distinction** — formula cells render in muted blue to distinguish them from literal data at a glance
- **Live formula preview** — while typing a formula, the current evaluated result appears in the formula bar as `→ value`
- **Cell reference highlighting** — when the cursor rests on a formula cell, all referenced cells are highlighted in the grid
- **Go-to definition** — `gd` in Normal mode jumps to the first cell referenced in the current cell's formula
- **Visual mode ex-commands** — press `:` from Visual/V-Row/V-Col mode to enter a command that applies to the entire selection (e.g. `:bold`, `:fg #ff0000`, `:sort`)
- **Repeat last change** — `.` in Normal mode replays the last insert or delete operation (like Vim)
- **Goto cell** — `g<letter>` jumps to a column; `:goto B15` jumps to any cell address
- **Transpose** — `:transpose` swaps rows and columns in the visual selection
- **Remove duplicates** — `:dedup` removes duplicate rows by the current cursor column
- **Cell notes** — `:note <text>` attaches a comment to the current cell; cells with notes show a `▸` corner marker; `:note` with no argument shows the current note; `:note!` clears it
- **Thousands separator** — `:fmt thousands` formats numbers with comma separators (`#,##0`); `:fmt t2` adds two decimal places
- **Formula tab-completion** — press `Tab` while typing a formula (`=SU…`) to cycle through matching function names
- **Time-based autosave** — configurable autosave interval in seconds (edit `autosave_interval` in `config.toml`; 0 = disabled)
- **Macros** — record key sequences to named registers (`qa` … `qz`), replay with `@a`, chain with `{N}@a`
- **Marks** — set named positions (`ma`), jump back (`'a`), and swap with `''`
- **Cell merging** — merge a visual selection into one spanning cell with `M` in Visual mode or `:merge`; unmerge with `U` or `:unmerge`; line wrap flows text into covered rows below the anchor
- **Line wrap** — toggle per-cell line wrap with `gw` (Normal mode) or `:wrap`; text reflows into merged rows below for vertical merges, and row height auto-expands for single cells
- **Cell styling** — bold, italic, underline, strikethrough, foreground/background colour, alignment, and number formats
- **Themes** — built-in theme picker (`:theme`) with multiple colour presets, saved to config
- **Formula reference picker** — press `Ctrl+R` inside a formula to navigate the grid and insert cell/range references interactively
- **Searchable help screen** — `:help` / `:h` opens a full-screen overlay with Keybindings and Formulas tabs; type to filter, `Tab` to switch tabs, `j`/`k` to scroll, `q` to close
- **Plugin system** — extend ASAT with Python via PyO3 (enabled by default); hook into cell changes, mode transitions, and file events; register custom formula functions from `~/.config/asat/init.py`; manage plugins with `:plugins`

---

## Installation

### Pre-built binaries (GitHub Releases)

Download a binary for your platform from the [v0.1.22 release](https://github.com/okt4v/ASAT/releases/tag/v0.1.22):

| Platform | Link |
|----------|------|
| Linux x86_64 (glibc) | [asat-x86_64-unknown-linux-gnu.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.22/asat-x86_64-unknown-linux-gnu.tar.gz) |
| Linux x86_64 (musl)  | [asat-x86_64-unknown-linux-musl.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.22/asat-x86_64-unknown-linux-musl.tar.gz) |
| Linux aarch64        | [asat-aarch64-unknown-linux-gnu.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.22/asat-aarch64-unknown-linux-gnu.tar.gz) |
| macOS arm64          | [asat-aarch64-apple-darwin.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.22/asat-aarch64-apple-darwin.tar.gz) |
| macOS x86_64         | [asat-x86_64-apple-darwin.tar.gz](https://github.com/okt4v/ASAT/releases/download/v0.1.22/asat-x86_64-apple-darwin.tar.gz) |
| Windows x86_64       | [asat-x86_64-pc-windows-msvc.zip](https://github.com/okt4v/ASAT/releases/download/v0.1.22/asat-x86_64-pc-windows-msvc.zip) |

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

### Debian / Ubuntu (apt)

```bash
curl -LO https://github.com/okt4v/ASAT/releases/download/v0.1.22/asat_0.1.21-1_amd64.deb
sudo apt install ./asat_0.1.21-1_amd64.deb
```

`apt install ./file.deb` resolves dependencies automatically and registers the package so `apt remove asat` works as expected.

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

> **Note:** The Python plugin engine is enabled by default. To build without it (smaller binary, no Python dependency): `cargo build --release --no-default-features`

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
| `ci"` / `ci(` / `ci[` / `ci{` | Change inner text object — clear content inside delimiters and edit |
| `.` | Repeat last change (last insert, delete, or destructive key) |
| `r` | Enter replace mode |
| `x` / `Del` / `D` | Delete cell content (supports count, e.g. `3x`) |
| `~` | Toggle case of text cell (supports count) |
| `J` | Join cell below into current cell (space-separated), clear below |
| `gd` | Jump to the first cell referenced in the current formula |
| `Ctrl+a` | Increment number / cycle date forward (day → month → weekday) |
| `Ctrl+x` | Decrement number / cycle date backward |
| `gw` | Toggle line wrap on current cell |
| `U` | Unmerge cell under cursor |
| `o` / `O` | Insert row below / above and enter insert mode |
| `u` | Undo |
| `Ctrl+r` | Redo |

### Normal Mode — Yank & Paste

| Key | Action |
|-----|--------|
| `yy` / `yr` | Yank current row → register + system clipboard |
| `yc` | Yank current cell → register + system clipboard |
| `yC` | Yank entire column → register + system clipboard |
| `yj` / `yk` | Yank row below / above cursor |
| `yS` | Copy current cell's style to style clipboard |
| `p` / `P` | Paste after / before cursor (supports count) |
| `pS` | Paste style clipboard onto current cell |

### Normal Mode — Rows & Columns

| Key | Action |
|-----|--------|
| `dd` | Delete current row |
| `dc` | Clear cell content (alias for `x`) |
| `dC` | Delete current column |
| `dj` / `dk` | Delete row below / above cursor |
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
| `M` | Merge selection into one spanning cell |
| `d` / `x` / `Del` | Delete selection |
| `c` / `s` | Clear selection and enter insert mode |
| `y` | Yank selection → register + system clipboard (TSV) |
| `S` | Insert `=SUM(range)` below the selection |
| `>` / `<` | Widen / narrow all columns in selection |
| `Ctrl+d` | Fill down — copy anchor row to all selected rows |
| `Ctrl+r` | Fill right — copy anchor column to all selected columns |
| `Ctrl+f` | Auto-fill series down — extends arithmetic, weekday, or month patterns |
| `Ctrl+e` | Auto-fill series right |
| `:` | Enter Command mode with the selection range pre-loaded (e.g. `:bold`, `:fg #hex`) |

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

Enter command mode with `:`. From Visual mode, `:` pre-loads the current selection range so style and formatting commands apply to all selected cells.

Tab-completion works in command mode — press `Tab` to cycle through matching commands.

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
| `:sort [COL] [!]` | Sort by column letter (e.g. `:sort A`, `:sort B!` = descending); defaults to cursor column (undoable) |
| `:s/pat/repl/[g][i]` | Find & replace in text cells — `g` = all occurrences, `i` = case-insensitive (undoable) |
| `:goto <addr>` / `:go <addr>` | Jump to cell address (e.g. `:goto B15`) |
| `:name <NAME> <range>` | Define a named range (e.g. `:name SALES A1:C10`) |
| `:filter <col> <op> <val>` | Hide rows where column does not match (ops: `=` `!=` `>` `<` `>=` `<=` `~`) |
| `:filter off` | Unhide all filtered rows |
| `:transpose` / `:tp` | Transpose the visual selection (swap rows and columns) |
| `:dedup` | Remove duplicate rows by the current cursor column |
| `:note [text]` | Set a note on current cell; `:note` shows it; `:note!` clears it |
| `:cf <range> <cond> <val> bg=#hex [fg=#hex]` | Add live conditional format rule (e.g. `:cf A1:C10 > 100 bg=#ff0000`); conditions: `>` `<` `>=` `<=` `=` `!=` `contains` `blank` `error` |
| `:cf clear` | Remove all conditional format rules from the active sheet |
| `:cf list` | Show number of active conditional format rules |
| `:filldown` / `:fd` | Fill the cursor cell value down to the selection end |
| `:fillright` / `:fr` | Fill the cursor cell value right to the selection end |
| `:fmt thousands` / `:fmt t2` | Thousands-separator number format (`#,##0` or `#,##0.00`) |
| `:freeze rows <N>` | Freeze top N rows as sticky header |
| `:freeze cols <N>` | Freeze left N columns as sticky header |
| `:freeze off` | Clear all frozen panes |
| `:merge` | Merge visual selection (or current cell) into one spanning cell |
| `:unmerge` | Unmerge the merged region under the cursor |
| `:wrap` / `:ww` | Toggle line wrap on current cell or selection |
| `:help` / `:h` | Open full-screen searchable help (Keybindings + Formulas tabs) |
| `:home` | Return to the welcome / home screen |
| `:plugins` | Open plugin manager TUI |

---

## Formula Engine

Start any cell with `=` to write a formula. Formulas re-evaluate automatically after every edit.

```
=SUM(A1:A10)
=IF(B2>100, "Over budget", "OK")
=AVERAGE(C1:C20) * 1.1
=CONCATENATE(A1, " ", B1)
=NOW()              → current date and time (recalculates every frame)
=RAND()             → random float between 0 and 1
```

**Supported functions:**

| Category | Functions |
|----------|-----------|
| Math | `SUM`, `AVERAGE`, `MIN`, `MAX`, `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `FLOOR`, `CEILING`, `MOD`, `POWER`, `SQRT`, `LN`, `LOG`, `LOG10`, `EXP`, `INT`, `TRUNC`, `SIGN` |
| Text | `LEN`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `UPPER`, `LOWER`, `PROPER`, `CONCATENATE`, `TEXT`, `VALUE`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `REPT` |
| Logic | `IF`, `AND`, `OR`, `NOT`, `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISERROR`, `IFERROR`, `ISLOGICAL` |
| Lookup | `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `CHOOSE` |
| Date | `NOW`, `TODAY`, `DATE`, `YEAR`, `MONTH`, `DAY` |
| Random | `RAND`, `RANDBETWEEN` |
| Statistical | `COUNT`, `COUNTA`, `SUMIF`, `COUNTIF`, `AVERAGEIF`, `MAXIFS`, `MINIFS`, `SUMPRODUCT`, `MEDIAN`, `STDEV`, `VAR`, `LARGE`, `SMALL`, `RANK`, `PERCENTILE`, `QUARTILE` |
| Finance | `PV`, `FV`, `PMT`, `NPER`, `RATE`, `NPV`, `IRR`, `MIRR`, `IPMT`, `PPMT`, `SLN`, `DDB`, `EFFECT`, `NOMINAL`, `CUMIPMT`, `CUMPRINC` |
| Constants | `TRUE`, `FALSE`, `PI()` |

**Volatile functions** (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`) recalculate on every frame, not just on cell edits.

**Circular references:** If a formula creates a dependency cycle (e.g. `A1=B1+1` and `B1=A1+1`), all cells in the cycle display `#CIRC!`. Break the cycle by editing one of the cells.

**Reference syntax:**

| Syntax | Meaning |
|--------|---------|
| `A1` | Relative cell reference (adjusts on paste) |
| `$A$1` | Absolute cell reference (fixed on paste) |
| `$A1` / `A$1` | Mixed reference (lock column or row only) |
| `A1:B10` | Range |
| `Sheet2!C4` | Cross-sheet reference |

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

## Help Screen

Press `:help` (or `:h`) to open a full-screen searchable help overlay.

- **Keybindings tab** — all Normal, Insert, and Visual mode keybinds organized by category
- **Formulas tab** — all 50+ built-in functions with one-line descriptions
- **Search** — type any text to filter entries across both tabs instantly
- **Navigation** — `j`/`k` to scroll, `Tab` to switch tabs, `q` or `Esc` to close

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

ASAT includes a Python plugin engine via PyO3, **enabled by default**. Place your script at `~/.config/asat/init.py` — it is loaded on startup and can be hot-reloaded with `:plugin reload`.

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

**Commands:** `:plugins` — open the plugin manager TUI (shows engine status and registered custom functions); `:plugin reload` — hot-reload `init.py`; `:plugin list` — list loaded handlers

**To build without Python** (smaller binary, no Python dependency):

```bash
cargo build --release --no-default-features
```

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
  asat-formula/   — Lexer, parser (AST), evaluator, 50+ built-in functions
  asat-io/        — CSV, XLSX, ODS, .asat drivers (calamine read / rust_xlsxwriter write)
  asat-tui/       — ratatui widgets: grid, formula bar, status bar, tab bar, command line
  asat-input/     — Modal state machine, InputState, AppAction enum
  asat-commands/  — Command trait, UndoStack, SetCell, InsertRow/Col, DeleteRow/Col
  asat-plugins/   — Plugin manager: PyO3 Python backend (enabled by default)
  asat-config/    — Config struct, config.toml parsing, ThemeConfig
  asat/           — Binary: main loop, AppAction dispatch, ex-command handler
```

---

## License

GNU General Public License v3.0 — see [LICENSE](LICENSE) for details.

Any fork or modified version distributed publicly must remain open source, credit the original authors, and use the same license.
