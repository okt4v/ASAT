/// Generates testing/demo.asat — a multi-sheet SPY stock analysis workbook
/// showcasing ASAT features: formulas, styles, conditional formatting, number
/// formats, cross-sheet refs, freeze rows, column widths, cell notes, merges.
///
/// Run from the workspace root:
///   cargo run --example gen_demo

use asat_core::*;
use asat_io::FileDriver;
use std::path::Path;

// ── Colour palette ────────────────────────────────────────────────────────────
const HDR_BG: Color = Color::rgb(15, 32, 68);      // dark navy
const HDR_FG: Color = Color::rgb(220, 230, 255);   // pale blue-white
const ACCENT: Color = Color::rgb(41, 98, 184);     // mid blue
const ACCENT2: Color = Color::rgb(22, 138, 101);   // teal-green
const DIM_BG: Color = Color::rgb(20, 24, 38);      // very dark blue-grey
const UP_FG: Color = Color::rgb(80, 220, 140);     // bright green
const DOWN_FG: Color = Color::rgb(220, 80, 80);    // bright red
const YELLOW: Color = Color::rgb(255, 200, 50);    // gold

// ── Helper: make a header style ───────────────────────────────────────────────
fn hdr(align: Alignment) -> CellStyle {
    CellStyle {
        bold: true,
        underline: true,
        bg: Some(HDR_BG),
        fg: Some(HDR_FG),
        align,
        ..Default::default()
    }
}

fn label_style() -> CellStyle {
    CellStyle {
        bold: true,
        fg: Some(HDR_FG),
        bg: Some(DIM_BG),
        align: Alignment::Left,
        ..Default::default()
    }
}

// ── Excel date serial (days since 1899-12-30) from unix days ──────────────────
fn date_serial(unix_days: i64) -> f64 {
    (unix_days + 25569) as f64
}

// ── SPY Jan 2025 price data ───────────────────────────────────────────────────
// (open, high, low, close, volume)
const SPY_DATA: &[(f64, f64, f64, f64, f64, &str)] = &[
    (584.12, 591.88, 582.49, 589.12, 72_300_000.0, "2025-01-02"),
    (589.20, 592.45, 587.36, 591.44, 58_100_000.0, "2025-01-03"),
    (591.80, 596.70, 590.30, 594.88, 72_400_000.0, "2025-01-06"),
    (595.20, 597.40, 590.80, 592.41, 68_200_000.0, "2025-01-07"),
    (591.50, 593.20, 582.30, 583.89, 89_000_000.0, "2025-01-08"),
    (582.00, 585.30, 578.40, 583.17, 78_500_000.0, "2025-01-09"),
    (582.80, 583.50, 575.20, 576.98, 95_100_000.0, "2025-01-10"),
    (577.50, 586.40, 576.80, 584.64, 82_300_000.0, "2025-01-13"),
    (585.20, 590.10, 583.70, 588.71, 71_200_000.0, "2025-01-14"),
    (589.30, 597.50, 588.10, 595.56, 88_400_000.0, "2025-01-15"),
    (595.80, 600.40, 594.20, 598.65, 76_100_000.0, "2025-01-16"),
    (598.90, 601.20, 595.60, 599.42, 69_500_000.0, "2025-01-17"),
    (601.20, 606.50, 599.80, 604.88, 74_300_000.0, "2025-01-21"),
    (605.10, 607.30, 601.40, 603.25, 66_800_000.0, "2025-01-22"),
    (603.50, 608.20, 602.10, 606.77, 70_200_000.0, "2025-01-23"),
    (607.20, 609.50, 604.80, 607.14, 63_400_000.0, "2025-01-24"),
    (607.00, 610.20, 604.30, 608.32, 71_600_000.0, "2025-01-27"),
    (608.80, 613.10, 607.50, 611.63, 68_900_000.0, "2025-01-28"),
    (612.00, 615.40, 608.90, 613.21, 72_500_000.0, "2025-01-29"),
    (613.50, 616.80, 610.20, 614.87, 75_000_000.0, "2025-01-30"),
];

// unix days for each trading date (Jan 2-30 2025, skipping weekends/holiday)
const UNIX_DAYS: &[i64] = &[
    20090, 20091, 20094, 20095, 20096, 20097, 20098,
    20101, 20102, 20103, 20104, 20105,
    20109, 20110, 20111, 20112,
    20115, 20116, 20117, 20118,
];

// ── Sheet 1: "Price Data" ─────────────────────────────────────────────────────
fn build_price_sheet(sheet: &mut Sheet) {
    sheet.freeze_rows = 2; // freeze title + header rows

    // Column widths: Date(12), Open(10), High(10), Low(10), Close(10), Volume(14)
    for (col, w) in [(0u32, 14u16), (1, 10), (2, 10), (3, 10), (4, 10), (5, 15)] {
        sheet.col_meta.insert(col, ColMeta { width: Some(w), ..Default::default() });
    }

    // Row 0: merged title
    sheet.set_cell(0, 0, Cell::with_style(
        CellValue::Text("  SPY — S&P 500 ETF  |  Daily Price History  |  Jan 2025".into()),
        CellStyle {
            bold: true,
            italic: true,
            bg: Some(ACCENT),
            fg: Some(Color::rgb(255, 255, 255)),
            align: Alignment::Center,
            ..Default::default()
        },
    ));
    sheet.add_merge(0, 0, 0, 5);

    // Row 1: column headers
    let hdrs = ["Date", "Open", "High", "Low", "Close", "Volume"];
    let aligns = [Alignment::Center, Alignment::Right, Alignment::Right,
                  Alignment::Right, Alignment::Right, Alignment::Right];
    for (col, (label, align)) in hdrs.iter().zip(aligns).enumerate() {
        sheet.set_cell(1, col as u32, Cell::with_style(
            CellValue::Text((*label).into()),
            hdr(align),
        ));
    }

    // Notes on headers
    sheet.notes.insert((1, 4), "Close price — used for all return calculations".into());
    sheet.notes.insert((1, 5), "Shares traded — formatted as #,##0".into());

    // Data rows 2-21
    let currency_style = |up: bool| CellStyle {
        fg: Some(if up { UP_FG } else { DOWN_FG }),
        align: Alignment::Right,
        format: Some(NumberFormat::Currency("$".into())),
        ..Default::default()
    };

    for (i, (&unix_day, row)) in UNIX_DAYS.iter().zip(SPY_DATA).enumerate() {
        let r = (i + 2) as u32;
        let (open, high, low, close, vol, _date_str) = *row;
        let up = close >= open;

        // Date column (A) — stored as Excel serial, formatted as date
        sheet.set_cell(r, 0, Cell::with_style(
            CellValue::Number(date_serial(unix_day)),
            CellStyle {
                fg: Some(Color::rgb(160, 180, 220)),
                align: Alignment::Center,
                format: Some(NumberFormat::Date(String::new())),
                ..Default::default()
            },
        ));

        // Open (B)
        sheet.set_cell(r, 1, Cell::with_style(
            CellValue::Number(open),
            CellStyle {
                fg: Some(Color::rgb(180, 195, 225)),
                align: Alignment::Right,
                format: Some(NumberFormat::Currency("$".into())),
                ..Default::default()
            },
        ));

        // High (C) — always green-tinted
        sheet.set_cell(r, 2, Cell::with_style(
            CellValue::Number(high),
            CellStyle {
                fg: Some(Color::rgb(100, 200, 120)),
                align: Alignment::Right,
                format: Some(NumberFormat::Currency("$".into())),
                ..Default::default()
            },
        ));

        // Low (D) — always red-tinted
        sheet.set_cell(r, 3, Cell::with_style(
            CellValue::Number(low),
            CellStyle {
                fg: Some(Color::rgb(220, 100, 100)),
                align: Alignment::Right,
                format: Some(NumberFormat::Currency("$".into())),
                ..Default::default()
            },
        ));

        // Close (E) — green if up, red if down; bold
        sheet.set_cell(r, 4, Cell::with_style(
            CellValue::Number(close),
            CellStyle {
                bold: true,
                ..currency_style(up)
            },
        ));

        // Volume (F) — thousands format
        sheet.set_cell(r, 5, Cell::with_style(
            CellValue::Number(vol),
            CellStyle {
                fg: Some(Color::rgb(160, 170, 200)),
                align: Alignment::Right,
                format: Some(NumberFormat::Thousands),
                ..Default::default()
            },
        ));
    }

    // Note on highest close
    sheet.notes.insert((16, 4), "Period high close: $604.88 on Jan 21".into());
    // Note on lowest close
    sheet.notes.insert((8, 4), "Period low close: $576.98 on Jan 10".into());
}

// ── Sheet 2: "Returns" ────────────────────────────────────────────────────────
fn build_returns_sheet(sheet: &mut Sheet) {
    sheet.freeze_rows = 2;

    for (col, w) in [(0u32, 14u16), (1, 11), (2, 11), (3, 11), (4, 11), (5, 13)] {
        sheet.col_meta.insert(col, ColMeta { width: Some(w), ..Default::default() });
    }

    // Row 0: title
    sheet.set_cell(0, 0, Cell::with_style(
        CellValue::Text("  SPY — Daily Returns & Moving Averages  |  Jan 2025".into()),
        CellStyle {
            bold: true,
            italic: true,
            bg: Some(ACCENT2),
            fg: Some(Color::rgb(255, 255, 255)),
            align: Alignment::Center,
            ..Default::default()
        },
    ));
    sheet.add_merge(0, 0, 0, 5);

    // Row 1: headers
    let hdrs = ["Date", "Close", "Daily Return", "5-Day SMA", "10-Day SMA", "Signal"];
    let aligns = [Alignment::Center, Alignment::Right, Alignment::Right,
                  Alignment::Right, Alignment::Right, Alignment::Center];
    for (col, (label, align)) in hdrs.iter().zip(aligns).enumerate() {
        sheet.set_cell(1, col as u32, Cell::with_style(
            CellValue::Text((*label).into()),
            hdr(align),
        ));
    }

    // Notes on headers
    sheet.notes.insert((1, 2), "Daily Return = (Close_n - Close_{n-1}) / Close_{n-1}".into());
    sheet.notes.insert((1, 3), "Simple moving average of last 5 closing prices".into());
    sheet.notes.insert((1, 4), "Simple moving average of last 10 closing prices".into());

    // Data rows 2-21 (row 2 = first day, no return yet)
    for (i, &_unix_day) in UNIX_DAYS.iter().enumerate() {
        let r = (i + 2) as u32;           // spreadsheet row (0-indexed)
        let pd_row = r;                    // same row number in Price Data sheet

        // Date — reference back to Price Data sheet
        sheet.set_cell(r, 0, Cell::with_style(
            CellValue::Formula(format!("'Price Data'!A{}", pd_row + 1)),
            CellStyle {
                fg: Some(Color::rgb(160, 180, 220)),
                align: Alignment::Center,
                format: Some(NumberFormat::Date(String::new())),
                ..Default::default()
            },
        ));

        // Close — reference Price Data E column
        sheet.set_cell(r, 1, Cell::with_style(
            CellValue::Formula(format!("'Price Data'!E{}", pd_row + 1)),
            CellStyle {
                fg: Some(Color::rgb(200, 215, 255)),
                align: Alignment::Right,
                format: Some(NumberFormat::Currency("$".into())),
                ..Default::default()
            },
        ));

        // Daily Return — only from row 3 onwards
        if i > 0 {
            sheet.set_cell(r, 2, Cell::with_style(
                CellValue::Formula(format!("(B{r1}-B{r0})/B{r0}", r1 = r + 1, r0 = r)),
                CellStyle {
                    align: Alignment::Right,
                    format: Some(NumberFormat::Percentage(2)),
                    ..Default::default()
                },
            ));
        } else {
            sheet.set_cell(r, 2, Cell::with_style(
                CellValue::Text("—".into()),
                CellStyle {
                    fg: Some(Color::rgb(100, 110, 140)),
                    align: Alignment::Center,
                    ..Default::default()
                },
            ));
        }

        // 5-Day SMA — needs at least 5 data points (rows 2-6, i.e. i >= 4)
        if i >= 4 {
            let start = r - 3; // B(r-3) through B(r+1)
            sheet.set_cell(r, 3, Cell::with_style(
                CellValue::Formula(format!("AVERAGE(B{}:B{})", start + 1, r + 1)),
                CellStyle {
                    fg: Some(YELLOW),
                    align: Alignment::Right,
                    format: Some(NumberFormat::Currency("$".into())),
                    ..Default::default()
                },
            ));
        } else {
            sheet.set_cell(r, 3, Cell::with_style(
                CellValue::Text("—".into()),
                CellStyle { fg: Some(Color::rgb(100, 110, 140)), align: Alignment::Center, ..Default::default() },
            ));
        }

        // 10-Day SMA — needs at least 10 data points (i >= 9)
        if i >= 9 {
            let start = r - 8;
            sheet.set_cell(r, 4, Cell::with_style(
                CellValue::Formula(format!("AVERAGE(B{}:B{})", start + 1, r + 1)),
                CellStyle {
                    fg: Some(Color::rgb(255, 160, 80)),
                    align: Alignment::Right,
                    format: Some(NumberFormat::Currency("$".into())),
                    ..Default::default()
                },
            ));
        } else {
            sheet.set_cell(r, 4, Cell::with_style(
                CellValue::Text("—".into()),
                CellStyle { fg: Some(Color::rgb(100, 110, 140)), align: Alignment::Center, ..Default::default() },
            ));
        }

        // Signal: ABOVE / BELOW / N/A based on 5-day SMA
        if i >= 4 {
            sheet.set_cell(r, 5, Cell::with_style(
                CellValue::Formula(format!("IF(B{r1}>D{r1},\"▲ ABOVE\",\"▼ BELOW\")", r1 = r + 1)),
                CellStyle {
                    bold: true,
                    align: Alignment::Center,
                    ..Default::default()
                },
            ));
        } else {
            sheet.set_cell(r, 5, Cell::with_style(
                CellValue::Text("N/A".into()),
                CellStyle { fg: Some(Color::rgb(100, 110, 140)), align: Alignment::Center, ..Default::default() },
            ));
        }
    }

    // Conditional formatting on Daily Return column (C, col 2) — rows 3-21 (0-indexed)
    // Green for positive returns
    sheet.conditional_formats.push(ConditionalFormat {
        row_start: 3,
        col_start: 2,
        row_end: 21,
        col_end: 2,
        condition: CfCondition::Gt(0.0),
        bg: None,
        fg: Some("#50dc8c".into()),
    });
    // Red for negative returns
    sheet.conditional_formats.push(ConditionalFormat {
        row_start: 3,
        col_start: 2,
        row_end: 21,
        col_end: 2,
        condition: CfCondition::Lt(0.0),
        bg: None,
        fg: Some("#dc5050".into()),
    });
}

// ── Sheet 3: "Summary" ────────────────────────────────────────────────────────
fn build_summary_sheet(sheet: &mut Sheet) {
    for (col, w) in [(0u32, 22u16), (1, 16), (2, 16), (3, 16)] {
        sheet.col_meta.insert(col, ColMeta { width: Some(w), ..Default::default() });
    }

    // Title row
    sheet.set_cell(0, 0, Cell::with_style(
        CellValue::Text("  SPY January 2025 — Performance Summary".into()),
        CellStyle {
            bold: true,
            italic: true,
            bg: Some(Color::rgb(80, 30, 100)),
            fg: Some(Color::rgb(240, 200, 255)),
            align: Alignment::Center,
            ..Default::default()
        },
    ));
    sheet.add_merge(0, 0, 0, 3);

    let mut row = 1u32;

    let kv = |sheet: &mut Sheet, r: u32, label: &str, formula: &str, fmt: Option<NumberFormat>| {
        sheet.set_cell(r, 0, Cell::with_style(
            CellValue::Text(label.into()),
            label_style(),
        ));
        sheet.set_cell(r, 1, Cell::with_style(
            CellValue::Formula(formula.into()),
            CellStyle {
                bg: Some(DIM_BG),
                fg: Some(Color::rgb(200, 220, 255)),
                bold: true,
                align: Alignment::Right,
                format: fmt,
                ..Default::default()
            },
        ));
    };

    // ── Section header ──
    let sec_hdr = |sheet: &mut Sheet, r: u32, text: &str| {
        sheet.set_cell(r, 0, Cell::with_style(
            CellValue::Text(format!("  {}", text)),
            CellStyle {
                bold: true,
                italic: true,
                underline_full: true,
                bg: Some(Color::rgb(30, 20, 55)),
                fg: Some(Color::rgb(180, 150, 230)),
                align: Alignment::Left,
                ..Default::default()
            },
        ));
        sheet.add_merge(r, 0, r, 3);
    };

    sec_hdr(sheet, row, "PERIOD"); row += 1;
    kv(sheet, row, "  Start Date", "'Price Data'!A3", Some(NumberFormat::Date(String::new()))); row += 1;
    kv(sheet, row, "  End Date",   "'Price Data'!A22", Some(NumberFormat::Date(String::new()))); row += 1;
    kv(sheet, row, "  Trading Days", "COUNT('Price Data'!E3:E22)", None); row += 1;

    row += 1; // spacer
    sec_hdr(sheet, row, "PRICE PERFORMANCE"); row += 1;
    kv(sheet, row, "  Start Price (Jan 2)", "'Price Data'!E3",  Some(NumberFormat::Currency("$".into()))); row += 1;
    kv(sheet, row, "  End Price (Jan 30)",  "'Price Data'!E22", Some(NumberFormat::Currency("$".into()))); row += 1;
    kv(sheet, row, "  Total Return",
        "('Price Data'!E22-'Price Data'!E3)/'Price Data'!E3",
        Some(NumberFormat::Percentage(2))); row += 1;
    kv(sheet, row, "  Period High (Close)", "MAX('Price Data'!E3:E22)", Some(NumberFormat::Currency("$".into()))); row += 1;
    kv(sheet, row, "  Period Low (Close)",  "MIN('Price Data'!E3:E22)", Some(NumberFormat::Currency("$".into()))); row += 1;
    kv(sheet, row, "  Period Range",
        "MAX('Price Data'!E3:E22)-MIN('Price Data'!E3:E22)",
        Some(NumberFormat::Currency("$".into()))); row += 1;

    row += 1; // spacer
    sec_hdr(sheet, row, "RETURN STATISTICS"); row += 1;
    kv(sheet, row, "  Avg Daily Return",  "AVERAGE(Returns!C4:C22)", Some(NumberFormat::Percentage(3))); row += 1;
    kv(sheet, row, "  Best Day",          "MAX(Returns!C4:C22)",     Some(NumberFormat::Percentage(2))); row += 1;
    kv(sheet, row, "  Worst Day",         "MIN(Returns!C4:C22)",     Some(NumberFormat::Percentage(2))); row += 1;
    kv(sheet, row, "  Daily Volatility",  "STDEV(Returns!C4:C22)",   Some(NumberFormat::Percentage(2))); row += 1;
    kv(sheet, row, "  Ann. Volatility (≈√252)",
        "STDEV(Returns!C4:C22)*SQRT(252)",
        Some(NumberFormat::Percentage(1))); row += 1;

    row += 1; // spacer
    sec_hdr(sheet, row, "VOLUME & BREADTH"); row += 1;
    kv(sheet, row, "  Avg Daily Volume",  "AVERAGE('Price Data'!F3:F22)", Some(NumberFormat::Thousands)); row += 1;
    kv(sheet, row, "  Total Volume",      "SUM('Price Data'!F3:F22)",     Some(NumberFormat::Thousands)); row += 1;
    kv(sheet, row, "  Days Advancing",    "COUNTIF(Returns!C4:C22,\">0\")", None); row += 1;
    kv(sheet, row, "  Days Declining",    "COUNTIF(Returns!C4:C22,\"<0\")", None);

    // Note on summary
    sheet.notes.insert((6, 1), "First close price on Jan 2 2025".into());
    sheet.notes.insert((7, 1), "Last close price on Jan 30 2025".into());
    sheet.notes.insert((15, 1), "STDEV of daily returns * √252 trading days/year".into());
}

// ── Main ──────────────────────────────────────────────────────────────────────
fn main() {
    let mut wb = Workbook::new();

    // Rename default sheet and build it
    wb.sheets[0].name = "Price Data".into();
    build_price_sheet(&mut wb.sheets[0]);

    // Returns sheet
    wb.add_sheet("Returns");
    let returns_idx = wb.sheets.len() - 1;
    build_returns_sheet(&mut wb.sheets[returns_idx]);

    // Summary sheet
    wb.add_sheet("Summary");
    let summary_idx = wb.sheets.len() - 1;
    build_summary_sheet(&mut wb.sheets[summary_idx]);

    wb.active_sheet = 0;

    let out = Path::new("testing/demo.asat");
    asat_io::asat_driver::AsatDriver
        .write(&wb, out)
        .expect("failed to write demo.asat");

    println!("Written {}", out.display());
    println!("  Sheet 1: Price Data   — {} cells", wb.sheets[0].cells.len());
    println!("  Sheet 2: Returns      — {} cells", wb.sheets[1].cells.len());
    println!("  Sheet 3: Summary      — {} cells", wb.sheets[2].cells.len());
}
