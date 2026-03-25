use asat_core::{Cell, CellError, CellValue, Workbook};
use asat_formula::evaluate;

fn wb_with_data() -> Workbook {
    let mut wb = Workbook::new();
    let s = wb.active_mut();
    // A1:A5 = 1,2,3,4,5
    for i in 0..5u32 {
        s.set_cell(i, 0, Cell::new(CellValue::Number((i + 1) as f64)));
    }
    // B1 = "hello"
    s.set_cell(0, 1, Cell::new(CellValue::Text("hello".into())));
    // C1 = TRUE
    s.set_cell(0, 2, Cell::new(CellValue::Boolean(true)));
    wb
}

/// Workbook with two sheets for cross-sheet reference tests.
/// Sheet1: A1=10, A2=20
/// Sheet2: A1=100, A2=200
fn wb_two_sheets() -> Workbook {
    let mut wb = Workbook::new();
    wb.active_mut().set_cell(0, 0, Cell::new(CellValue::Number(10.0)));
    wb.active_mut().set_cell(1, 0, Cell::new(CellValue::Number(20.0)));
    wb.add_sheet("Sheet2");
    wb.sheets[1].set_cell(0, 0, Cell::new(CellValue::Number(100.0)));
    wb.sheets[1].set_cell(1, 0, Cell::new(CellValue::Number(200.0)));
    wb
}

fn eval(formula: &str, wb: &Workbook) -> CellValue {
    evaluate(formula, wb, 0, 5, 5) // eval from outside the data range
}

#[test]
fn test_sum() {
    let wb = wb_with_data();
    assert_eq!(eval("SUM(A1:A5)", &wb), CellValue::Number(15.0));
}

#[test]
fn test_average() {
    let wb = wb_with_data();
    assert_eq!(eval("AVERAGE(A1:A5)", &wb), CellValue::Number(3.0));
}

#[test]
fn test_min_max() {
    let wb = wb_with_data();
    assert_eq!(eval("MIN(A1:A5)", &wb), CellValue::Number(1.0));
    assert_eq!(eval("MAX(A1:A5)", &wb), CellValue::Number(5.0));
}

#[test]
fn test_count() {
    let wb = wb_with_data();
    assert_eq!(eval("COUNT(A1:A5)", &wb), CellValue::Number(5.0));
}

#[test]
fn test_arithmetic() {
    let wb = wb_with_data();
    assert_eq!(eval("2+3", &wb), CellValue::Number(5.0));
    assert_eq!(eval("10-4", &wb), CellValue::Number(6.0));
    assert_eq!(eval("3*4", &wb), CellValue::Number(12.0));
    assert_eq!(eval("10/4", &wb), CellValue::Number(2.5));
    assert_eq!(eval("2^10", &wb), CellValue::Number(1024.0));
}

#[test]
fn test_if() {
    let wb = wb_with_data();
    assert_eq!(
        eval("IF(1>0,\"yes\",\"no\")", &wb),
        CellValue::Text("yes".into())
    );
    assert_eq!(
        eval("IF(1<0,\"yes\",\"no\")", &wb),
        CellValue::Text("no".into())
    );
}

#[test]
fn test_string_functions() {
    let wb = wb_with_data();
    assert_eq!(eval("LEN(\"hello\")", &wb), CellValue::Number(5.0));
    assert_eq!(
        eval("UPPER(\"hello\")", &wb),
        CellValue::Text("HELLO".into())
    );
    assert_eq!(
        eval("LEFT(\"hello\",3)", &wb),
        CellValue::Text("hel".into())
    );
    assert_eq!(
        eval("RIGHT(\"hello\",3)", &wb),
        CellValue::Text("llo".into())
    );
    assert_eq!(
        eval("MID(\"hello\",2,3)", &wb),
        CellValue::Text("ell".into())
    );
}

#[test]
fn test_concat() {
    let wb = wb_with_data();
    assert_eq!(
        eval("\"foo\"&\"bar\"", &wb),
        CellValue::Text("foobar".into())
    );
    assert_eq!(
        eval("CONCATENATE(\"a\",\"b\",\"c\")", &wb),
        CellValue::Text("abc".into())
    );
}

#[test]
fn test_div_by_zero() {
    let wb = wb_with_data();
    assert_eq!(
        eval("1/0", &wb),
        CellValue::Error(asat_core::CellError::Div0)
    );
}

#[test]
fn test_cell_ref() {
    let wb = wb_with_data();
    assert_eq!(eval("A1+A2", &wb), CellValue::Number(3.0));
}

#[test]
fn test_abs_round() {
    let wb = wb_with_data();
    assert_eq!(eval("ABS(-5)", &wb), CellValue::Number(5.0));
    #[allow(clippy::approx_constant)]
    let expected = 3.14;
    assert_eq!(eval("ROUND(3.14159,2)", &wb), CellValue::Number(expected));
}

// ── SUMIF / COUNTIF ──────────────────────────────────────────────────────────

/// Workbook for SUMIF/COUNTIF tests:
/// A1:A5 = "apple","banana","apple","cherry","apple"
/// B1:B5 = 10, 20, 30, 40, 50
fn wb_sumif() -> Workbook {
    let mut wb = Workbook::new();
    let s = wb.active_mut();
    let fruits = ["apple", "banana", "apple", "cherry", "apple"];
    let values = [10.0f64, 20.0, 30.0, 40.0, 50.0];
    for (i, (f, v)) in fruits.iter().zip(values.iter()).enumerate() {
        s.set_cell(i as u32, 0, Cell::new(CellValue::Text(f.to_string())));
        s.set_cell(i as u32, 1, Cell::new(CellValue::Number(*v)));
    }
    wb
}

#[test]
fn test_sumif_equal() {
    let wb = wb_sumif();
    // Sum B where A = "apple": 10+30+50 = 90
    assert_eq!(
        evaluate("SUMIF(A1:A5,\"apple\",B1:B5)", &wb, 0, 6, 0),
        CellValue::Number(90.0)
    );
}

#[test]
fn test_sumif_no_match() {
    let wb = wb_sumif();
    assert_eq!(
        evaluate("SUMIF(A1:A5,\"mango\",B1:B5)", &wb, 0, 6, 0),
        CellValue::Number(0.0)
    );
}

#[test]
fn test_countif_equal() {
    let wb = wb_sumif();
    // Count A where A = "apple": 3
    assert_eq!(
        evaluate("COUNTIF(A1:A5,\"apple\")", &wb, 0, 6, 0),
        CellValue::Number(3.0)
    );
}

#[test]
fn test_countif_numeric_comparison() {
    let wb = wb_with_data();
    // A1:A5 = 1,2,3,4,5 — count values > 3
    assert_eq!(
        evaluate("COUNTIF(A1:A5,\">3\")", &wb, 0, 6, 0),
        CellValue::Number(2.0)
    );
}

#[test]
fn test_sumif_numeric_gt() {
    let wb = wb_with_data();
    // A1:A5 = 1,2,3,4,5; sum where value > 3 → 4+5 = 9
    assert_eq!(
        evaluate("SUMIF(A1:A5,\">3\")", &wb, 0, 6, 0),
        CellValue::Number(9.0)
    );
}

// ── Error propagation ────────────────────────────────────────────────────────

#[test]
fn test_error_propagates_through_arithmetic() {
    let wb = wb_with_data();
    // 1/0 produces #DIV/0!, adding a number to it should still be an error
    assert_eq!(
        evaluate("(1/0)+1", &wb, 0, 6, 0),
        CellValue::Error(CellError::Div0)
    );
}

#[test]
fn test_error_propagates_through_sum() {
    let mut wb = Workbook::new();
    // A1 = #DIV/0! (via formula stored as error value)
    wb.active_mut()
        .set_cell(0, 0, Cell::new(CellValue::Error(CellError::Div0)));
    // SUM over a range containing an error cell should propagate the error
    assert_eq!(
        evaluate("A1+1", &wb, 0, 1, 0),
        CellValue::Error(CellError::Div0)
    );
}

#[test]
fn test_iferror_catches_div0() {
    let wb = wb_with_data();
    assert_eq!(
        evaluate("IFERROR(1/0,\"caught\")", &wb, 0, 6, 0),
        CellValue::Text("caught".into())
    );
}

#[test]
fn test_iferror_passes_through_valid() {
    let wb = wb_with_data();
    assert_eq!(
        evaluate("IFERROR(42,\"caught\")", &wb, 0, 6, 0),
        CellValue::Number(42.0)
    );
}

#[test]
fn test_division_by_zero_is_div0_error() {
    let wb = wb_with_data();
    assert_eq!(
        evaluate("1/0", &wb, 0, 6, 0),
        CellValue::Error(CellError::Div0)
    );
}

// ── Cross-sheet references ────────────────────────────────────────────────────

#[test]
fn test_cross_sheet_cell_ref() {
    let wb = wb_two_sheets();
    // Evaluate on sheet 0, reference Sheet2!A1 = 100
    assert_eq!(
        evaluate("Sheet2!A1", &wb, 0, 2, 0),
        CellValue::Number(100.0)
    );
}

#[test]
fn test_cross_sheet_arithmetic() {
    let wb = wb_two_sheets();
    // Sheet1!A1=10, Sheet2!A1=100 → sum = 110
    assert_eq!(
        evaluate("A1+Sheet2!A1", &wb, 0, 2, 0),
        CellValue::Number(110.0)
    );
}

#[test]
fn test_cross_sheet_range_sum() {
    let wb = wb_two_sheets();
    // SUM(Sheet2!A1:A2) = 100+200 = 300
    assert_eq!(
        evaluate("SUM(Sheet2!A1:A2)", &wb, 0, 2, 0),
        CellValue::Number(300.0)
    );
}
