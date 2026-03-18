use asat_core::{Cell, CellValue, Workbook};
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
