use crate::functions;
use crate::parser::Expr;
use asat_core::{CellError, CellValue, Workbook};

pub struct EvalContext<'a> {
    pub workbook: &'a Workbook,
    pub sheet_idx: usize,
    pub row: u32,
    pub col: u32,
}

#[derive(Default)]
pub struct Evaluator;

impl Evaluator {
    pub fn new() -> Self {
        Evaluator
    }

    pub fn eval(&self, expr: &Expr, ctx: &EvalContext<'_>) -> CellValue {
        match expr {
            Expr::Number(n) => CellValue::Number(*n),
            Expr::Text(s) => CellValue::Text(s.clone()),
            Expr::Boolean(b) => CellValue::Boolean(*b),

            Expr::CellRef {
                sheet, col, row, ..
            } => {
                let sheet_idx = self.resolve_sheet(sheet, ctx);
                if let Some(s) = ctx.workbook.sheet(sheet_idx) {
                    s.get_value(*row, *col).clone()
                } else {
                    CellValue::Error(CellError::Ref)
                }
            }

            Expr::RangeRef { .. } => {
                // Ranges are handled by functions; bare range is an error
                CellValue::Error(CellError::Value)
            }

            Expr::UnaryMinus(e) => match self.eval(e, ctx) {
                CellValue::Number(n) => CellValue::Number(-n),
                CellValue::Error(e) => CellValue::Error(e),
                _ => CellValue::Error(CellError::Value),
            },
            Expr::UnaryPlus(e) => self.eval(e, ctx),

            Expr::Add(a, b) => self.binop_num(a, b, ctx, |x, y| x + y),
            Expr::Sub(a, b) => self.binop_num(a, b, ctx, |x, y| x - y),
            Expr::Mul(a, b) => self.binop_num(a, b, ctx, |x, y| x * y),
            Expr::Div(a, b) => {
                let lv = self.eval(a, ctx);
                let rv = self.eval(b, ctx);
                match (to_number(&lv), to_number(&rv)) {
                    (_, Some(0.0)) => CellValue::Error(CellError::Div0),
                    (Some(x), Some(y)) => CellValue::Number(x / y),
                    _ => propagate_error(&lv, &rv),
                }
            }
            Expr::Pow(a, b) => self.binop_num(a, b, ctx, |x, y| x.powf(y)),

            Expr::Concat(a, b) => {
                let lv = self.eval(a, ctx);
                let rv = self.eval(b, ctx);
                match (&lv, &rv) {
                    (CellValue::Error(e), _) => CellValue::Error(e.clone()),
                    (_, CellValue::Error(e)) => CellValue::Error(e.clone()),
                    _ => CellValue::Text(format!("{}{}", to_text(&lv), to_text(&rv))),
                }
            }

            Expr::Eq(a, b) => self.compare(a, b, ctx, |o| o == std::cmp::Ordering::Equal),
            Expr::Neq(a, b) => self.compare(a, b, ctx, |o| o != std::cmp::Ordering::Equal),
            Expr::Lt(a, b) => self.compare(a, b, ctx, |o| o == std::cmp::Ordering::Less),
            Expr::Lte(a, b) => self.compare(a, b, ctx, |o| o != std::cmp::Ordering::Greater),
            Expr::Gt(a, b) => self.compare(a, b, ctx, |o| o == std::cmp::Ordering::Greater),
            Expr::Gte(a, b) => self.compare(a, b, ctx, |o| o != std::cmp::Ordering::Less),

            Expr::Call { name, args } => {
                // Check if this is a named range used as a scalar (zero args)
                if args.is_empty() {
                    let upper = name.to_ascii_uppercase();
                    if let Some(range) = ctx.workbook.named_ranges.get(&upper) {
                        // Single-cell named range → return that cell's value
                        if range.row_start == range.row_end && range.col_start == range.col_end {
                            if let Some(s) = ctx.workbook.sheet(range.sheet) {
                                return s.get_value(range.row_start, range.col_start).clone();
                            }
                        }
                        // Multi-cell: return #VALUE — use in a function like SUM(rangeName)
                        return CellValue::Error(CellError::Value);
                    }
                }
                functions::call(name, args, ctx, self)
            }
        }
    }

    /// Expand a range expression into individual CellValues
    pub fn expand_range(&self, expr: &Expr, ctx: &EvalContext<'_>) -> Vec<CellValue> {
        // Handle named ranges: a Call with no args whose name matches a named range
        if let Expr::Call { name, args } = expr {
            if args.is_empty() {
                let upper = name.to_ascii_uppercase();
                if let Some(range) = ctx.workbook.named_ranges.get(&upper) {
                    if let Some(s) = ctx.workbook.sheet(range.sheet) {
                        let mut vals = Vec::new();
                        for r in range.row_start..=range.row_end {
                            for c in range.col_start..=range.col_end {
                                vals.push(s.get_value(r, c).clone());
                            }
                        }
                        return vals;
                    }
                }
            }
        }
        match expr {
            Expr::RangeRef {
                sheet,
                col1,
                row1,
                col2,
                row2,
                ..
            } => {
                let sheet_idx = self.resolve_sheet(sheet, ctx);
                if let Some(s) = ctx.workbook.sheet(sheet_idx) {
                    let mut vals = Vec::new();
                    let (r_min, r_max) = ((*row1).min(*row2), (*row1).max(*row2));
                    let (c_min, c_max) = ((*col1).min(*col2), (*col1).max(*col2));
                    for r in r_min..=r_max {
                        for c in c_min..=c_max {
                            vals.push(s.get_value(r, c).clone());
                        }
                    }
                    vals
                } else {
                    vec![CellValue::Error(CellError::Ref)]
                }
            }
            // Single cell ref in range context
            Expr::CellRef {
                sheet, col, row, ..
            } => {
                let sheet_idx = self.resolve_sheet(sheet, ctx);
                if let Some(s) = ctx.workbook.sheet(sheet_idx) {
                    vec![s.get_value(*row, *col).clone()]
                } else {
                    vec![CellValue::Error(CellError::Ref)]
                }
            }
            // Literal value
            other => vec![self.eval(other, ctx)],
        }
    }

    fn resolve_sheet(&self, sheet: &Option<String>, ctx: &EvalContext<'_>) -> usize {
        if let Some(name) = sheet {
            ctx.workbook
                .sheets
                .iter()
                .position(|s| s.name.eq_ignore_ascii_case(name))
                .unwrap_or(ctx.sheet_idx)
        } else {
            ctx.sheet_idx
        }
    }

    fn binop_num(
        &self,
        a: &Expr,
        b: &Expr,
        ctx: &EvalContext<'_>,
        f: impl Fn(f64, f64) -> f64,
    ) -> CellValue {
        let lv = self.eval(a, ctx);
        let rv = self.eval(b, ctx);
        match (to_number(&lv), to_number(&rv)) {
            (Some(x), Some(y)) => CellValue::Number(f(x, y)),
            _ => propagate_error(&lv, &rv),
        }
    }

    fn compare(
        &self,
        a: &Expr,
        b: &Expr,
        ctx: &EvalContext<'_>,
        pred: impl Fn(std::cmp::Ordering) -> bool,
    ) -> CellValue {
        let lv = self.eval(a, ctx);
        let rv = self.eval(b, ctx);
        if let (CellValue::Error(e), _) = (&lv, &rv) {
            return CellValue::Error(e.clone());
        }
        if let (_, CellValue::Error(e)) = (&lv, &rv) {
            return CellValue::Error(e.clone());
        }
        let ord = compare_values(&lv, &rv);
        CellValue::Boolean(pred(ord))
    }
}

// ── Value coercions ───────────────────────────────────────────────────────────

pub fn to_number(v: &CellValue) -> Option<f64> {
    match v {
        CellValue::Number(n) => Some(*n),
        CellValue::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
        CellValue::Text(s) => s.parse::<f64>().ok(),
        CellValue::Empty => Some(0.0),
        _ => None,
    }
}

pub fn to_bool(v: &CellValue) -> Option<bool> {
    match v {
        CellValue::Boolean(b) => Some(*b),
        CellValue::Number(n) => Some(*n != 0.0),
        CellValue::Text(s) => match s.to_uppercase().as_str() {
            "TRUE" => Some(true),
            "FALSE" => Some(false),
            _ => None,
        },
        CellValue::Empty => Some(false),
        _ => None,
    }
}

pub fn to_text(v: &CellValue) -> String {
    v.display()
}

fn compare_values(a: &CellValue, b: &CellValue) -> std::cmp::Ordering {
    use CellValue::*;
    match (a, b) {
        (Number(x), Number(y)) => x.partial_cmp(y).unwrap_or(std::cmp::Ordering::Equal),
        (Text(x), Text(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
        (Boolean(x), Boolean(y)) => x.cmp(y),
        (Empty, Empty) => std::cmp::Ordering::Equal,
        // Type ordering: Empty < Number < Text < Boolean (Excel-like)
        (Empty, _) => std::cmp::Ordering::Less,
        (_, Empty) => std::cmp::Ordering::Greater,
        (Number(_), Text(_)) | (Number(_), Boolean(_)) => std::cmp::Ordering::Less,
        (Text(_), Boolean(_)) => std::cmp::Ordering::Less,
        _ => std::cmp::Ordering::Greater,
    }
}

fn propagate_error(a: &CellValue, b: &CellValue) -> CellValue {
    match a {
        CellValue::Error(e) => CellValue::Error(e.clone()),
        _ => match b {
            CellValue::Error(e) => CellValue::Error(e.clone()),
            _ => CellValue::Error(CellError::Value),
        },
    }
}
