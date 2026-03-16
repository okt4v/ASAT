pub mod evaluator;
pub mod functions;
pub mod lexer;
pub mod parser;

pub use evaluator::{EvalContext, Evaluator};
pub use lexer::{LexError, Token};
pub use parser::{Expr, ParseError};

use asat_core::{CellValue, Workbook};
use thiserror::Error;

#[derive(Debug, Error)]
pub enum FormulaError {
    #[error("lex error: {0}")]
    Lex(#[from] LexError),
    #[error("parse error: {0}")]
    Parse(#[from] ParseError),
    #[error("eval error: {0}")]
    Eval(String),
}

/// Top-level: evaluate a formula string in the context of a workbook/cell.
/// `formula` should NOT include the leading '='.
pub fn evaluate(
    formula: &str,
    workbook: &Workbook,
    sheet_idx: usize,
    row: u32,
    col: u32,
) -> CellValue {
    match evaluate_inner(formula, workbook, sheet_idx, row, col) {
        Ok(v) => v,
        Err(FormulaError::Lex(_)) => CellValue::Error(asat_core::CellError::Name),
        Err(FormulaError::Parse(_)) => CellValue::Error(asat_core::CellError::Value),
        Err(FormulaError::Eval(_)) => CellValue::Error(asat_core::CellError::Value),
    }
}

/// Collect all cell coordinates referenced by a formula string (without leading '=').
/// Expands ranges to individual cells. Returns a Vec of (row, col) pairs.
pub fn collect_cell_refs(formula: &str) -> Vec<(u32, u32)> {
    let tokens = match lexer::lex(formula) {
        Ok(t) => t,
        Err(_) => return Vec::new(),
    };
    let expr = match parser::parse(&tokens) {
        Ok(e) => e,
        Err(_) => return Vec::new(),
    };
    let mut refs = Vec::new();
    collect_refs_from_expr(&expr, &mut refs);
    refs
}

fn collect_refs_from_expr(expr: &parser::Expr, out: &mut Vec<(u32, u32)>) {
    match expr {
        parser::Expr::CellRef { row, col, .. } => {
            out.push((*row, *col));
        }
        parser::Expr::RangeRef {
            row1,
            col1,
            row2,
            col2,
            ..
        } => {
            let r_min = (*row1).min(*row2);
            let r_max = (*row1).max(*row2);
            let c_min = (*col1).min(*col2);
            let c_max = (*col1).max(*col2);
            // Clamp large ranges to avoid blowing up memory
            let r_max = r_max.min(r_min + 999);
            let c_max = c_max.min(c_min + 99);
            for r in r_min..=r_max {
                for c in c_min..=c_max {
                    out.push((r, c));
                }
            }
        }
        parser::Expr::UnaryMinus(e) | parser::Expr::UnaryPlus(e) => {
            collect_refs_from_expr(e, out);
        }
        parser::Expr::Add(a, b)
        | parser::Expr::Sub(a, b)
        | parser::Expr::Mul(a, b)
        | parser::Expr::Div(a, b)
        | parser::Expr::Pow(a, b)
        | parser::Expr::Concat(a, b)
        | parser::Expr::Eq(a, b)
        | parser::Expr::Neq(a, b)
        | parser::Expr::Lt(a, b)
        | parser::Expr::Lte(a, b)
        | parser::Expr::Gt(a, b)
        | parser::Expr::Gte(a, b) => {
            collect_refs_from_expr(a, out);
            collect_refs_from_expr(b, out);
        }
        parser::Expr::Call { args, .. } => {
            for arg in args {
                collect_refs_from_expr(arg, out);
            }
        }
        _ => {}
    }
}

fn evaluate_inner(
    formula: &str,
    workbook: &Workbook,
    sheet_idx: usize,
    row: u32,
    col: u32,
) -> Result<CellValue, FormulaError> {
    let tokens = lexer::lex(formula)?;
    let expr = parser::parse(&tokens)?;
    let ctx = EvalContext {
        workbook,
        sheet_idx,
        row,
        col,
    };
    let evaluator = Evaluator::new();
    Ok(evaluator.eval(&expr, &ctx))
}
