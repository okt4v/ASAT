pub mod lexer;
pub mod parser;
pub mod evaluator;
pub mod functions;

pub use evaluator::{EvalContext, Evaluator};
pub use parser::{Expr, ParseError};
pub use lexer::{Token, LexError};

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

fn evaluate_inner(
    formula: &str,
    workbook: &Workbook,
    sheet_idx: usize,
    row: u32,
    col: u32,
) -> Result<CellValue, FormulaError> {
    let tokens = lexer::lex(formula)?;
    let expr = parser::parse(&tokens)?;
    let ctx = EvalContext { workbook, sheet_idx, row, col };
    let evaluator = Evaluator::new();
    Ok(evaluator.eval(&expr, &ctx))
}
