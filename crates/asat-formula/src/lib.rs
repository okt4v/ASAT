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
/// Includes cross-sheet references (sheet qualifier is ignored — only coords returned).
/// Use `collect_same_sheet_refs` for dependency/cycle analysis.
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
    collect_refs_from_expr(&expr, &mut refs, false);
    refs
}

/// Collect cell coordinates referenced by a formula, skipping cross-sheet references.
/// Used for same-sheet dependency/circular-reference analysis — cross-sheet refs
/// cannot create cycles on the current sheet and must be excluded.
pub fn collect_same_sheet_refs(formula: &str) -> Vec<(u32, u32)> {
    let tokens = match lexer::lex(formula) {
        Ok(t) => t,
        Err(_) => return Vec::new(),
    };
    let expr = match parser::parse(&tokens) {
        Ok(e) => e,
        Err(_) => return Vec::new(),
    };
    let mut refs = Vec::new();
    collect_refs_from_expr(&expr, &mut refs, true);
    refs
}

fn collect_refs_from_expr(expr: &parser::Expr, out: &mut Vec<(u32, u32)>, same_sheet_only: bool) {
    match expr {
        parser::Expr::CellRef {
            row, col, sheet, ..
        } => {
            if !same_sheet_only || sheet.is_none() {
                out.push((*row, *col));
            }
        }
        parser::Expr::RangeRef {
            row1,
            col1,
            row2,
            col2,
            sheet,
            ..
        } => {
            if same_sheet_only && sheet.is_some() {
                return;
            }
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
            collect_refs_from_expr(e, out, same_sheet_only);
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
            collect_refs_from_expr(a, out, same_sheet_only);
            collect_refs_from_expr(b, out, same_sheet_only);
        }
        parser::Expr::Call { args, .. } => {
            for arg in args {
                collect_refs_from_expr(arg, out, same_sheet_only);
            }
        }
        _ => {}
    }
}

/// Adjust cell references in a formula when pasting.
/// `d_row` / `d_col` are the signed offsets from the source to the destination cell.
/// Absolute references (`$A` or `$1`) are left unchanged; relative references are shifted.
/// Returns the adjusted formula string (without leading `=`).
pub fn adjust_formula_refs(formula: &str, d_row: i64, d_col: i64) -> String {
    let tokens = match lexer::lex(formula) {
        Ok(t) => t,
        Err(_) => return formula.to_string(),
    };
    let mut out = String::new();
    for tok in &tokens {
        match tok {
            Token::CellRef {
                sheet,
                col,
                row,
                abs_col,
                abs_row,
            } => {
                if let Some(s) = sheet {
                    out.push_str(s);
                    out.push('!');
                }
                let new_col = if *abs_col {
                    *col
                } else {
                    (*col as i64 + d_col).max(0) as u32
                };
                let new_row = if *abs_row {
                    *row
                } else {
                    (*row as i64 + d_row).max(0) as u32
                };
                if *abs_col {
                    out.push('$');
                }
                out.push_str(&asat_core::col_to_letter(new_col));
                if *abs_row {
                    out.push('$');
                }
                out.push_str(&(new_row + 1).to_string());
            }
            Token::RangeRef {
                sheet,
                col1,
                row1,
                abs_col1,
                abs_row1,
                col2,
                row2,
                abs_col2,
                abs_row2,
            } => {
                if let Some(s) = sheet {
                    out.push_str(s);
                    out.push('!');
                }
                let nc1 = if *abs_col1 {
                    *col1
                } else {
                    (*col1 as i64 + d_col).max(0) as u32
                };
                let nr1 = if *abs_row1 {
                    *row1
                } else {
                    (*row1 as i64 + d_row).max(0) as u32
                };
                let nc2 = if *abs_col2 {
                    *col2
                } else {
                    (*col2 as i64 + d_col).max(0) as u32
                };
                let nr2 = if *abs_row2 {
                    *row2
                } else {
                    (*row2 as i64 + d_row).max(0) as u32
                };
                if *abs_col1 {
                    out.push('$');
                }
                out.push_str(&asat_core::col_to_letter(nc1));
                if *abs_row1 {
                    out.push('$');
                }
                out.push_str(&(nr1 + 1).to_string());
                out.push(':');
                if *abs_col2 {
                    out.push('$');
                }
                out.push_str(&asat_core::col_to_letter(nc2));
                if *abs_row2 {
                    out.push('$');
                }
                out.push_str(&(nr2 + 1).to_string());
            }
            Token::Number(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    out.push_str(&format!("{}", *n as i64));
                } else {
                    out.push_str(&format!("{}", n));
                }
            }
            Token::Text(s) => {
                out.push('"');
                out.push_str(&s.replace('"', "\"\""));
                out.push('"');
            }
            Token::Boolean(b) => out.push_str(if *b { "TRUE" } else { "FALSE" }),
            Token::Ident(s) => out.push_str(s),
            Token::Plus => out.push('+'),
            Token::Minus => out.push('-'),
            Token::Star => out.push('*'),
            Token::Slash => out.push('/'),
            Token::Caret => out.push('^'),
            Token::Eq => out.push('='),
            Token::Neq => out.push_str("<>"),
            Token::Lt => out.push('<'),
            Token::Lte => out.push_str("<="),
            Token::Gt => out.push('>'),
            Token::Gte => out.push_str(">="),
            Token::Ampersand => out.push('&'),
            Token::LParen => out.push('('),
            Token::RParen => out.push(')'),
            Token::Comma => out.push(','),
            Token::Colon => out.push(':'),
            Token::Semicolon => out.push(';'),
            Token::Eof => {}
        }
    }
    out
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
