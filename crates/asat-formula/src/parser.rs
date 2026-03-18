use crate::lexer::Token;
use thiserror::Error;

#[derive(Debug, Error, Clone)]
pub enum ParseError {
    #[error("unexpected token: {0:?}")]
    UnexpectedToken(Token),
    #[error("unexpected end of input")]
    UnexpectedEof,
    #[error("expected '{0}'")]
    Expected(&'static str),
}

/// Abstract Syntax Tree node
#[derive(Debug, Clone)]
pub enum Expr {
    Number(f64),
    Text(String),
    Boolean(bool),
    CellRef {
        sheet: Option<String>,
        col: u32,
        row: u32,
        abs_col: bool,
        abs_row: bool,
    },
    RangeRef {
        sheet: Option<String>,
        col1: u32,
        row1: u32,
        abs_col1: bool,
        abs_row1: bool,
        col2: u32,
        row2: u32,
        abs_col2: bool,
        abs_row2: bool,
    },

    UnaryMinus(Box<Expr>),
    UnaryPlus(Box<Expr>),

    // Binary arithmetic
    Add(Box<Expr>, Box<Expr>),
    Sub(Box<Expr>, Box<Expr>),
    Mul(Box<Expr>, Box<Expr>),
    Div(Box<Expr>, Box<Expr>),
    Pow(Box<Expr>, Box<Expr>),
    Concat(Box<Expr>, Box<Expr>), // &

    // Comparisons
    Eq(Box<Expr>, Box<Expr>),
    Neq(Box<Expr>, Box<Expr>),
    Lt(Box<Expr>, Box<Expr>),
    Lte(Box<Expr>, Box<Expr>),
    Gt(Box<Expr>, Box<Expr>),
    Gte(Box<Expr>, Box<Expr>),

    // Function call
    Call {
        name: String,
        args: Vec<Expr>,
    },
}

struct Parser<'a> {
    tokens: &'a [Token],
    pos: usize,
}

impl<'a> Parser<'a> {
    fn new(tokens: &'a [Token]) -> Self {
        Parser { tokens, pos: 0 }
    }

    fn peek(&self) -> &Token {
        self.tokens.get(self.pos).unwrap_or(&Token::Eof)
    }

    fn advance(&mut self) -> &Token {
        let t = self.tokens.get(self.pos).unwrap_or(&Token::Eof);
        if self.pos < self.tokens.len() {
            self.pos += 1;
        }
        t
    }

    #[allow(dead_code)]
    fn expect_ident(&mut self) -> Result<String, ParseError> {
        match self.advance().clone() {
            Token::Ident(s) => Ok(s),
            t => Err(ParseError::UnexpectedToken(t)),
        }
    }

    /// Parse expression (lowest precedence: comparison)
    fn parse_expr(&mut self) -> Result<Expr, ParseError> {
        self.parse_concat()
    }

    fn parse_concat(&mut self) -> Result<Expr, ParseError> {
        let mut left = self.parse_comparison()?;
        loop {
            if matches!(self.peek(), Token::Ampersand) {
                self.advance();
                let right = self.parse_comparison()?;
                left = Expr::Concat(Box::new(left), Box::new(right));
            } else {
                break;
            }
        }
        Ok(left)
    }

    fn parse_comparison(&mut self) -> Result<Expr, ParseError> {
        let mut left = self.parse_additive()?;
        loop {
            let op = match self.peek() {
                Token::Eq => Expr::Eq as fn(_, _) -> _,
                Token::Neq => Expr::Neq as fn(_, _) -> _,
                Token::Lt => Expr::Lt as fn(_, _) -> _,
                Token::Lte => Expr::Lte as fn(_, _) -> _,
                Token::Gt => Expr::Gt as fn(_, _) -> _,
                Token::Gte => Expr::Gte as fn(_, _) -> _,
                _ => break,
            };
            self.advance();
            let right = self.parse_additive()?;
            left = op(Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn parse_additive(&mut self) -> Result<Expr, ParseError> {
        let mut left = self.parse_multiplicative()?;
        loop {
            match self.peek() {
                Token::Plus => {
                    self.advance();
                    let r = self.parse_multiplicative()?;
                    left = Expr::Add(Box::new(left), Box::new(r));
                }
                Token::Minus => {
                    self.advance();
                    let r = self.parse_multiplicative()?;
                    left = Expr::Sub(Box::new(left), Box::new(r));
                }
                _ => break,
            }
        }
        Ok(left)
    }

    fn parse_multiplicative(&mut self) -> Result<Expr, ParseError> {
        let mut left = self.parse_power()?;
        loop {
            match self.peek() {
                Token::Star => {
                    self.advance();
                    let r = self.parse_power()?;
                    left = Expr::Mul(Box::new(left), Box::new(r));
                }
                Token::Slash => {
                    self.advance();
                    let r = self.parse_power()?;
                    left = Expr::Div(Box::new(left), Box::new(r));
                }
                _ => break,
            }
        }
        Ok(left)
    }

    fn parse_power(&mut self) -> Result<Expr, ParseError> {
        let base = self.parse_unary()?;
        if matches!(self.peek(), Token::Caret) {
            self.advance();
            let exp = self.parse_unary()?; // right-associative
            Ok(Expr::Pow(Box::new(base), Box::new(exp)))
        } else {
            Ok(base)
        }
    }

    fn parse_unary(&mut self) -> Result<Expr, ParseError> {
        match self.peek() {
            Token::Minus => {
                self.advance();
                Ok(Expr::UnaryMinus(Box::new(self.parse_unary()?)))
            }
            Token::Plus => {
                self.advance();
                Ok(Expr::UnaryPlus(Box::new(self.parse_unary()?)))
            }
            _ => self.parse_primary(),
        }
    }

    fn parse_primary(&mut self) -> Result<Expr, ParseError> {
        match self.peek().clone() {
            Token::Number(n) => {
                self.advance();
                Ok(Expr::Number(n))
            }
            Token::Text(s) => {
                self.advance();
                Ok(Expr::Text(s))
            }
            Token::Boolean(b) => {
                self.advance();
                Ok(Expr::Boolean(b))
            }

            Token::CellRef {
                sheet,
                col,
                row,
                abs_col,
                abs_row,
            } => {
                self.advance();
                Ok(Expr::CellRef {
                    sheet,
                    col,
                    row,
                    abs_col,
                    abs_row,
                })
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
                self.advance();
                Ok(Expr::RangeRef {
                    sheet,
                    col1,
                    row1,
                    abs_col1,
                    abs_row1,
                    col2,
                    row2,
                    abs_col2,
                    abs_row2,
                })
            }

            Token::Ident(name) => {
                self.advance();
                // Could be a function call
                if matches!(self.peek(), Token::LParen) {
                    self.advance(); // consume '('
                    let mut args = Vec::new();
                    if !matches!(self.peek(), Token::RParen) {
                        args.push(self.parse_expr()?);
                        while matches!(self.peek(), Token::Comma | Token::Semicolon) {
                            self.advance();
                            if matches!(self.peek(), Token::RParen) {
                                break;
                            }
                            args.push(self.parse_expr()?);
                        }
                    }
                    if !matches!(self.peek(), Token::RParen) {
                        return Err(ParseError::Expected(")"));
                    }
                    self.advance(); // consume ')'
                    Ok(Expr::Call {
                        name: name.to_uppercase(),
                        args,
                    })
                } else {
                    // Named range or unknown identifier
                    Err(ParseError::UnexpectedToken(Token::Ident(name)))
                }
            }

            Token::LParen => {
                self.advance();
                let expr = self.parse_expr()?;
                if !matches!(self.peek(), Token::RParen) {
                    return Err(ParseError::Expected(")"));
                }
                self.advance();
                Ok(expr)
            }

            Token::Eof => Err(ParseError::UnexpectedEof),
            t => Err(ParseError::UnexpectedToken(t)),
        }
    }
}

pub fn parse(tokens: &[Token]) -> Result<Expr, ParseError> {
    let mut p = Parser::new(tokens);
    let expr = p.parse_expr()?;
    // Make sure we consumed everything (except Eof)
    if !matches!(p.peek(), Token::Eof) {
        return Err(ParseError::UnexpectedToken(p.peek().clone()));
    }
    Ok(expr)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::lexer::lex;

    fn parse_str(s: &str) -> Result<Expr, ParseError> {
        let tokens = lex(s).unwrap();
        parse(&tokens)
    }

    #[test]
    fn test_parse_number() {
        assert!(matches!(parse_str("42").unwrap(), Expr::Number(42.0)));
    }

    #[test]
    fn test_parse_add() {
        assert!(matches!(parse_str("1+2").unwrap(), Expr::Add(_, _)));
    }

    #[test]
    fn test_parse_function() {
        let e = parse_str("SUM(A1:A3)").unwrap();
        if let Expr::Call { name, args } = e {
            assert_eq!(name, "SUM");
            assert_eq!(args.len(), 1);
        } else {
            panic!("expected Call");
        }
    }
}
