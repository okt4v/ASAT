use thiserror::Error;

#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    // Literals
    Number(f64),
    Text(String),
    Boolean(bool),

    // Cell refs
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

    // Identifiers (function names)
    Ident(String),

    // Operators
    Plus,
    Minus,
    Star,
    Slash,
    Caret, // ^ = power
    Eq,
    Neq,
    Lt,
    Lte,
    Gt,
    Gte,
    Ampersand, // string concat

    // Punctuation
    LParen,
    RParen,
    Comma,
    Colon,
    Semicolon,

    // Structural
    Eof,
}

#[derive(Debug, Error, Clone)]
pub enum LexError {
    #[error("unexpected character '{0}'")]
    UnexpectedChar(char),
    #[error("invalid cell reference '{0}'")]
    InvalidCellRef(String),
    #[error("unterminated string")]
    UnterminatedString,
}

pub fn lex(input: &str) -> Result<Vec<Token>, LexError> {
    let mut tokens = Vec::new();
    let chars: Vec<char> = input.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        let c = chars[i];

        // Skip whitespace
        if c.is_ascii_whitespace() {
            i += 1;
            continue;
        }

        // Numbers
        if c.is_ascii_digit() || (c == '.' && i + 1 < chars.len() && chars[i + 1].is_ascii_digit())
        {
            let start = i;
            while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                i += 1;
            }
            // Optional exponent
            if i < chars.len() && (chars[i] == 'e' || chars[i] == 'E') {
                i += 1;
                if i < chars.len() && (chars[i] == '+' || chars[i] == '-') {
                    i += 1;
                }
                while i < chars.len() && chars[i].is_ascii_digit() {
                    i += 1;
                }
            }
            let s: String = chars[start..i].iter().collect();
            let n: f64 = s.parse().map_err(|_| LexError::UnexpectedChar(c))?;
            tokens.push(Token::Number(n));
            continue;
        }

        // String literals (double-quoted)
        if c == '"' {
            i += 1;
            let mut s = String::new();
            loop {
                if i >= chars.len() {
                    return Err(LexError::UnterminatedString);
                }
                if chars[i] == '"' {
                    // Escaped quote: ""
                    if i + 1 < chars.len() && chars[i + 1] == '"' {
                        s.push('"');
                        i += 2;
                    } else {
                        i += 1;
                        break;
                    }
                } else {
                    s.push(chars[i]);
                    i += 1;
                }
            }
            tokens.push(Token::Text(s));
            continue;
        }

        // Sheet-qualified references or identifiers: Sheet1!A1 or just A1 or FUNCNAME
        if c.is_ascii_alphabetic() || c == '\'' {
            // Collect identifier
            let start = i;
            // Sheet names can be 'Sheet Name'! (quoted)
            let sheet_prefix: Option<String>;
            if c == '\'' {
                // Quoted sheet name
                i += 1;
                let mut sname = String::new();
                loop {
                    if i >= chars.len() {
                        return Err(LexError::UnterminatedString);
                    }
                    if chars[i] == '\'' {
                        i += 1;
                        break;
                    }
                    sname.push(chars[i]);
                    i += 1;
                }
                if i < chars.len() && chars[i] == '!' {
                    i += 1;
                    sheet_prefix = Some(sname);
                } else {
                    // Not a sheet prefix — push as text? This is unusual, treat as error
                    return Err(LexError::InvalidCellRef(sname));
                }
            } else {
                // Collect alphanumeric run
                while i < chars.len() && (chars[i].is_ascii_alphanumeric() || chars[i] == '_') {
                    i += 1;
                }
                let ident: String = chars[start..i].iter().collect();

                // Check if next char is '!' → sheet prefix
                if i < chars.len() && chars[i] == '!' {
                    i += 1;
                    sheet_prefix = Some(ident);
                } else {
                    // Check: is it TRUE/FALSE?
                    match ident.to_uppercase().as_str() {
                        "TRUE" => {
                            tokens.push(Token::Boolean(true));
                            continue;
                        }
                        "FALSE" => {
                            tokens.push(Token::Boolean(false));
                            continue;
                        }
                        _ => {}
                    }
                    // Try to parse as cell ref (e.g. A1, $B$3, AB12)
                    if let Some(tok) = try_parse_cell_ref(&ident, None) {
                        tokens.push(tok);
                        // Check for range colon
                        if i < chars.len() && chars[i] == ':' {
                            i += 1;
                            // Collect next ref
                            let _ref2_start = i;
                            let mut abs_col2 = false;
                            let mut abs_row2 = false;
                            if i < chars.len() && chars[i] == '$' {
                                abs_col2 = true;
                                i += 1;
                            }
                            let col2_start = i;
                            while i < chars.len() && chars[i].is_ascii_alphabetic() {
                                i += 1;
                            }
                            let col2_str: String = chars[col2_start..i].iter().collect();
                            if i < chars.len() && chars[i] == '$' {
                                abs_row2 = true;
                                i += 1;
                            }
                            let row2_start = i;
                            while i < chars.len() && chars[i].is_ascii_digit() {
                                i += 1;
                            }
                            let row2_str: String = chars[row2_start..i].iter().collect();
                            // Replace the last token with a range ref
                            if let Some(Token::CellRef {
                                sheet,
                                col,
                                row,
                                abs_col,
                                abs_row,
                            }) = tokens.pop()
                            {
                                if let (Some(col2), Ok(row2)) = (
                                    asat_core::letter_to_col(&col2_str),
                                    row2_str.parse::<u32>().map(|r| r.saturating_sub(1)),
                                ) {
                                    tokens.push(Token::RangeRef {
                                        sheet,
                                        col1: col,
                                        row1: row,
                                        abs_col1: abs_col,
                                        abs_row1: abs_row,
                                        col2,
                                        row2,
                                        abs_col2,
                                        abs_row2,
                                    });
                                }
                            }
                        }
                    } else {
                        // It's a function name or named range
                        tokens.push(Token::Ident(ident));
                    }
                    continue;
                }
            }

            // We have a sheet prefix — parse cell ref after it
            let sheet = sheet_prefix;
            let _cell_start = i;
            let mut abs_col = false;
            let mut abs_row = false;
            if i < chars.len() && chars[i] == '$' {
                abs_col = true;
                i += 1;
            }
            let col_start = i;
            while i < chars.len() && chars[i].is_ascii_alphabetic() {
                i += 1;
            }
            let col_str: String = chars[col_start..i].iter().collect();
            if i < chars.len() && chars[i] == '$' {
                abs_row = true;
                i += 1;
            }
            let row_start = i;
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
            let row_str: String = chars[row_start..i].iter().collect();

            let col = asat_core::letter_to_col(&col_str)
                .ok_or_else(|| LexError::InvalidCellRef(col_str.clone()))?;
            let row = row_str
                .parse::<u32>()
                .map(|r| r.saturating_sub(1))
                .map_err(|_| LexError::InvalidCellRef(row_str.clone()))?;

            let tok = Token::CellRef {
                sheet: sheet.clone(),
                col,
                row,
                abs_col,
                abs_row,
            };
            // Check for range
            if i < chars.len() && chars[i] == ':' {
                i += 1;
                // Skip optional sheet prefix (same sheet assumed)
                let mut abs_col2 = false;
                let mut abs_row2 = false;
                if i < chars.len() && chars[i] == '$' {
                    abs_col2 = true;
                    i += 1;
                }
                let col2_start = i;
                while i < chars.len() && chars[i].is_ascii_alphabetic() {
                    i += 1;
                }
                let col2_str: String = chars[col2_start..i].iter().collect();
                if i < chars.len() && chars[i] == '$' {
                    abs_row2 = true;
                    i += 1;
                }
                let row2_start = i;
                while i < chars.len() && chars[i].is_ascii_digit() {
                    i += 1;
                }
                let row2_str: String = chars[row2_start..i].iter().collect();
                let col2 = asat_core::letter_to_col(&col2_str)
                    .ok_or_else(|| LexError::InvalidCellRef(col2_str.clone()))?;
                let row2 = row2_str
                    .parse::<u32>()
                    .map(|r| r.saturating_sub(1))
                    .map_err(|_| LexError::InvalidCellRef(row2_str.clone()))?;
                tokens.push(Token::RangeRef {
                    sheet,
                    col1: col,
                    row1: row,
                    abs_col1: abs_col,
                    abs_row1: abs_row,
                    col2,
                    row2,
                    abs_col2,
                    abs_row2,
                });
            } else {
                tokens.push(tok);
            }
            continue;
        }

        // $ prefix for absolute refs: collect as part of cell ref
        if c == '$' {
            i += 1;
            let abs_col = true;
            let col_start = i;
            while i < chars.len() && chars[i].is_ascii_alphabetic() {
                i += 1;
            }
            let col_str: String = chars[col_start..i].iter().collect();
            let mut abs_row = false;
            if i < chars.len() && chars[i] == '$' {
                abs_row = true;
                i += 1;
            }
            let row_start = i;
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
            let row_str: String = chars[row_start..i].iter().collect();
            let col = asat_core::letter_to_col(&col_str)
                .ok_or_else(|| LexError::InvalidCellRef(col_str.clone()))?;
            let row = row_str
                .parse::<u32>()
                .map(|r| r.saturating_sub(1))
                .map_err(|_| LexError::InvalidCellRef(row_str.clone()))?;
            tokens.push(Token::CellRef {
                sheet: None,
                col,
                row,
                abs_col,
                abs_row,
            });
            continue;
        }

        // Operators and punctuation
        let tok = match c {
            '+' => Token::Plus,
            '-' => Token::Minus,
            '*' => Token::Star,
            '/' => Token::Slash,
            '^' => Token::Caret,
            '&' => Token::Ampersand,
            '(' => Token::LParen,
            ')' => Token::RParen,
            ',' => Token::Comma,
            ';' => Token::Semicolon,
            '<' => {
                if i + 1 < chars.len() {
                    match chars[i + 1] {
                        '=' => {
                            i += 1;
                            Token::Lte
                        }
                        '>' => {
                            i += 1;
                            Token::Neq
                        }
                        _ => Token::Lt,
                    }
                } else {
                    Token::Lt
                }
            }
            '>' => {
                if i + 1 < chars.len() && chars[i + 1] == '=' {
                    i += 1;
                    Token::Gte
                } else {
                    Token::Gt
                }
            }
            '=' => {
                if i + 1 < chars.len() && chars[i + 1] == '=' {
                    i += 1;
                }
                Token::Eq
            }
            '!' => {
                if i + 1 < chars.len() && chars[i + 1] == '=' {
                    i += 1;
                    Token::Neq
                } else {
                    return Err(LexError::UnexpectedChar(c));
                }
            }
            _ => return Err(LexError::UnexpectedChar(c)),
        };
        tokens.push(tok);
        i += 1;
    }

    tokens.push(Token::Eof);
    Ok(tokens)
}

/// Try to parse a string like "A1", "$B$3", "AB12" as a cell reference token.
fn try_parse_cell_ref(s: &str, sheet: Option<String>) -> Option<Token> {
    let bytes = s.as_bytes();
    let mut i = 0;
    let mut abs_col = false;
    let mut abs_row = false;

    if i < bytes.len() && bytes[i] == b'$' {
        abs_col = true;
        i += 1;
    }
    let col_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == col_start {
        return None;
    } // no letters
    let col_str: &str = &s[col_start..i];

    if i < bytes.len() && bytes[i] == b'$' {
        abs_row = true;
        i += 1;
    }
    let row_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i != bytes.len() || i == row_start {
        return None;
    } // trailing chars or no digits

    let row_str: &str = &s[row_start..i];
    let col = asat_core::letter_to_col(col_str)?;
    let row = row_str.parse::<u32>().ok()?.checked_sub(1)?;

    Some(Token::CellRef {
        sheet,
        col,
        row,
        abs_col,
        abs_row,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_lex_basic() {
        let tokens = lex("1+2").unwrap();
        assert_eq!(tokens[0], Token::Number(1.0));
        assert_eq!(tokens[1], Token::Plus);
        assert_eq!(tokens[2], Token::Number(2.0));
    }

    #[test]
    fn test_lex_cell_ref() {
        let tokens = lex("A1").unwrap();
        assert_eq!(
            tokens[0],
            Token::CellRef {
                sheet: None,
                col: 0,
                row: 0,
                abs_col: false,
                abs_row: false
            }
        );
    }

    #[test]
    fn test_lex_range() {
        let tokens = lex("A1:B3").unwrap();
        assert!(matches!(tokens[0], Token::RangeRef { .. }));
    }

    #[test]
    fn test_lex_function() {
        let tokens = lex("SUM(A1:A3)").unwrap();
        assert_eq!(tokens[0], Token::Ident("SUM".to_string()));
        assert_eq!(tokens[1], Token::LParen);
    }
}
