use asat_core::{CellError, CellValue};
use crate::evaluator::{EvalContext, Evaluator, to_number, to_bool, to_text};
use crate::parser::Expr;

/// Dispatch a function call by name
pub fn call(name: &str, args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    // Evaluate all non-range args eagerly; range args are expanded lazily by each function
    match name {
        "SUM"         => fn_sum(args, ctx, ev),
        "AVERAGE" | "AVG" => fn_average(args, ctx, ev),
        "COUNT"       => fn_count(args, ctx, ev),
        "COUNTA"      => fn_counta(args, ctx, ev),
        "MIN"         => fn_min(args, ctx, ev),
        "MAX"         => fn_max(args, ctx, ev),
        "IF"          => fn_if(args, ctx, ev),
        "AND"         => fn_and(args, ctx, ev),
        "OR"          => fn_or(args, ctx, ev),
        "NOT"         => fn_not(args, ctx, ev),
        "ABS"         => fn_abs(args, ctx, ev),
        "ROUND"       => fn_round(args, ctx, ev),
        "ROUNDUP"     => fn_roundup(args, ctx, ev),
        "ROUNDDOWN"   => fn_rounddown(args, ctx, ev),
        "FLOOR"       => fn_floor(args, ctx, ev),
        "CEILING"     => fn_ceiling(args, ctx, ev),
        "MOD"         => fn_mod(args, ctx, ev),
        "POWER"       => fn_power(args, ctx, ev),
        "SQRT"        => fn_sqrt(args, ctx, ev),
        "LN"          => fn_ln(args, ctx, ev),
        "LOG"         => fn_log(args, ctx, ev),
        "LOG10"       => fn_log10(args, ctx, ev),
        "EXP"         => fn_exp(args, ctx, ev),
        "INT"         => fn_int(args, ctx, ev),
        "TRUNC"       => fn_trunc(args, ctx, ev),
        "SIGN"        => fn_sign(args, ctx, ev),

        "LEN"         => fn_len(args, ctx, ev),
        "LEFT"        => fn_left(args, ctx, ev),
        "RIGHT"       => fn_right(args, ctx, ev),
        "MID"         => fn_mid(args, ctx, ev),
        "TRIM"        => fn_trim(args, ctx, ev),
        "UPPER"       => fn_upper(args, ctx, ev),
        "LOWER"       => fn_lower(args, ctx, ev),
        "PROPER"      => fn_proper(args, ctx, ev),
        "CONCATENATE" | "CONCAT" => fn_concat(args, ctx, ev),
        "TEXT"        => fn_text(args, ctx, ev),
        "VALUE"       => fn_value(args, ctx, ev),
        "FIND"        => fn_find(args, ctx, ev),
        "SEARCH"      => fn_search(args, ctx, ev),
        "SUBSTITUTE"  => fn_substitute(args, ctx, ev),
        "REPLACE"     => fn_replace(args, ctx, ev),
        "REPT"        => fn_rept(args, ctx, ev),

        "ISNUMBER"    => fn_isnumber(args, ctx, ev),
        "ISTEXT"      => fn_istext(args, ctx, ev),
        "ISBLANK"     => fn_isblank(args, ctx, ev),
        "ISERROR"     => fn_iserror(args, ctx, ev),
        "IFERROR"     => fn_iferror(args, ctx, ev),
        "ISLOGICAL"   => fn_islogical(args, ctx, ev),

        "TRUE"        => CellValue::Boolean(true),
        "FALSE"       => CellValue::Boolean(false),
        "PI"          => CellValue::Number(std::f64::consts::PI),

        _ => {
            // Fall through to plugin-registered custom functions before giving up.
            if asat_core::has_custom_fn(name) {
                // Evaluate all args eagerly (no range support for custom fns in v1)
                let evaluated: Vec<CellValue> = args.iter().map(|a| ev.eval(a, ctx)).collect();
                asat_core::call_custom_fn(name, &evaluated)
                    .unwrap_or(CellValue::Error(CellError::Value))
            } else {
                CellValue::Error(CellError::Name)
            }
        }
    }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/// Expand all args (including ranges) into flat list of CellValues
fn expand_args(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> Vec<CellValue> {
    args.iter().flat_map(|a| ev.expand_range(a, ctx)).collect()
}

fn first_num(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> Option<f64> {
    args.first().and_then(|a| to_number(&ev.eval(a, ctx)))
}

fn first_two_nums(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> Option<(f64, f64)> {
    if args.len() < 2 { return None; }
    let a = to_number(&ev.eval(&args[0], ctx))?;
    let b = to_number(&ev.eval(&args[1], ctx))?;
    Some((a, b))
}

fn require_args(args: &[Expr], n: usize) -> Result<(), CellValue> {
    if args.len() < n {
        Err(CellValue::Error(CellError::Value))
    } else {
        Ok(())
    }
}

// ── Math Functions ────────────────────────────────────────────────────────────

fn fn_sum(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let vals = expand_args(args, ctx, ev);
    let mut sum = 0.0;
    for v in &vals {
        match v {
            CellValue::Error(e) => return CellValue::Error(e.clone()),
            _ => if let Some(n) = to_number(v) { sum += n; }
        }
    }
    CellValue::Number(sum)
}

fn fn_average(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let vals = expand_args(args, ctx, ev);
    let mut sum = 0.0;
    let mut cnt = 0u32;
    for v in &vals {
        match v {
            CellValue::Error(e) => return CellValue::Error(e.clone()),
            CellValue::Empty => {}
            _ => if let Some(n) = to_number(v) { sum += n; cnt += 1; }
        }
    }
    if cnt == 0 { CellValue::Error(CellError::Div0) } else { CellValue::Number(sum / cnt as f64) }
}

fn fn_count(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let vals = expand_args(args, ctx, ev);
    let cnt = vals.iter().filter(|v| matches!(v, CellValue::Number(_))).count();
    CellValue::Number(cnt as f64)
}

fn fn_counta(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let vals = expand_args(args, ctx, ev);
    let cnt = vals.iter().filter(|v| !matches!(v, CellValue::Empty)).count();
    CellValue::Number(cnt as f64)
}

fn fn_min(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let vals = expand_args(args, ctx, ev);
    let nums: Vec<f64> = vals.iter().filter_map(|v| to_number(v)).collect();
    if nums.is_empty() { return CellValue::Error(CellError::Value); }
    CellValue::Number(nums.iter().cloned().fold(f64::INFINITY, f64::min))
}

fn fn_max(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let vals = expand_args(args, ctx, ev);
    let nums: Vec<f64> = vals.iter().filter_map(|v| to_number(v)).collect();
    if nums.is_empty() { return CellValue::Error(CellError::Value); }
    CellValue::Number(nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max))
}

fn fn_abs(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) { CellValue::Number(n.abs()) }
    else { CellValue::Error(CellError::Value) }
}

fn fn_round(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some((n, places)) = first_two_nums(args, ctx, ev) {
        let factor = 10f64.powi(places as i32);
        CellValue::Number((n * factor).round() / factor)
    } else { CellValue::Error(CellError::Value) }
}

fn fn_roundup(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some((n, places)) = first_two_nums(args, ctx, ev) {
        let factor = 10f64.powi(places as i32);
        CellValue::Number((n * factor).ceil() / factor)
    } else { CellValue::Error(CellError::Value) }
}

fn fn_rounddown(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some((n, places)) = first_two_nums(args, ctx, ev) {
        let factor = 10f64.powi(places as i32);
        CellValue::Number((n * factor).floor() / factor)
    } else { CellValue::Error(CellError::Value) }
}

fn fn_floor(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) { CellValue::Number(n.floor()) }
    else { CellValue::Error(CellError::Value) }
}

fn fn_ceiling(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) { CellValue::Number(n.ceil()) }
    else { CellValue::Error(CellError::Value) }
}

fn fn_mod(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some((n, d)) = first_two_nums(args, ctx, ev) {
        if d == 0.0 { CellValue::Error(CellError::Div0) }
        else { CellValue::Number(n % d) }
    } else { CellValue::Error(CellError::Value) }
}

fn fn_power(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some((n, e)) = first_two_nums(args, ctx, ev) { CellValue::Number(n.powf(e)) }
    else { CellValue::Error(CellError::Value) }
}

fn fn_sqrt(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) {
        if n < 0.0 { CellValue::Error(CellError::Num) }
        else { CellValue::Number(n.sqrt()) }
    } else { CellValue::Error(CellError::Value) }
}

fn fn_ln(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) {
        if n <= 0.0 { CellValue::Error(CellError::Num) }
        else { CellValue::Number(n.ln()) }
    } else { CellValue::Error(CellError::Value) }
}

fn fn_log(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.is_empty() { return CellValue::Error(CellError::Value); }
    let n = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let base = if args.len() > 1 { to_number(&ev.eval(&args[1], ctx)).unwrap_or(10.0) } else { 10.0 };
    if n <= 0.0 || base <= 0.0 || base == 1.0 { CellValue::Error(CellError::Num) }
    else { CellValue::Number(n.log(base)) }
}

fn fn_log10(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) {
        if n <= 0.0 { CellValue::Error(CellError::Num) }
        else { CellValue::Number(n.log10()) }
    } else { CellValue::Error(CellError::Value) }
}

fn fn_exp(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) { CellValue::Number(n.exp()) }
    else { CellValue::Error(CellError::Value) }
}

fn fn_int(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) { CellValue::Number(n.floor()) }
    else { CellValue::Error(CellError::Value) }
}

fn fn_trunc(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) { CellValue::Number(n.trunc()) }
    else { CellValue::Error(CellError::Value) }
}

fn fn_sign(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(n) = first_num(args, ctx, ev) {
        CellValue::Number(if n > 0.0 { 1.0 } else if n < 0.0 { -1.0 } else { 0.0 })
    } else { CellValue::Error(CellError::Value) }
}

// ── Logical Functions ─────────────────────────────────────────────────────────

fn fn_if(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let cond = ev.eval(&args[0], ctx);
    match to_bool(&cond) {
        Some(true) => ev.eval(&args[1], ctx),
        Some(false) => args.get(2).map(|e| ev.eval(e, ctx)).unwrap_or(CellValue::Boolean(false)),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_and(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    for a in args {
        match to_bool(&ev.eval(a, ctx)) {
            Some(false) => return CellValue::Boolean(false),
            None => return CellValue::Error(CellError::Value),
            _ => {}
        }
    }
    CellValue::Boolean(true)
}

fn fn_or(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    for a in args {
        match to_bool(&ev.eval(a, ctx)) {
            Some(true) => return CellValue::Boolean(true),
            None => return CellValue::Error(CellError::Value),
            _ => {}
        }
    }
    CellValue::Boolean(false)
}

fn fn_not(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if let Some(b) = args.first().and_then(|a| to_bool(&ev.eval(a, ctx))) {
        CellValue::Boolean(!b)
    } else { CellValue::Error(CellError::Value) }
}

fn fn_iferror(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let v = ev.eval(&args[0], ctx);
    if matches!(v, CellValue::Error(_)) {
        ev.eval(&args[1], ctx)
    } else {
        v
    }
}

// ── Type-check Functions ──────────────────────────────────────────────────────

fn fn_isnumber(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let v = args.first().map(|a| ev.eval(a, ctx)).unwrap_or(CellValue::Empty);
    CellValue::Boolean(matches!(v, CellValue::Number(_)))
}

fn fn_istext(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let v = args.first().map(|a| ev.eval(a, ctx)).unwrap_or(CellValue::Empty);
    CellValue::Boolean(matches!(v, CellValue::Text(_)))
}

fn fn_isblank(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let v = args.first().map(|a| ev.eval(a, ctx)).unwrap_or(CellValue::Empty);
    CellValue::Boolean(matches!(v, CellValue::Empty))
}

fn fn_iserror(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let v = args.first().map(|a| ev.eval(a, ctx)).unwrap_or(CellValue::Empty);
    CellValue::Boolean(matches!(v, CellValue::Error(_)))
}

fn fn_islogical(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let v = args.first().map(|a| ev.eval(a, ctx)).unwrap_or(CellValue::Empty);
    CellValue::Boolean(matches!(v, CellValue::Boolean(_)))
}

// ── String Functions ──────────────────────────────────────────────────────────

fn fn_len(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let v = args.first().map(|a| ev.eval(a, ctx)).unwrap_or(CellValue::Empty);
    CellValue::Number(to_text(&v).chars().count() as f64)
}

fn fn_left(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.is_empty() { return CellValue::Error(CellError::Value); }
    let s = to_text(&ev.eval(&args[0], ctx));
    let n = if args.len() > 1 { to_number(&ev.eval(&args[1], ctx)).unwrap_or(1.0) as usize } else { 1 };
    CellValue::Text(s.chars().take(n).collect())
}

fn fn_right(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.is_empty() { return CellValue::Error(CellError::Value); }
    let s = to_text(&ev.eval(&args[0], ctx));
    let n = if args.len() > 1 { to_number(&ev.eval(&args[1], ctx)).unwrap_or(1.0) as usize } else { 1 };
    let chars: Vec<char> = s.chars().collect();
    let start = chars.len().saturating_sub(n);
    CellValue::Text(chars[start..].iter().collect())
}

fn fn_mid(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let s = to_text(&ev.eval(&args[0], ctx));
    let start = to_number(&ev.eval(&args[1], ctx)).unwrap_or(1.0) as usize;
    let len = to_number(&ev.eval(&args[2], ctx)).unwrap_or(0.0) as usize;
    let chars: Vec<char> = s.chars().collect();
    let start = start.saturating_sub(1);
    CellValue::Text(chars.get(start..).unwrap_or(&[]).iter().take(len).collect())
}

fn fn_trim(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let s = args.first().map(|a| to_text(&ev.eval(a, ctx))).unwrap_or_default();
    // Excel TRIM: leading/trailing whitespace + collapse internal whitespace to single space
    let trimmed = s.split_whitespace().collect::<Vec<_>>().join(" ");
    CellValue::Text(trimmed)
}

fn fn_upper(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let s = args.first().map(|a| to_text(&ev.eval(a, ctx))).unwrap_or_default();
    CellValue::Text(s.to_uppercase())
}

fn fn_lower(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let s = args.first().map(|a| to_text(&ev.eval(a, ctx))).unwrap_or_default();
    CellValue::Text(s.to_lowercase())
}

fn fn_proper(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let s = args.first().map(|a| to_text(&ev.eval(a, ctx))).unwrap_or_default();
    let result: String = s.chars().enumerate().map(|(i, c)| {
        // Capitalize after whitespace or start
        let prev = s.chars().nth(i.saturating_sub(1));
        if i == 0 || prev.map(|p| !p.is_alphabetic()).unwrap_or(true) {
            c.to_uppercase().next().unwrap_or(c)
        } else {
            c.to_lowercase().next().unwrap_or(c)
        }
    }).collect();
    CellValue::Text(result)
}

fn fn_concat(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let vals = expand_args(args, ctx, ev);
    let mut result = String::new();
    for v in &vals {
        if let CellValue::Error(e) = v { return CellValue::Error(e.clone()); }
        result.push_str(&to_text(v));
    }
    CellValue::Text(result)
}

fn fn_text(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.is_empty() { return CellValue::Error(CellError::Value); }
    let v = ev.eval(&args[0], ctx);
    // For now, just convert to display string (full format support is Phase 5)
    CellValue::Text(to_text(&v))
}

fn fn_value(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let s = args.first().map(|a| to_text(&ev.eval(a, ctx))).unwrap_or_default();
    if let Ok(n) = s.trim().parse::<f64>() {
        CellValue::Number(n)
    } else {
        CellValue::Error(CellError::Value)
    }
}

fn fn_find(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let find = to_text(&ev.eval(&args[0], ctx));
    let within = to_text(&ev.eval(&args[1], ctx));
    let start = if args.len() > 2 {
        to_number(&ev.eval(&args[2], ctx)).unwrap_or(1.0) as usize - 1
    } else { 0 };
    if let Some(pos) = within[start..].find(&find) {
        CellValue::Number((start + pos + 1) as f64)
    } else {
        CellValue::Error(CellError::Value)
    }
}

fn fn_search(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    // Case-insensitive find
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let find = to_text(&ev.eval(&args[0], ctx)).to_lowercase();
    let within = to_text(&ev.eval(&args[1], ctx)).to_lowercase();
    let start = if args.len() > 2 {
        to_number(&ev.eval(&args[2], ctx)).unwrap_or(1.0) as usize - 1
    } else { 0 };
    if let Some(pos) = within[start..].find(&find) {
        CellValue::Number((start + pos + 1) as f64)
    } else {
        CellValue::Error(CellError::Value)
    }
}

fn fn_substitute(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let text = to_text(&ev.eval(&args[0], ctx));
    let old = to_text(&ev.eval(&args[1], ctx));
    let new = to_text(&ev.eval(&args[2], ctx));
    // Optional: which occurrence (1-indexed); if absent, replace all
    let instance = if args.len() > 3 {
        to_number(&ev.eval(&args[3], ctx)).map(|n| n as usize)
    } else { None };

    if let Some(n) = instance {
        let mut result = text.clone();
        let mut count = 0;
        let mut search_start = 0;
        while let Some(pos) = result[search_start..].find(&old) {
            count += 1;
            let abs_pos = search_start + pos;
            if count == n {
                result = format!("{}{}{}", &result[..abs_pos], &new, &result[abs_pos + old.len()..]);
                break;
            }
            search_start = abs_pos + old.len();
        }
        CellValue::Text(result)
    } else {
        CellValue::Text(text.replace(&old, &new))
    }
}

fn fn_replace(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 4 { return CellValue::Error(CellError::Value); }
    let text = to_text(&ev.eval(&args[0], ctx));
    let start = to_number(&ev.eval(&args[1], ctx)).unwrap_or(1.0) as usize;
    let num_chars = to_number(&ev.eval(&args[2], ctx)).unwrap_or(0.0) as usize;
    let new_text = to_text(&ev.eval(&args[3], ctx));
    let chars: Vec<char> = text.chars().collect();
    let start = start.saturating_sub(1);
    let end = (start + num_chars).min(chars.len());
    let mut result: Vec<char> = chars[..start].to_vec();
    result.extend(new_text.chars());
    result.extend_from_slice(&chars[end..]);
    CellValue::Text(result.iter().collect())
}

fn fn_rept(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let text = to_text(&ev.eval(&args[0], ctx));
    let n = to_number(&ev.eval(&args[1], ctx)).unwrap_or(0.0) as usize;
    CellValue::Text(text.repeat(n))
}
