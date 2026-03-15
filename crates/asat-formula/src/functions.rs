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

        // ── Statistical ──
        "SUMIF"       => fn_sumif(args, ctx, ev),
        "COUNTIF"     => fn_countif(args, ctx, ev),
        "SUMPRODUCT"  => fn_sumproduct(args, ctx, ev),
        "LARGE"       => fn_large(args, ctx, ev),
        "SMALL"       => fn_small(args, ctx, ev),
        "MEDIAN"      => fn_median(args, ctx, ev),
        "STDEV"       => fn_stdev(args, ctx, ev),
        "VAR"         => fn_var(args, ctx, ev),

        // ── Finance ──
        "PV"          => fn_pv(args, ctx, ev),
        "FV"          => fn_fv(args, ctx, ev),
        "PMT"         => fn_pmt(args, ctx, ev),
        "NPER"        => fn_nper(args, ctx, ev),
        "RATE"        => fn_rate(args, ctx, ev),
        "NPV"         => fn_npv(args, ctx, ev),
        "IRR"         => fn_irr(args, ctx, ev),
        "MIRR"        => fn_mirr(args, ctx, ev),
        "IPMT"        => fn_ipmt(args, ctx, ev),
        "PPMT"        => fn_ppmt(args, ctx, ev),
        "SLN"         => fn_sln(args, ctx, ev),
        "DDB"         => fn_ddb(args, ctx, ev),
        "EFFECT"      => fn_effect(args, ctx, ev),
        "NOMINAL"     => fn_nominal(args, ctx, ev),
        "CUMIPMT"     => fn_cumipmt(args, ctx, ev),
        "CUMPRINC"    => fn_cumprinc(args, ctx, ev),

        "AVERAGEIF"   => fn_averageif(args, ctx, ev),
        "MAXIFS"      => fn_maxifs(args, ctx, ev),
        "MINIFS"      => fn_minifs(args, ctx, ev),
        "RANK"        => fn_rank(args, ctx, ev),
        "PERCENTILE"  => fn_percentile(args, ctx, ev),
        "QUARTILE"    => fn_quartile(args, ctx, ev),
        "XLOOKUP"     => fn_xlookup(args, ctx, ev),
        "CHOOSE"      => fn_choose(args, ctx, ev),

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

// ── Statistical Functions ─────────────────────────────────────────────────────

/// SUMIF(range, criteria, [sum_range])
fn fn_sumif(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let test_vals = ev.expand_range(&args[0], ctx);
    let criteria  = ev.eval(&args[1], ctx);
    let sum_vals  = if args.len() > 2 { ev.expand_range(&args[2], ctx) } else { test_vals.clone() };
    let mut sum = 0.0;
    for (tv, sv) in test_vals.iter().zip(sum_vals.iter()) {
        if criteria_match(tv, &criteria) {
            if let Some(n) = to_number(sv) { sum += n; }
        }
    }
    CellValue::Number(sum)
}

/// COUNTIF(range, criteria)
fn fn_countif(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let vals     = ev.expand_range(&args[0], ctx);
    let criteria = ev.eval(&args[1], ctx);
    let count    = vals.iter().filter(|v| criteria_match(v, &criteria)).count();
    CellValue::Number(count as f64)
}

/// Match a cell value against a SUMIF/COUNTIF criteria.
/// Supports: number equality, text equality (case-insensitive), and
/// comparison strings like ">10", "<=5", "<>0".
fn criteria_match(val: &CellValue, criteria: &CellValue) -> bool {
    match criteria {
        CellValue::Number(n) => to_number(val).map(|v| v == *n).unwrap_or(false),
        CellValue::Boolean(b) => matches!(val, CellValue::Boolean(v) if v == b),
        CellValue::Text(s) => {
            // Try comparison operators first: ">5", "<=10", "<>0", ">=3"
            for (op, rest) in [(">=", &s[..]), ("<=", s.as_str()), ("<>", s.as_str()), (">", s.as_str()), ("<", s.as_str())] {
                if s.starts_with(op) {
                    let rhs = &s[op.len()..];
                    if let Ok(rhs_n) = rhs.parse::<f64>() {
                        if let Some(lhs_n) = to_number(val) {
                            return match op {
                                ">="  => lhs_n >= rhs_n,
                                "<="  => lhs_n <= rhs_n,
                                "<>"  => lhs_n != rhs_n,
                                ">"   => lhs_n > rhs_n,
                                "<"   => lhs_n < rhs_n,
                                _     => false,
                            };
                        }
                    }
                    let _ = rest; // suppress unused warning
                    break;
                }
            }
            // Fall back to case-insensitive text match
            match val {
                CellValue::Text(t) => t.to_lowercase() == s.to_lowercase(),
                _ => false,
            }
        }
        _ => false,
    }
}

/// SUMPRODUCT(array1, array2, ...)  — element-wise multiply then sum
fn fn_sumproduct(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.is_empty() { return CellValue::Error(CellError::Value); }
    let arrays: Vec<Vec<f64>> = args.iter()
        .map(|a| ev.expand_range(a, ctx).into_iter().filter_map(|v| to_number(&v)).collect())
        .collect();
    let len = arrays[0].len();
    if arrays.iter().any(|a| a.len() != len) { return CellValue::Error(CellError::Value); }
    let mut sum = 0.0;
    for i in 0..len {
        sum += arrays.iter().map(|a| a[i]).product::<f64>();
    }
    CellValue::Number(sum)
}

/// LARGE(range, k)  — k-th largest value
fn fn_large(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let mut nums: Vec<f64> = ev.expand_range(&args[0], ctx).into_iter().filter_map(|v| to_number(&v)).collect();
    let k = to_number(&ev.eval(&args[1], ctx)).unwrap_or(1.0) as usize;
    if nums.is_empty() || k == 0 || k > nums.len() { return CellValue::Error(CellError::Num); }
    nums.sort_by(|a, b| b.partial_cmp(a).unwrap_or(std::cmp::Ordering::Equal));
    CellValue::Number(nums[k - 1])
}

/// SMALL(range, k)  — k-th smallest value
fn fn_small(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let mut nums: Vec<f64> = ev.expand_range(&args[0], ctx).into_iter().filter_map(|v| to_number(&v)).collect();
    let k = to_number(&ev.eval(&args[1], ctx)).unwrap_or(1.0) as usize;
    if nums.is_empty() || k == 0 || k > nums.len() { return CellValue::Error(CellError::Num); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    CellValue::Number(nums[k - 1])
}

/// MEDIAN(range)
fn fn_median(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let mut nums: Vec<f64> = expand_args(args, ctx, ev).into_iter().filter_map(|v| to_number(&v)).collect();
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let mid = nums.len() / 2;
    let median = if nums.len() % 2 == 0 { (nums[mid - 1] + nums[mid]) / 2.0 } else { nums[mid] };
    CellValue::Number(median)
}

/// STDEV(range)  — sample standard deviation
fn fn_stdev(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let nums: Vec<f64> = expand_args(args, ctx, ev).into_iter().filter_map(|v| to_number(&v)).collect();
    if nums.len() < 2 { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let variance = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64;
    CellValue::Number(variance.sqrt())
}

/// VAR(range)  — sample variance
fn fn_var(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    let nums: Vec<f64> = expand_args(args, ctx, ev).into_iter().filter_map(|v| to_number(&v)).collect();
    if nums.len() < 2 { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let variance = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64;
    CellValue::Number(variance)
}

// ── Financial Functions ───────────────────────────────────────────────────────
//
// Convention (matches Excel):
//   pv   = present value  (money received is positive)
//   fv   = future value   (default 0)
//   pmt  = payment amount per period (cash out is negative)
//   nper = number of periods
//   rate = interest rate per period
//   type = 0 → payments at end of period (ordinary annuity)
//          1 → payments at beginning of period (annuity due)
//
// Core identity (rate ≠ 0):
//   pv·(1+r)^n  +  pmt·(1+r·type)·((1+r)^n − 1)/r  +  fv  =  0

/// Helper: compute (1+rate)^nper, returning Err on domain errors.
fn r1n(rate: f64, nper: f64) -> f64 { (1.0 + rate).powf(nper) }

/// PV(rate, nper, pmt, [fv=0], [type=0])
fn fn_pv(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let rate = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pmt  = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv   = args.get(3).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    let typ  = args.get(4).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    if rate == 0.0 {
        return CellValue::Number(-(pmt * nper + fv));
    }
    let rn = r1n(rate, nper);
    CellValue::Number(-(pmt * (1.0 - 1.0 / rn) / rate * (1.0 + rate * typ) + fv / rn))
}

/// FV(rate, nper, pmt, [pv=0], [type=0])
fn fn_fv(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let rate = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pmt  = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv   = args.get(3).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    let typ  = args.get(4).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    if rate == 0.0 {
        return CellValue::Number(-(pv + pmt * nper));
    }
    let rn = r1n(rate, nper);
    CellValue::Number(-(pv * rn + pmt * (rn - 1.0) / rate * (1.0 + rate * typ)))
}

/// PMT(rate, nper, pv, [fv=0], [type=0])
fn fn_pmt(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let rate = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv   = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv   = args.get(3).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    let typ  = args.get(4).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    if nper == 0.0 { return CellValue::Error(CellError::Div0); }
    if rate == 0.0 {
        return CellValue::Number(-(pv + fv) / nper);
    }
    let rn = r1n(rate, nper);
    CellValue::Number(-(pv * rn + fv) * rate / ((rn - 1.0) * (1.0 + rate * typ)))
}

/// NPER(rate, pmt, pv, [fv=0], [type=0])
fn fn_nper(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let rate = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pmt  = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv   = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv   = args.get(3).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    let typ  = args.get(4).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    if rate == 0.0 {
        if pmt == 0.0 { return CellValue::Error(CellError::Div0); }
        return CellValue::Number(-(pv + fv) / pmt);
    }
    let adj = pmt * (1.0 + rate * typ);
    let num = adj - fv * rate;
    let den = adj + pv * rate;
    if den == 0.0 || num / den <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((num / den).ln() / (1.0 + rate).ln())
}

/// RATE(nper, pmt, pv, [fv=0], [type=0], [guess=0.1])  — Newton-Raphson
fn fn_rate(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let nper = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pmt  = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv   = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv   = args.get(3).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    let typ  = args.get(4).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    let mut rate = args.get(5).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.1);
    // f(r) = pv·(1+r)^n + pmt·(1+r·type)·((1+r)^n−1)/r + fv
    let f = |r: f64| -> f64 {
        let rn = r1n(r, nper);
        pv * rn + pmt * (1.0 + r * typ) * (rn - 1.0) / r + fv
    };
    for _ in 0..200 {
        let fx  = f(rate);
        let dfx = (f(rate + 1e-8) - fx) / 1e-8;
        if dfx.abs() < 1e-15 { break; }
        let new_rate = rate - fx / dfx;
        if (new_rate - rate).abs() < 1e-10 { return CellValue::Number(new_rate); }
        rate = new_rate;
        if rate <= -1.0 { return CellValue::Error(CellError::Num); }
    }
    CellValue::Error(CellError::Num)
}

/// NPV(rate, value1, value2, ...)  — Net present value of cashflows
fn fn_npv(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.is_empty() { return CellValue::Error(CellError::Value); }
    let rate = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let cashflows: Vec<f64> = args[1..].iter()
        .flat_map(|a| ev.expand_range(a, ctx))
        .filter_map(|v| to_number(&v))
        .collect();
    if cashflows.is_empty() { return CellValue::Error(CellError::Value); }
    let npv: f64 = cashflows.iter().enumerate()
        .map(|(i, cf)| cf / r1n(rate, i as f64 + 1.0))
        .sum();
    CellValue::Number(npv)
}

/// IRR(values, [guess=0.1])  — Internal rate of return (Newton-Raphson)
fn fn_irr(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.is_empty() { return CellValue::Error(CellError::Value); }
    let cfs: Vec<f64> = ev.expand_range(&args[0], ctx).into_iter().filter_map(|v| to_number(&v)).collect();
    if cfs.len() < 2 { return CellValue::Error(CellError::Value); }
    let mut rate = args.get(1).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.1);
    let npv = |r: f64| -> f64 {
        cfs.iter().enumerate().map(|(i, &cf)| cf / r1n(r, i as f64)).sum()
    };
    for _ in 0..200 {
        let fx  = npv(rate);
        let dfx = (npv(rate + 1e-8) - fx) / 1e-8;
        if dfx.abs() < 1e-15 { break; }
        let new_rate = rate - fx / dfx;
        if (new_rate - rate).abs() < 1e-10 { return CellValue::Number(new_rate); }
        rate = new_rate;
        if rate <= -1.0 { return CellValue::Error(CellError::Num); }
    }
    CellValue::Error(CellError::Num)
}

/// MIRR(values, finance_rate, reinvest_rate)  — Modified IRR
fn fn_mirr(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let cfs: Vec<f64> = ev.expand_range(&args[0], ctx).into_iter().filter_map(|v| to_number(&v)).collect();
    let fin_r   = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let reinv_r = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let n = cfs.len();
    if n < 2 { return CellValue::Error(CellError::Value); }
    let pv_neg: f64 = cfs.iter().enumerate()
        .filter(|(_, &cf)| cf < 0.0)
        .map(|(i, &cf)| cf / r1n(fin_r, i as f64))
        .sum();
    let fv_pos: f64 = cfs.iter().enumerate()
        .filter(|(_, &cf)| cf > 0.0)
        .map(|(i, &cf)| cf * r1n(reinv_r, (n - 1 - i) as f64))
        .sum();
    if pv_neg == 0.0 || fv_pos == 0.0 { return CellValue::Error(CellError::Div0); }
    CellValue::Number((fv_pos / -pv_neg).powf(1.0 / (n - 1) as f64) - 1.0)
}

/// IPMT(rate, per, nper, pv, [fv=0], [type=0])  — Interest portion of payment
fn fn_ipmt(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 4 { return CellValue::Error(CellError::Value); }
    let rate = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let per  = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv   = match to_number(&ev.eval(&args[3], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv   = args.get(4).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    let typ  = args.get(5).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    if per < 1.0 || per > nper { return CellValue::Error(CellError::Num); }
    // Compute PMT, then balance at start of period, then interest = balance × rate
    let pmt = if rate == 0.0 { -(pv + fv) / nper }
              else { let rn = r1n(rate, nper); -(pv * rn + fv) * rate / ((rn - 1.0) * (1.0 + rate * typ)) };
    let k = per - 1.0 + typ;
    let bal = if rate == 0.0 { pv + pmt * k }
              else { pv * r1n(rate, k) + pmt * (r1n(rate, k) - 1.0) / rate };
    CellValue::Number(bal * rate)
}

/// PPMT(rate, per, nper, pv, [fv=0], [type=0])  — Principal portion of payment
fn fn_ppmt(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    // PPMT = PMT - IPMT
    // Re-build args without the `per` arg for PMT (PMT doesn't take per)
    let pmt_args: Vec<Expr> = [0, 2, 3, 4, 5].iter()
        .filter_map(|&i| args.get(i).cloned())
        .collect();
    match (fn_pmt(&pmt_args, ctx, ev), fn_ipmt(args, ctx, ev)) {
        (CellValue::Number(p), CellValue::Number(i)) => CellValue::Number(p - i),
        (CellValue::Error(e), _) | (_, CellValue::Error(e)) => CellValue::Error(e),
        _ => CellValue::Error(CellError::Value),
    }
}

/// SLN(cost, salvage, life)  — Straight-line depreciation per period
fn fn_sln(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let cost    = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let salvage = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let life    = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if life == 0.0 { return CellValue::Error(CellError::Div0); }
    CellValue::Number((cost - salvage) / life)
}

/// DDB(cost, salvage, life, period, [factor=2])  — Double-declining balance depreciation
fn fn_ddb(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 4 { return CellValue::Error(CellError::Value); }
    let cost    = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let salvage = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let life    = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let period  = match to_number(&ev.eval(&args[3], ctx)) { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let factor  = args.get(4).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(2.0);
    if life == 0.0 { return CellValue::Error(CellError::Div0); }
    let rate = factor / life;
    let mut book = cost;
    let mut dep  = 0.0;
    for _ in 0..period {
        dep   = (book - salvage).max(0.0).min(book * rate);
        book -= dep;
    }
    CellValue::Number(dep)
}

/// EFFECT(nominal_rate, npery)  — Effective annual interest rate
fn fn_effect(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let nom   = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let npery = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if npery < 1.0 || nom <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((1.0 + nom / npery).powf(npery) - 1.0)
}

/// NOMINAL(effect_rate, npery)  — Nominal annual interest rate
fn fn_nominal(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let eff   = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let npery = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if npery < 1.0 || eff <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(((1.0 + eff).powf(1.0 / npery) - 1.0) * npery)
}

/// CUMIPMT(rate, nper, pv, start_period, end_period, type)  — Cumulative interest paid
fn fn_cumipmt(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 6 { return CellValue::Error(CellError::Value); }
    let rate  = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper  = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let pv    = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let start = match to_number(&ev.eval(&args[3], ctx)) { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let end   = match to_number(&ev.eval(&args[4], ctx)) { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let typ   = match to_number(&ev.eval(&args[5], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if start < 1 || end > nper || start > end { return CellValue::Error(CellError::Num); }
    let rn  = r1n(rate, nper as f64);
    let pmt = if rate == 0.0 { -pv / nper as f64 }
              else { -(pv * rn) * rate / ((rn - 1.0) * (1.0 + rate * typ)) };
    let mut total = 0.0;
    for per in start..=end {
        let k   = per as f64 - 1.0 + typ;
        let bal = if rate == 0.0 { pv + pmt * k }
                  else { pv * r1n(rate, k) + pmt * (r1n(rate, k) - 1.0) / rate };
        total += bal * rate;
    }
    CellValue::Number(total)
}

/// CUMPRINC(rate, nper, pv, start_period, end_period, type)  — Cumulative principal paid
fn fn_cumprinc(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 6 { return CellValue::Error(CellError::Value); }
    let rate  = match to_number(&ev.eval(&args[0], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper  = match to_number(&ev.eval(&args[1], ctx)) { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let pv    = match to_number(&ev.eval(&args[2], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let start = match to_number(&ev.eval(&args[3], ctx)) { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let end   = match to_number(&ev.eval(&args[4], ctx)) { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let typ   = match to_number(&ev.eval(&args[5], ctx)) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if start < 1 || end > nper || start > end { return CellValue::Error(CellError::Num); }
    let rn  = r1n(rate, nper as f64);
    let pmt = if rate == 0.0 { -pv / nper as f64 }
              else { -(pv * rn) * rate / ((rn - 1.0) * (1.0 + rate * typ)) };
    let mut total = 0.0;
    for per in start..=end {
        let k   = per as f64 - 1.0 + typ;
        let bal = if rate == 0.0 { pv + pmt * k }
                  else { pv * r1n(rate, k) + pmt * (r1n(rate, k) - 1.0) / rate };
        total += pmt - bal * rate;
    }
    CellValue::Number(total)
}

// ── Extended Statistical Functions ───────────────────────────────────────────

/// AVERAGEIF(range, criteria, [avg_range])
fn fn_averageif(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let test_vals = ev.expand_range(&args[0], ctx);
    let criteria  = ev.eval(&args[1], ctx);
    let avg_vals  = if args.len() > 2 { ev.expand_range(&args[2], ctx) } else { test_vals.clone() };
    let mut sum = 0.0;
    let mut cnt = 0u32;
    for (tv, av) in test_vals.iter().zip(avg_vals.iter()) {
        if criteria_match(tv, &criteria) {
            if let Some(n) = to_number(av) { sum += n; cnt += 1; }
        }
    }
    if cnt == 0 { CellValue::Error(CellError::Div0) } else { CellValue::Number(sum / cnt as f64) }
}

/// MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2, ...])
fn fn_maxifs(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let max_vals = ev.expand_range(&args[0], ctx);
    // Collect pairs of (criteria_range, criteria)
    let mut criteria_pairs: Vec<(Vec<CellValue>, CellValue)> = Vec::new();
    let mut i = 1;
    while i + 1 < args.len() {
        criteria_pairs.push((ev.expand_range(&args[i], ctx), ev.eval(&args[i + 1], ctx)));
        i += 2;
    }
    let mut result = f64::NEG_INFINITY;
    let mut found = false;
    for (idx, mv) in max_vals.iter().enumerate() {
        let all_match = criteria_pairs.iter().all(|(cr, crit)| {
            cr.get(idx).map(|v| criteria_match(v, crit)).unwrap_or(false)
        });
        if all_match {
            if let Some(n) = to_number(mv) {
                if n > result { result = n; found = true; }
            }
        }
    }
    if found { CellValue::Number(result) } else { CellValue::Number(0.0) }
}

/// MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2, ...])
fn fn_minifs(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let min_vals = ev.expand_range(&args[0], ctx);
    let mut criteria_pairs: Vec<(Vec<CellValue>, CellValue)> = Vec::new();
    let mut i = 1;
    while i + 1 < args.len() {
        criteria_pairs.push((ev.expand_range(&args[i], ctx), ev.eval(&args[i + 1], ctx)));
        i += 2;
    }
    let mut result = f64::INFINITY;
    let mut found = false;
    for (idx, mv) in min_vals.iter().enumerate() {
        let all_match = criteria_pairs.iter().all(|(cr, crit)| {
            cr.get(idx).map(|v| criteria_match(v, crit)).unwrap_or(false)
        });
        if all_match {
            if let Some(n) = to_number(mv) {
                if n < result { result = n; found = true; }
            }
        }
    }
    if found { CellValue::Number(result) } else { CellValue::Number(0.0) }
}

/// RANK(number, ref, [order=0])  — 0=descending (largest=1), 1=ascending
fn fn_rank(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let num = match to_number(&ev.eval(&args[0], ctx)) {
        Some(n) => n,
        None    => return CellValue::Error(CellError::Value),
    };
    let vals: Vec<f64> = ev.expand_range(&args[1], ctx).into_iter().filter_map(|v| to_number(&v)).collect();
    let order = args.get(2).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0);
    let ascending = order != 0.0;
    // Count how many values are strictly better (larger if desc, smaller if asc)
    let rank = vals.iter().filter(|&&v| {
        if ascending { v < num } else { v > num }
    }).count() + 1;
    // Check num exists in vals
    if !vals.iter().any(|&v| (v - num).abs() < 1e-10) {
        return CellValue::Error(CellError::NA);
    }
    CellValue::Number(rank as f64)
}

/// PERCENTILE(range, k)  — k is 0..1 (inclusive)
fn fn_percentile(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let mut nums: Vec<f64> = ev.expand_range(&args[0], ctx).into_iter().filter_map(|v| to_number(&v)).collect();
    let k = match to_number(&ev.eval(&args[1], ctx)) {
        Some(n) if (0.0..=1.0).contains(&n) => n,
        _ => return CellValue::Error(CellError::Num),
    };
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let n = nums.len() as f64;
    let pos = k * (n - 1.0);
    let lo = pos.floor() as usize;
    let hi = pos.ceil() as usize;
    if lo == hi {
        CellValue::Number(nums[lo])
    } else {
        let frac = pos - lo as f64;
        CellValue::Number(nums[lo] + frac * (nums[hi] - nums[lo]))
    }
}

/// QUARTILE(range, quart)  — quart: 0=min, 1=Q1, 2=median, 3=Q3, 4=max
fn fn_quartile(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let quart = match to_number(&ev.eval(&args[1], ctx)) {
        Some(n) if (0.0..=4.0).contains(&n) => n as u8,
        _ => return CellValue::Error(CellError::Num),
    };
    let k = match quart {
        0 => 0.0,
        1 => 0.25,
        2 => 0.5,
        3 => 0.75,
        4 => 1.0,
        _ => return CellValue::Error(CellError::Num),
    };
    // Reuse percentile logic
    let percentile_args = &args[..1];
    let mut nums: Vec<f64> = ev.expand_range(&args[0], ctx).into_iter().filter_map(|v| to_number(&v)).collect();
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let n = nums.len() as f64;
    let pos = k * (n - 1.0);
    let lo = pos.floor() as usize;
    let hi = pos.ceil() as usize;
    let _ = percentile_args; // suppress unused warning
    if lo == hi {
        CellValue::Number(nums[lo])
    } else {
        let frac = pos - lo as f64;
        CellValue::Number(nums[lo] + frac * (nums[hi] - nums[lo]))
    }
}

/// XLOOKUP(lookup, lookup_array, return_array, [not_found], [match_mode=0])
/// match_mode: 0=exact, 1=next larger, -1=next smaller, 2=wildcard
fn fn_xlookup(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 3 { return CellValue::Error(CellError::Value); }
    let lookup    = ev.eval(&args[0], ctx);
    let lookup_arr = ev.expand_range(&args[1], ctx);
    let return_arr = ev.expand_range(&args[2], ctx);
    let not_found  = args.get(3).map(|a| ev.eval(a, ctx));
    let match_mode = args.get(4).and_then(|a| to_number(&ev.eval(a, ctx))).unwrap_or(0.0) as i64;

    // Find position of match
    let pos = match match_mode {
        0 => {
            // Exact match
            lookup_arr.iter().position(|v| values_equal(v, &lookup))
        }
        1 => {
            // Next larger or equal
            let target = to_number(&lookup).unwrap_or(0.0);
            lookup_arr.iter()
                .enumerate()
                .filter_map(|(i, v)| to_number(v).map(|n| (i, n)))
                .filter(|(_, n)| *n >= target)
                .min_by(|(_, a), (_, b)| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal))
                .map(|(i, _)| i)
        }
        -1 => {
            // Next smaller or equal
            let target = to_number(&lookup).unwrap_or(0.0);
            lookup_arr.iter()
                .enumerate()
                .filter_map(|(i, v)| to_number(v).map(|n| (i, n)))
                .filter(|(_, n)| *n <= target)
                .max_by(|(_, a), (_, b)| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal))
                .map(|(i, _)| i)
        }
        _ => lookup_arr.iter().position(|v| values_equal(v, &lookup)),
    };

    match pos {
        Some(i) => return_arr.get(i).cloned().unwrap_or(CellValue::Error(CellError::NA)),
        None    => not_found.unwrap_or(CellValue::Error(CellError::NA)),
    }
}

/// CHOOSE(index, val1, val2, ...)  — 1-based index
fn fn_choose(args: &[Expr], ctx: &EvalContext<'_>, ev: &Evaluator) -> CellValue {
    if args.len() < 2 { return CellValue::Error(CellError::Value); }
    let idx = match to_number(&ev.eval(&args[0], ctx)) {
        Some(n) if n >= 1.0 => n as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    let choices = &args[1..];
    if idx > choices.len() { return CellValue::Error(CellError::Value); }
    ev.eval(&choices[idx - 1], ctx)
}

/// Helper: check if two CellValues are equal for XLOOKUP exact match
fn values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Number(x), CellValue::Number(y)) => (x - y).abs() < 1e-10,
        (CellValue::Text(x), CellValue::Text(y))     => x.eq_ignore_ascii_case(y),
        (CellValue::Boolean(x), CellValue::Boolean(y)) => x == y,
        (CellValue::Empty, CellValue::Empty) => true,
        _ => false,
    }
}
