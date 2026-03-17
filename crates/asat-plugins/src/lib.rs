//! Plugin system for ASAT.
//!
//! Enable with `--features asat-plugins/python` (requires Python 3.8+).
//! Without the feature all methods are no-ops.
//!
//! # init.py example
//!
//! ```python
//! import asat
//!
//! @asat.on("open")
//! def on_open(path):
//!     asat.notify("Opened: " + str(path))
//!
//! @asat.on("cell_change")
//! def on_change(sheet, row, col, old_val, new_val):
//!     if isinstance(new_val, (int, float)) and new_val > 1_000_000:
//!         asat.notify("Very large value entered!")
//!
//! @asat.function("DOUBLE")
//! def double(x):
//!     return x * 2
//!
//! @asat.function("GREET")
//! def greet(name):
//!     return "Hello, " + str(name) + "!"
//! ```

use asat_core::{CellValue, Workbook};

// ── Events ────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum PluginEvent {
    Open {
        path: Option<String>,
    },
    PreSave {
        path: String,
    },
    PostSave {
        path: String,
    },
    CellChange {
        sheet: usize,
        row: u32,
        col: u32,
        old: CellValue,
        new: CellValue,
    },
    ModeChange {
        mode: String,
    },
    SheetChange {
        from: usize,
        to: usize,
    },
}

impl PluginEvent {
    pub fn event_name(&self) -> &'static str {
        match self {
            PluginEvent::Open { .. } => "open",
            PluginEvent::PreSave { .. } => "pre_save",
            PluginEvent::PostSave { .. } => "post_save",
            PluginEvent::CellChange { .. } => "cell_change",
            PluginEvent::ModeChange { .. } => "mode_change",
            PluginEvent::SheetChange { .. } => "sheet_change",
        }
    }
}

// ── Plugin outputs ────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum PluginOutput {
    Notify(String),
    Command(String),
    /// `sheet == usize::MAX` → use caller's active sheet.
    SetCell {
        sheet: usize,
        row: u32,
        col: u32,
        value: CellValue,
    },
}

// ── Plugin manager ────────────────────────────────────────────────────────────

pub struct PluginManager {
    pending_events: Vec<PluginEvent>,
    #[cfg(feature = "python")]
    initialized: bool,
}

impl Default for PluginManager {
    fn default() -> Self {
        Self::new()
    }
}

impl PluginManager {
    pub fn new() -> Self {
        PluginManager {
            pending_events: Vec::new(),
            #[cfg(feature = "python")]
            initialized: false,
        }
    }

    pub fn push_event(&mut self, event: PluginEvent) {
        self.pending_events.push(event);
    }

    pub fn load_init_script(&mut self) {
        #[cfg(feature = "python")]
        {
            self.initialized = python::init();
        }
    }

    /// Reload init.py without restarting ASAT.
    pub fn reload(&mut self) {
        #[cfg(feature = "python")]
        {
            python::reset();
            self.initialized = python::init();
        }
    }

    /// Human-readable status line for `:plugin list`.
    pub fn info(&self) -> String {
        #[cfg(feature = "python")]
        {
            let handler_count = python::handler_count();
            let fn_count = asat_core::list_custom_fns().len();
            if self.initialized {
                return format!(
                    "Plugin engine: active — {} event handler(s), {} custom function(s)  \
                     [init.py: ~/.config/asat/init.py]",
                    handler_count, fn_count
                );
            } else {
                return "Plugin engine: disabled (no init.py or load error)".to_string();
            }
        }
        #[cfg(not(feature = "python"))]
        "Plugin engine: not compiled in (rebuild with --features asat-plugins/python)".to_string()
    }

    pub fn drain(&mut self, workbook: &Workbook) -> Vec<PluginOutput> {
        let events = std::mem::take(&mut self.pending_events);
        #[cfg(feature = "python")]
        if self.initialized && !events.is_empty() {
            return python::dispatch(events, workbook);
        }
        let _ = (events, workbook);
        Vec::new()
    }
}

// ── Python back-end ───────────────────────────────────────────────────────────

#[cfg(feature = "python")]
mod python {
    use std::collections::HashMap;
    use std::sync::Arc;
    use std::sync::{Mutex, OnceLock};

    use pyo3::prelude::*;
    use pyo3::types::{PyDict, PyList, PyModule, PyTuple};

    use super::{PluginEvent, PluginOutput};
    use asat_core::{register_custom_fn, CellError, CellValue, CustomFn, Workbook};

    // ── Global state ─────────────────────────────────────────────────────────

    #[derive(Default)]
    struct State {
        handlers: HashMap<String, Vec<Py<PyAny>>>,
        outputs: Vec<PluginOutput>,
    }

    static STATE: OnceLock<Mutex<State>> = OnceLock::new();
    fn state() -> &'static Mutex<State> {
        STATE.get_or_init(|| Mutex::new(State::default()))
    }

    // Workbook pointer — valid only during dispatch().
    std::thread_local! {
        static WB_PTR: std::cell::Cell<*const Workbook> =
            std::cell::Cell::new(std::ptr::null());
    }

    // ── Value conversions ─────────────────────────────────────────────────────

    fn cv_to_py(py: Python<'_>, v: &CellValue) -> Py<PyAny> {
        match v {
            CellValue::Number(n) => n.into_pyobject(py).unwrap().into_any().unbind(),
            CellValue::Text(s) => s.into_pyobject(py).unwrap().into_any().unbind(),
            CellValue::Boolean(b) => {
                let bound = pyo3::types::PyBool::new(py, *b);
                bound.as_any().clone().unbind()
            }
            CellValue::Empty => py.None(),
            CellValue::Error(e) => format!("#{:?}", e)
                .into_pyobject(py)
                .unwrap()
                .into_any()
                .unbind(),
            CellValue::Formula(f) => f.into_pyobject(py).unwrap().into_any().unbind(),
        }
    }

    fn py_to_cv(py: Python<'_>, obj: &Py<PyAny>) -> CellValue {
        let bound = obj.bind(py);
        if bound.is_none() {
            CellValue::Empty
        } else if let Ok(b) = bound.extract::<bool>() {
            CellValue::Boolean(b)
        } else if let Ok(n) = bound.extract::<f64>() {
            CellValue::Number(n)
        } else if let Ok(n) = bound.extract::<i64>() {
            CellValue::Number(n as f64)
        } else if let Ok(s) = bound.extract::<String>() {
            CellValue::Text(s)
        } else {
            CellValue::Text(bound.str().map(|s| s.to_string()).unwrap_or_default())
        }
    }

    // ── Native functions exposed to Python ────────────────────────────────────

    /// Internal: register an event handler.
    #[pyfunction]
    fn _register_handler(event_name: String, handler: Py<PyAny>) {
        state()
            .lock()
            .unwrap()
            .handlers
            .entry(event_name.to_ascii_lowercase())
            .or_default()
            .push(handler);
    }

    /// Internal: register a custom formula function.
    #[pyfunction]
    fn _register_fn(py: Python<'_>, func_name: String, handler: Py<PyAny>) {
        let h = Arc::new(handler);
        let f: CustomFn = Arc::new(move |args: &[CellValue]| {
            Python::attach(|py| {
                let py_args: Vec<Py<PyAny>> = args.iter().map(|v| cv_to_py(py, v)).collect();
                let tuple = PyTuple::new(py, &py_args).unwrap();
                match h.bind(py).call1(tuple) {
                    Ok(r) => py_to_cv(py, &r.unbind()),
                    Err(_) => CellValue::Error(CellError::Value),
                }
            })
        });
        register_custom_fn(&func_name.to_ascii_uppercase(), f);
        let _ = py;
    }

    /// Show a notification in the TUI.
    #[pyfunction]
    fn notify(message: String) {
        state()
            .lock()
            .unwrap()
            .outputs
            .push(PluginOutput::Notify(message));
    }

    /// Execute an ex-command.
    #[pyfunction]
    fn command(cmd: String) {
        state()
            .lock()
            .unwrap()
            .outputs
            .push(PluginOutput::Command(cmd));
    }

    /// Read a cell from the active sheet (0-indexed row, col).
    #[pyfunction]
    fn get_cell(py: Python<'_>, row: u32, col: u32) -> Py<PyAny> {
        WB_PTR.with(|ptr| {
            let raw = ptr.get();
            if raw.is_null() {
                return py.None();
            }
            // SAFETY: valid during dispatch() only.
            let wb = unsafe { &*raw };
            cv_to_py(py, wb.active().get_value(row, col))
        })
    }

    /// Queue a cell update.
    #[pyfunction]
    fn set_cell(py: Python<'_>, row: u32, col: u32, value: Py<PyAny>) {
        let cv = py_to_cv(py, &value);
        state().lock().unwrap().outputs.push(PluginOutput::SetCell {
            sheet: usize::MAX,
            row,
            col,
            value: cv,
        });
    }

    // ── The `asat` Python shim (handles decorator API) ────────────────────────

    /// Python code that builds the user-facing `asat` module around the native functions.
    const ASAT_SHIM: &str = r#"
class _Asat:
    """The `asat` module — available inside init.py as `import asat`."""

    def on(self, event_name, handler=None):
        """Register an event handler.  Use as @asat.on("open") or asat.on("open", fn)."""
        if handler is not None:
            _register_handler(event_name, handler)
            return handler
        def decorator(fn):
            _register_handler(event_name, fn)
            return fn
        return decorator

    def function(self, func_name, handler=None):
        """Register a custom formula function.  Use as @asat.function("DOUBLE") or directly."""
        if handler is not None:
            _register_fn(func_name, handler)
            return handler
        def decorator(fn):
            _register_fn(func_name, fn)
            return fn
        return decorator

    def notify(self, message):
        """Show a notification bubble in the TUI."""
        _notify(str(message))

    def command(self, cmd):
        """Execute an ex-command (e.g. ':w')."""
        _command(str(cmd))

    def get_cell(self, row, col):
        """Read a cell value from the active sheet (0-indexed row, col)."""
        return _get_cell(row, col)

    def set_cell(self, row, col, value):
        """Set a cell value in the active sheet (0-indexed row, col)."""
        _set_cell(row, col, value)

    def read(self, address):
        """Read a cell by Excel-style address e.g. asat.read('B3')."""
        row, col = self._parse_address(address)
        return _get_cell(row, col)

    def write(self, address, value):
        """Write a cell by Excel-style address e.g. asat.write('B3', 42)."""
        row, col = self._parse_address(address)
        _set_cell(row, col, value)

    def _parse_address(self, address):
        """Convert 'B3' → (row=2, col=1) (0-indexed)."""
        addr = address.strip().upper()
        col_str = ''
        row_str = ''
        for ch in addr:
            if ch.isalpha():
                col_str += ch
            else:
                row_str += ch
        if not col_str or not row_str:
            raise ValueError(f"Invalid cell address: {address!r}")
        col = 0
        for ch in col_str:
            col = col * 26 + (ord(ch) - ord('A') + 1)
        col -= 1  # 0-indexed
        row = int(row_str) - 1  # 0-indexed
        return row, col

asat = _Asat()
"#;

    // ── Reset (for reload) ────────────────────────────────────────────────────

    /// Clear all registered handlers and custom functions, ready for re-init.
    pub fn reset() {
        if let Ok(mut s) = state().lock() {
            s.handlers.clear();
            s.outputs.clear();
        }
        // Clear plugin-registered custom formula functions
        for name in asat_core::list_custom_fns() {
            asat_core::unregister_custom_fn(&name);
        }
    }

    /// Count total registered event handlers (for :plugin list).
    pub fn handler_count() -> usize {
        state()
            .lock()
            .map(|s| s.handlers.values().map(|v| v.len()).sum())
            .unwrap_or(0)
    }

    // ── Init ─────────────────────────────────────────────────────────────────

    pub fn init() -> bool {
        let script_path = script_path();
        if !script_path.exists() {
            return true; // no script — Python still "initialised"
        }

        let user_code = match std::fs::read_to_string(&script_path) {
            Ok(c) => c,
            Err(e) => {
                eprintln!("[asat] Cannot read init.py: {}", e);
                return false;
            }
        };

        match Python::attach(|py| -> PyResult<()> {
            // Build the native module
            let native = PyModule::new(py, "_asat_native")?;
            native.add_function(wrap_pyfunction!(_register_handler, &native)?)?;
            native.add_function(wrap_pyfunction!(_register_fn, &native)?)?;
            native.add_function(wrap_pyfunction!(notify, &native)?)?;
            native.add_function(wrap_pyfunction!(command, &native)?)?;
            native.add_function(wrap_pyfunction!(get_cell, &native)?)?;
            native.add_function(wrap_pyfunction!(set_cell, &native)?)?;

            // Globals for the shim + user script
            let globals = PyDict::new(py);
            globals.set_item("__builtins__", py.import("builtins")?)?;
            // Expose native functions directly into globals
            globals.set_item("_register_handler", native.getattr("_register_handler")?)?;
            globals.set_item("_register_fn", native.getattr("_register_fn")?)?;
            globals.set_item("_notify", native.getattr("notify")?)?;
            globals.set_item("_command", native.getattr("command")?)?;
            globals.set_item("_get_cell", native.getattr("get_cell")?)?;
            globals.set_item("_set_cell", native.getattr("set_cell")?)?;

            // Run the shim — defines `asat` in globals
            let shim_cstr =
                std::ffi::CString::new(ASAT_SHIM).expect("ASAT_SHIM contains no null bytes");
            py.run(shim_cstr.as_c_str(), Some(&globals), None)?;

            // Now run the user's init.py
            let code_cstr = std::ffi::CString::new(user_code.as_str()).map_err(|_| {
                pyo3::exceptions::PySyntaxError::new_err("init.py contains null bytes")
            })?;
            py.run(code_cstr.as_c_str(), Some(&globals), None)?;

            Ok(())
        }) {
            Ok(_) => {
                eprintln!("[asat] init.py loaded");
                true
            }
            Err(e) => {
                eprintln!("[asat] init.py error:\n{}", e);
                false
            }
        }
    }

    // ── Dispatch ─────────────────────────────────────────────────────────────

    pub fn dispatch(events: Vec<PluginEvent>, wb: &Workbook) -> Vec<PluginOutput> {
        WB_PTR.with(|ptr| ptr.set(wb as *const Workbook));

        Python::attach(|py| {
            for event in &events {
                // Snapshot handlers while holding the lock, then release before calling
                let handlers: Vec<Py<PyAny>> = {
                    let s = state().lock().unwrap();
                    match s.handlers.get(event.event_name()) {
                        Some(v) => v.iter().map(|h| h.clone_ref(py)).collect(),
                        None => continue,
                    }
                };

                let args = make_args(py, event);
                let tuple = match PyTuple::new(py, &args) {
                    Ok(t) => t,
                    Err(_) => continue,
                };

                for h in handlers {
                    if let Err(e) = h.bind(py).call1(tuple.clone()) {
                        eprintln!("[asat] handler error ({}): {}", event.event_name(), e);
                    }
                }
            }
        });

        WB_PTR.with(|ptr| ptr.set(std::ptr::null()));

        std::mem::take(&mut state().lock().unwrap().outputs)
    }

    fn make_args(py: Python<'_>, event: &PluginEvent) -> Vec<Py<PyAny>> {
        match event {
            PluginEvent::Open { path } => vec![path
                .as_deref()
                .map(|s| s.into_pyobject(py).unwrap().into_any().unbind())
                .unwrap_or_else(|| py.None())],
            PluginEvent::PreSave { path } | PluginEvent::PostSave { path } => {
                vec![path.into_pyobject(py).unwrap().into_any().unbind()]
            }
            PluginEvent::CellChange {
                sheet,
                row,
                col,
                old,
                new,
            } => vec![
                sheet.into_pyobject(py).unwrap().into_any().unbind(),
                row.into_pyobject(py).unwrap().into_any().unbind(),
                col.into_pyobject(py).unwrap().into_any().unbind(),
                cv_to_py(py, old),
                cv_to_py(py, new),
            ],
            PluginEvent::ModeChange { mode } => {
                vec![mode.into_pyobject(py).unwrap().into_any().unbind()]
            }
            PluginEvent::SheetChange { from, to } => vec![
                from.into_pyobject(py).unwrap().into_any().unbind(),
                to.into_pyobject(py).unwrap().into_any().unbind(),
            ],
        }
    }

    fn script_path() -> std::path::PathBuf {
        let base = std::env::var("XDG_CONFIG_HOME")
            .map(std::path::PathBuf::from)
            .unwrap_or_else(|_| {
                dirs::home_dir()
                    .unwrap_or_else(|| std::path::PathBuf::from("."))
                    .join(".config")
            });
        base.join("asat").join("init.py")
    }
}
