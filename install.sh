#!/usr/bin/env bash
set -e

INSTALL_DIR="${HOME}/.local/bin"
BINARY="asat"

echo "Building ASAT (release)..."
cargo build --release

echo "Installing to ${INSTALL_DIR}/${BINARY}..."
mkdir -p "${INSTALL_DIR}"
cp "target/release/${BINARY}" "${INSTALL_DIR}/${BINARY}"
chmod +x "${INSTALL_DIR}/${BINARY}"

echo "Done. ASAT installed to ${INSTALL_DIR}/${BINARY}"

# Warn if ~/.local/bin is not on PATH
if ! echo "${PATH}" | grep -q "${INSTALL_DIR}"; then
    echo ""
    echo "  Note: ${INSTALL_DIR} is not in your PATH."
    echo "  Add this to your ~/.bashrc or ~/.zshrc:"
    echo ""
    echo "    export PATH=\"\${HOME}/.local/bin:\${PATH}\""
    echo ""
fi
