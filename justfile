# uv selects/creates python-bindings/.venv and supplies maturin.
# No shell activation is required; Windows uses cmd.exe without PowerShell.
[windows]
set shell := ["cmd.exe", "/c"]

# Build release wheels. Example: just build "--interpreter python3.15"
build args="":
    uv run --project python-bindings --group dev maturin build --release --locked --manifest-path python-bindings/Cargo.toml {{args}}

# Build a source archive that uv can resolve across Python versions.
sdist args="":
    uv run --project python-bindings --group dev maturin sdist --manifest-path python-bindings/Cargo.toml --out dist {{args}}

# Install the extension into python-bindings/.venv for local development.
dev args="":
    uv run --project python-bindings --group dev maturin develop --release --locked --manifest-path python-bindings/Cargo.toml {{args}}

# Run Python integration tests using the project's managed environment.
test args="":
    uv run --project python-bindings --group dev --extra test python -m pytest python-bindings/tests -q {{args}}
