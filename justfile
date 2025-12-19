set shell := ["pwsh.exe", "-c"]

build POLARS="":
    python-bindings\.venv\Scripts\activate
    cd python-bindings && maturin build --release {{POLARS}}

dev POLARS="":
    python-bindings\.venv\Scripts\activate
    cd python-bindings && maturin develop --release {{POLARS}}
