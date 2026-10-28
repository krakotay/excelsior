"""Reject release tags/inputs that do not match the version being built."""
import os
from pathlib import Path
import tomllib

version = tomllib.loads(Path("Cargo.toml").read_text())["workspace"]["package"]["version"]
expected = os.environ.get("EXPECTED_VERSION", "").removeprefix("v")
if expected and expected != version:
    raise SystemExit(f"Requested version {expected!r} does not match Cargo.toml {version!r}")
print(f"Building excelsior-fast {version}")
