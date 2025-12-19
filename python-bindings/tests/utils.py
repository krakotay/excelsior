import time
from contextlib import contextmanager

@contextmanager
def measure(label: str = ""):
    start = time.perf_counter()
    try:
        yield
    finally:
        end = time.perf_counter()
        delta = end - start
        if label:
            print(f"[{label}] {delta:.6f} s")
        else:
            print(f"{delta:.6f} s")
