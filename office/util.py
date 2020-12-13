import sys
from contextlib import contextmanager
from io import StringIO


@contextmanager
def capture(stdout_curr=None, stderr_curr=None):
    stdout_prev, stderr_prev = sys.stdout, sys.stderr
    if stdout_curr is None:
        stdout_curr = StringIO()
    if stderr_curr is None:
        stderr_curr = StringIO()
    try:
        sys.stdout, sys.stderr = stdout_curr, stderr_curr
        yield stdout_curr, stderr_curr
    finally:
        sys.stdout, sys.stderr = stdout_prev, stderr_prev


def inch(value, reverse=False):
    if reverse:
        return value / 72
    return value * 72


def rgb(r, g, b):
    return r + (g * 256) + (b * 256 ** 2)
