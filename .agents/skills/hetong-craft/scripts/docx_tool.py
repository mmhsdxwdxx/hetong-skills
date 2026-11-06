#!/usr/bin/env python3
"""Compatibility entrypoint; implementation lives in hetong-word."""
from pathlib import Path as _Path
import sys as _sys
_target = _Path(__file__).resolve().parents[2] / "hetong-word" / "scripts" / _Path(__file__).name
if str(_target.parent) not in _sys.path:
    _sys.path.insert(0, str(_target.parent))
__file__ = str(_target)
exec(compile(_target.read_bytes(), str(_target), "exec"), globals())
