"""schemas.py — load + validate JSON instruction docs (plan / decision / patch).

These schemas are the standardized v1 contract distilled from qual-supply
(W3, 2026-05-25). Any downstream project producing plan/decision/patch JSON
for the doctools block/caption/image group modules should pass through
`validate()` first.

Usage::

    from lib.schemas import validate, load_schema

    err = validate(data, "plan")
    if err:
        sys.exit(f"plan validation failed: {err}")

Falls back to minimal manual `required`/`const` checks when the `jsonschema`
library is not installed (lib/schemas.py keeps doctools stdlib-only by default).
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Optional

SCHEMA_DIR = Path(__file__).resolve().parent.parent / "schemas"

_SCHEMA_CACHE: dict[str, dict[str, Any]] = {}


def load_schema(name: str) -> dict[str, Any]:
    """Load schema by short name ('plan' / 'decision' / 'patch').

    Looks for ``<SCHEMA_DIR>/<name>.schema.json`` and caches the parsed dict.
    """
    if name in _SCHEMA_CACHE:
        return _SCHEMA_CACHE[name]
    path = SCHEMA_DIR / f"{name}.schema.json"
    if not path.exists():
        raise FileNotFoundError(f"schema not found: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    _SCHEMA_CACHE[name] = data
    return data


def _manual_validate(data: Any, schema: dict[str, Any], path: str = "$") -> Optional[str]:
    """Minimal fallback validator (no jsonschema lib).

    Covers: type=object/array/integer/string/const, required, properties recursion,
    array items. Enough to catch missing version/required-fields mismatches that
    cause downstream KeyError. Not a full draft-07 implementation.
    """
    t = schema.get("type")
    if "const" in schema:
        if data != schema["const"]:
            return f"{path}: expected const {schema['const']!r}, got {data!r}"
    if "enum" in schema:
        if data not in schema["enum"]:
            return f"{path}: expected one of {schema['enum']}, got {data!r}"
    if t == "object":
        if not isinstance(data, dict):
            return f"{path}: expected object, got {type(data).__name__}"
        for k in schema.get("required", []):
            if k not in data:
                return f"{path}: missing required field {k!r}"
        for k, sub in schema.get("properties", {}).items():
            if k in data:
                err = _manual_validate(data[k], sub, f"{path}.{k}")
                if err:
                    return err
    elif t == "array":
        if not isinstance(data, list):
            return f"{path}: expected array, got {type(data).__name__}"
        item_schema = schema.get("items")
        if isinstance(item_schema, dict):
            for i, it in enumerate(data):
                # oneOf: pass if any alternative validates
                if "oneOf" in item_schema:
                    matched = False
                    last_err = None
                    for alt in item_schema["oneOf"]:
                        e = _manual_validate(it, alt, f"{path}[{i}]")
                        if e is None:
                            matched = True
                            break
                        last_err = e
                    if not matched:
                        return f"{path}[{i}]: none of oneOf matched (last err: {last_err})"
                else:
                    err = _manual_validate(it, item_schema, f"{path}[{i}]")
                    if err:
                        return err
    elif t == "integer":
        if not isinstance(data, int) or isinstance(data, bool):
            return f"{path}: expected integer, got {type(data).__name__}"
    elif t == "string":
        if not isinstance(data, str):
            return f"{path}: expected string, got {type(data).__name__}"
    return None


def validate(data: Any, schema_name: str) -> Optional[str]:
    """Validate ``data`` against schema ``<schema_name>.schema.json``.

    Returns None on success, an error string on failure.
    Tries ``jsonschema`` library first; falls back to manual checker.
    """
    schema = load_schema(schema_name)
    try:
        import jsonschema  # type: ignore

        try:
            jsonschema.validate(data, schema)
            return None
        except jsonschema.ValidationError as e:  # type: ignore[attr-defined]
            return f"{'/'.join(str(x) for x in e.path) or '$'}: {e.message}"
        except Exception as e:
            return f"jsonschema error: {type(e).__name__}: {e}"
    except ImportError:
        return _manual_validate(data, schema)


def load_and_validate(json_path: str | Path, schema_name: str) -> tuple[Optional[dict], Optional[str]]:
    """Convenience: read JSON file from disk + validate. Returns (data, err)."""
    p = Path(json_path)
    if not p.exists():
        return None, f"file not found: {p}"
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        return None, f"json parse error: {e}"
    err = validate(data, schema_name)
    return (None, err) if err else (data, None)


__all__ = ["load_schema", "validate", "load_and_validate", "SCHEMA_DIR"]
