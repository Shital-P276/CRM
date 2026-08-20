import json
import threading

import config

_lock = threading.Lock()
_cache = None


def _path():
    return config.DATA_DIR / "settings.json"


def _load() -> dict:
    global _cache
    if _cache is not None:
        return _cache
    path = _path()
    if path.exists():
        try:
            _cache = json.loads(path.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            _cache = {}
    else:
        _cache = {}
    return _cache


def _save(data: dict) -> None:
    global _cache
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp = _path().with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(_path())
    _cache = data


def get_append_direction(workbook: str) -> str:
    with _lock:
        data = _load()
        value = (data.get(workbook) or {}).get("append_direction", config.DEFAULT_APPEND_DIRECTION)
    return "top" if value == "top" else "bottom"


def set_append_direction(workbook: str, direction: str) -> None:
    if direction not in ("top", "bottom"):
        raise ValueError(f"invalid append direction: {direction}")
    with _lock:
        data = _load()
        data.setdefault(workbook, {})["append_direction"] = direction
        _save(data)


def get_duplicate_columns(workbook: str, sheet: str) -> list:
    with _lock:
        data = _load()
        workbook_settings = data.get(workbook) or {}
        configured = (workbook_settings.get("duplicate_check_columns_by_sheet") or {}).get(sheet)
    return list(configured) if configured else []


def set_duplicate_columns(workbook: str, sheet: str, columns: list) -> None:
    if not isinstance(columns, list) or not all(isinstance(c, str) for c in columns):
        raise ValueError("duplicate_check_columns must be a list of column names")
    with _lock:
        data = _load()
        workbook_settings = data.setdefault(workbook, {})
        by_sheet = workbook_settings.setdefault("duplicate_check_columns_by_sheet", {})
        by_sheet[sheet] = list(columns)
        _save(data)


def all_settings() -> dict:
    with _lock:
        return json.loads(json.dumps(_load()))


_COLUMN_TYPES = ("text", "number", "date", "amount")


def get_column_types(workbook: str, sheet: str) -> dict:
    """Declared column types for a sheet, e.g. {"Amount": "number"}.

    Empty sheet columns have no data to infer a type from, so these
    declarations are the durable record used by load_sheet() as a
    fallback when inference can't classify a column."""
    with _lock:
        data = _load()
        workbook_settings = data.get(workbook) or {}
        configured = (workbook_settings.get("column_types_by_sheet") or {}).get(sheet)
    return dict(configured) if isinstance(configured, dict) else {}


def set_column_types(workbook: str, sheet: str, columns: dict) -> None:
    if not isinstance(columns, dict) or not all(
        isinstance(name, str) and col_type in _COLUMN_TYPES for name, col_type in columns.items()
    ):
        raise ValueError("column types must map column names to text|number|date|amount")
    with _lock:
        data = _load()
        workbook_settings = data.setdefault(workbook, {})
        by_sheet = workbook_settings.setdefault("column_types_by_sheet", {})
        by_sheet[sheet] = dict(columns)
        _save(data)


def set_column_type(workbook: str, sheet: str, name: str, col_type: str) -> None:
    if not isinstance(name, str) or not name or col_type not in _COLUMN_TYPES:
        raise ValueError(f"invalid column type: {col_type}")
    with _lock:
        data = _load()
        workbook_settings = data.setdefault(workbook, {})
        by_sheet = workbook_settings.setdefault("column_types_by_sheet", {})
        by_sheet.setdefault(sheet, {})[name] = col_type
        _save(data)


def remove_column_type(workbook: str, sheet: str, name: str) -> None:
    with _lock:
        data = _load()
        by_sheet = (data.get(workbook) or {}).get("column_types_by_sheet")
        sheet_map = (by_sheet or {}).get(sheet)
        if sheet_map and name in sheet_map:
            sheet_map.pop(name, None)
            _save(data)


def rename_column_type(workbook: str, sheet: str, old_name: str, new_name: str) -> None:
    with _lock:
        data = _load()
        by_sheet = (data.get(workbook) or {}).get("column_types_by_sheet")
        sheet_map = (by_sheet or {}).get(sheet)
        if sheet_map and old_name in sheet_map:
            sheet_map[new_name] = sheet_map.pop(old_name)
            _save(data)


_TOTALS_MODES = ("off", "all", "visible", "flagged", "visible_flagged")


def get_totals(workbook: str, sheet: str) -> dict:
    """Per-column totals mode for a sheet, e.g. {"Amount": "all"}.

    Modes: off (no total), all (every row), visible (rows matching the
    current search), flagged (rows with the flag set), visible_flagged
    (search-visible rows that also have the flag set)."""
    with _lock:
        data = _load()
        workbook_settings = data.get(workbook) or {}
        configured = (workbook_settings.get("totals_by_sheet") or {}).get(sheet)
    if not isinstance(configured, dict):
        return {}
    return {name: mode for name, mode in configured.items() if mode in _TOTALS_MODES}


def set_totals(workbook: str, sheet: str, modes: dict) -> None:
    if not isinstance(modes, dict) or not all(
        isinstance(name, str) and mode in _TOTALS_MODES for name, mode in modes.items()
    ):
        raise ValueError(f"totals modes must map column names to one of {_TOTALS_MODES}")
    with _lock:
        data = _load()
        workbook_settings = data.setdefault(workbook, {})
        by_sheet = workbook_settings.setdefault("totals_by_sheet", {})
        by_sheet[sheet] = dict(modes)
        _save(data)


def set_total(workbook: str, sheet: str, name: str, mode: str) -> None:
    if not isinstance(name, str) or not name or mode not in _TOTALS_MODES:
        raise ValueError(f"invalid totals mode: {mode}")
    with _lock:
        data = _load()
        workbook_settings = data.setdefault(workbook, {})
        by_sheet = workbook_settings.setdefault("totals_by_sheet", {})
        by_sheet.setdefault(sheet, {})[name] = mode
        _save(data)


def remove_total(workbook: str, sheet: str, name: str) -> None:
    with _lock:
        data = _load()
        by_sheet = (data.get(workbook) or {}).get("totals_by_sheet")
        sheet_map = (by_sheet or {}).get(sheet)
        if sheet_map and name in sheet_map:
            sheet_map.pop(name, None)
            _save(data)


def get_last_opened() -> dict:
    """Global 'last opened' workbook + sheet pointer (single-user app)."""
    with _lock:
        value = _load().get("last_opened")
    if not isinstance(value, dict):
        return {"workbook": None, "sheet": None}
    wb = value.get("workbook")
    sheet = value.get("sheet")
    return {
        "workbook": wb if isinstance(wb, str) else None,
        "sheet": sheet if isinstance(sheet, str) else None,
    }


def set_last_opened(workbook: str, sheet: str) -> None:
    if not isinstance(workbook, str) or not workbook:
        raise ValueError("invalid workbook")
    if not isinstance(sheet, str) or not sheet:
        raise ValueError("invalid sheet")
    with _lock:
        data = _load()
        data["last_opened"] = {"workbook": workbook, "sheet": sheet}
        _save(data)