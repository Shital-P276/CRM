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


def all_settings() -> dict:
    with _lock:
        return json.loads(json.dumps(_load()))