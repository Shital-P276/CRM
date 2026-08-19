import io
import os
import time
import zipfile
from functools import wraps

import openpyxl
from flask import Flask, jsonify, render_template, request, send_file
from flask_limiter import Limiter
from flask_limiter.errors import RateLimitExceeded
from flask_limiter.util import get_remote_address

import auth
import config
import data_layer
import security
import settings

limiter = Limiter(key_func=get_remote_address, default_limits=["60 per minute"])


@limiter.limit(lambda: config.LOGIN_RATE_LIMIT)
def api_login():
    data = request.get_json(silent=True) or {}
    password = str(data.get("password", ""))
    if not auth.password_configured():
        return jsonify({"error": "server not configured", "code": 503}), 503
    if auth.login(password):
        return jsonify({"authenticated": True, "csrf_token": security.get_csrf_token()})
    return jsonify({"error": "invalid password", "code": 401}), 401


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["SECRET_KEY"] = config.get_secret_key() or os.urandom(32).hex()
    app.config["SESSION_COOKIE_NAME"] = config.SESSION_COOKIE_NAME
    app.config["SESSION_COOKIE_HTTPONLY"] = True
    app.config["SESSION_COOKIE_SAMESITE"] = "Strict"
    app.config["SESSION_COOKIE_SECURE"] = bool(os.environ.get("COOKIE_SECURE", "0") == "1")
    app.config["MAX_CONTENT_LENGTH"] = config.MAX_UPLOAD_BYTES
    app.config["JSON_SORT_KEYS"] = False
    limiter.init_app(app)

    @app.before_request
    def csrf_guard():
        if request.method in ("GET", "HEAD", "OPTIONS"):
            return None
        if not session_csrf_token_present():
            return None
        if not security.validate_csrf():
            return jsonify({"error": "invalid csrf token", "code": 403}), 403
        return None

    def session_csrf_token_present() -> bool:
        from flask import session
        return bool(session.get("csrf_token"))

    @app.after_request
    def add_headers(response):
        return security.security_headers(response)

    @app.errorhandler(data_layer.SheetError)
    def handle_sheet_error(error):
        message = str(error)
        code = 404 if "not found" in message else 400
        return jsonify({"error": message, "code": code}), code

    @app.errorhandler(413)
    def handle_too_large(_):
        return jsonify({"error": "file too large", "code": 413}), 413

    @app.errorhandler(RateLimitExceeded)
    def handle_rate_limit(exc):
        retry_after = 60
        try:
            reset_at, _ = limiter.limiter.get_window_stats(
                exc.limit.limit,
                get_remote_address(),
                exc.limit.scope_for(request.endpoint, request.method),
            )
            retry_after = max(1, int(reset_at + 1 - time.time()))
        except Exception:
            pass
        return jsonify({
            "error": f"Too many login attempts. Try again in {retry_after} seconds.",
            "retry_after": retry_after,
            "code": 429,
        }), 429

    def login_required(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            if not auth.is_authenticated():
                return jsonify({"error": "unauthorized", "code": 401}), 401
            return func(*args, **kwargs)
        return wrapper

    def require_wb_sheet():
        wb = security.safe_workbook_name(request.args.get("wb", ""))
        sheet = security.safe_sheet_name(request.args.get("sheet", ""))
        if not wb:
            raise data_layer.SheetError("invalid workbook name")
        if not sheet:
            raise data_layer.SheetError("invalid sheet name")
        return wb, sheet

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.get("/api/health")
    def api_health():
        return jsonify({"status": "ok"})

    @app.get("/api/session")
    def api_session():
        return jsonify({
            "authenticated": auth.is_authenticated(),
            "csrf_token": security.get_csrf_token(),
            "password_configured": auth.password_configured(),
        })

    app.add_url_rule("/api/login", "api_login", api_login, methods=["POST"])

    @app.post("/api/logout")
    @login_required
    def api_logout():
        auth.logout()
        return jsonify({"logged_out": True})

    @app.get("/api/workbooks")
    @login_required
    def api_workbooks():
        return jsonify({"workbooks": data_layer.list_workbooks()})

    @app.get("/api/sheets")
    @login_required
    def api_sheets():
        wb = security.safe_workbook_name(request.args.get("wb", ""))
        if not wb:
            raise data_layer.SheetError("invalid workbook name")
        return jsonify({"sheets": data_layer.list_sheets(wb)})

    @app.get("/api/sheet-data")
    @login_required
    def api_sheet_data():
        wb, sheet = require_wb_sheet()
        loaded = data_layer.load_sheet(wb, sheet)
        rows = []
        for idx in range(loaded.row_count):
            values = {h: loaded.cell_value(h, idx) for h in loaded.headers}
            flagged = values.get(config.RESERVED_COLUMN, "").upper() == "TRUE"
            rows.append({"excel_row": loaded.excel_row_of(idx), "values": values, "flagged": flagged})
        return jsonify({
            "headers": loaded.headers,
            "rows": rows,
            "formula_cols": loaded.formula_cols,
            "numeric_cols": loaded.numeric_cols,
            "date_cols": loaded.date_cols,
            "has_flagged": loaded.has_flagged,
            "append_direction": settings.get_append_direction(wb),
            "warnings": loaded.warnings,
        })

    def _body_wb_sheet():
        data = request.get_json(silent=True) or {}
        wb = security.safe_workbook_name(data.get("wb", ""))
        sheet = security.safe_sheet_name(data.get("sheet", ""))
        if not wb:
            raise data_layer.SheetError("invalid workbook name")
        if not sheet:
            raise data_layer.SheetError("invalid sheet name")
        return wb, sheet

    def _string_values(raw) -> dict:
        return {str(k): str(v) for k, v in (raw or {}).items()}

    @app.post("/api/rows")
    @login_required
    def api_add_row():
        data = request.get_json(silent=True) or {}
        wb, sheet = _body_wb_sheet()
        with data_layer.file_lock(wb):
            loaded = data_layer.load_sheet(wb, sheet)
            try:
                result = data_layer.add_row(
                    loaded, _string_values(data.get("values")), force=bool(data.get("force"))
                )
            except data_layer.DuplicateError as error:
                return jsonify({"error": "duplicates", "duplicates": error.duplicates, "code": 409}), 409
            index = result["excel_row"] - loaded.data_start_row
            values = {h: loaded.cell_value(h, index) for h in loaded.headers}
            result["row"] = {
                "excel_row": result["excel_row"],
                "values": values,
                "flagged": values.get(config.RESERVED_COLUMN, "").upper() == "TRUE",
            }
            return jsonify(result)

    @app.put("/api/rows/<int:excel_row>")
    @login_required
    def api_update_row(excel_row):
        data = request.get_json(silent=True) or {}
        wb, sheet = _body_wb_sheet()
        with data_layer.file_lock(wb):
            loaded = data_layer.load_sheet(wb, sheet)
            try:
                result = data_layer.update_row(
                    loaded, excel_row, _string_values(data.get("values")), force=bool(data.get("force"))
                )
            except data_layer.DuplicateError as error:
                return jsonify({"error": "duplicates", "duplicates": error.duplicates, "code": 409}), 409
            return jsonify(result)

    @app.delete("/api/rows/<int:excel_row>")
    @login_required
    def api_delete_row(excel_row):
        data = request.get_json(silent=True) or {}
        wb, sheet = _body_wb_sheet()
        with data_layer.file_lock(wb):
            loaded = data_layer.load_sheet(wb, sheet)
            data_layer.delete_row(loaded, excel_row)
            return jsonify({"deleted": True, "excel_row": excel_row})

    @app.post("/api/rows/<int:excel_row>/flag")
    @login_required
    def api_flag_row(excel_row):
        data = request.get_json(silent=True) or {}
        wb, sheet = _body_wb_sheet()
        with data_layer.file_lock(wb):
            loaded = data_layer.load_sheet(wb, sheet)
            return jsonify(data_layer.toggle_flag(loaded, excel_row))

    @app.post("/api/sheets")
    @login_required
    def api_add_sheet():
        data = request.get_json(silent=True) or {}
        wb = security.safe_workbook_name(data.get("wb", ""))
        sheet = security.safe_sheet_name(data.get("name", ""))
        if not wb:
            raise data_layer.SheetError("invalid workbook name")
        if not sheet:
            raise data_layer.SheetError("invalid sheet name")
        with data_layer.file_lock(wb):
            data_layer.add_sheet(wb, sheet)
            return jsonify({"sheets": data_layer.list_sheets(wb)})

    @app.post("/api/upload")
    @login_required
    def api_upload():
        upload = request.files.get("file")
        if upload is None or not upload.filename:
            raise data_layer.SheetError("no file provided")
        name = os.path.basename(upload.filename)
        if not name.endswith(config.ALLOWED_EXTENSION):
            raise data_layer.SheetError("only .xlsx files are allowed")
        content = upload.read()
        if len(content) > config.MAX_UPLOAD_BYTES:
            raise data_layer.SheetError("file too large")
        if not zipfile.is_zipfile(io.BytesIO(content)):
            raise data_layer.SheetError("not a valid xlsx file")
        try:
            openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=False)
        except Exception as exc:
            raise data_layer.SheetError("not a valid xlsx file") from exc
        name = security.safe_workbook_name(name)
        if not name:
            raise data_layer.SheetError("invalid workbook name")
        with data_layer.file_lock(name):
            dest = config.DATA_DIR / name
            if dest.exists():
                data_layer.create_backup(name)
            tmp = dest.with_suffix(".xlsx.tmp")
            tmp.write_bytes(content)
            tmp.replace(dest)
        return jsonify({"workbooks": data_layer.list_workbooks(), "name": name})

    @app.get("/api/download")
    @login_required
    def api_download():
        wb = security.safe_workbook_name(request.args.get("wb", ""))
        if not wb:
            raise data_layer.SheetError("invalid workbook name")
        path = config.DATA_DIR / wb
        if not path.exists():
            raise data_layer.SheetError("workbook not found")
        return send_file(path, as_attachment=True, download_name=wb,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.get("/api/backups")
    @login_required
    def api_backups():
        wb = security.safe_workbook_name(request.args.get("wb", ""))
        if not wb:
            raise data_layer.SheetError("invalid workbook name")
        return jsonify({"backups": data_layer.list_backups(wb)})

    @app.get("/api/backups/<filename>/download")
    @login_required
    def api_backup_download(filename):
        path = data_layer.backup_path(filename)
        return send_file(path, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.get("/api/settings")
    @login_required
    def api_get_settings():
        wb = security.safe_workbook_name(request.args.get("wb", ""))
        if not wb:
            raise data_layer.SheetError("invalid workbook name")
        return jsonify({"append_direction": settings.get_append_direction(wb)})

    @app.put("/api/settings")
    @login_required
    def api_put_settings():
        data = request.get_json(silent=True) or {}
        wb = security.safe_workbook_name(data.get("wb", ""))
        if not wb:
            raise data_layer.SheetError("invalid workbook name")
        direction = str(data.get("append_direction", ""))
        settings.set_append_direction(wb, direction)
        return jsonify({"append_direction": settings.get_append_direction(wb)})

    return app


app = create_app()

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "5000")), debug=False)