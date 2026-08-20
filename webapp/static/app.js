// Customer Tracker — frontend SPA (vanilla JS, no build step)
"use strict";

const state = {
  csrfToken: null,
  workbook: null,
  sheet: null,
  headers: [],
  rows: [],
  formulaCols: new Set(),
  numericCols: new Set(),
  dateCols: new Set(),
  appendDirection: "bottom",
  sortCol: null,
  sortDir: 1,
  search: "",
  flaggedOnly: false,
  selected: new Set(),
  totals: {},
};

const el = (id) => document.getElementById(id);
const loginView = el("login-view");
const appView = el("app-view");
const tableHead = el("table-head");
const tableBody = el("table-body");
const tableFoot = el("table-foot");
const modalRoot = el("modal-root");
const toastRoot = el("toast-root");
const warningsBanner = el("warnings-banner");

// ---------------- API helper ----------------

async function api(path, options = {}) {
  const opts = {
    method: options.method || "GET",
    headers: { ...(options.headers || {}) },
  };
  if (options.body instanceof FormData) {
    opts.body = options.body;
  } else if (options.body !== undefined) {
    opts.headers["Content-Type"] = "application/json";
    opts.body = JSON.stringify(options.body);
  }
  if (opts.method !== "GET" && state.csrfToken) {
    opts.headers["X-CSRF-Token"] = state.csrfToken;
  }
  const res = await fetch(path, opts);

  if (res.status === 401) {
    showLogin();
    throw new ApiError("session expired", 401, null);
  }
  if (res.status === 429) {
    let data = {};
    try { data = await res.json(); } catch (_) {}
    throw new ApiError(data.error || "Too many attempts.", 429, data);
  }

  let data = null;
  try { data = await res.json(); } catch (_) {}

  if (!res.ok) {
    throw new ApiError((data && data.error) || `Request failed (${res.status})`, res.status, data);
  }
  return data;
}

class ApiError extends Error {
  constructor(message, status, data) {
    super(message);
    this.status = status;
    this.data = data;
  }
}

// ---------------- Toasts ----------------

function toast(message, kind = "info") {
  const node = document.createElement("div");
  node.className = `toast ${kind}`;
  node.textContent = message;
  toastRoot.appendChild(node);
  setTimeout(() => {
    node.classList.add("fade");
    setTimeout(() => node.remove(), 300);
  }, 3800);
}

// ---------------- View switching ----------------

function showLogin() {
  loginView.classList.remove("hidden");
  appView.classList.add("hidden");
  el("login-password").value = "";
  el("login-error").textContent = "";
}

function showApp() {
  loginView.classList.add("hidden");
  appView.classList.remove("hidden");
}

// ---------------- Auth ----------------

el("login-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const password = el("login-password").value;
  const btn = el("btn-login");
  const errNode = el("login-error");
  errNode.textContent = "";
  btn.disabled = true;
  try {
    const data = await api("/api/login", { method: "POST", body: { password } });
    state.csrfToken = data.csrf_token || null;
    showApp();
    await loadWorkbooks();
  } catch (err) {
    if (err.status === 429 && err.data && err.data.retry_after) {
      errNode.textContent = err.message;
      let remaining = err.data.retry_after;
      const timer = setInterval(() => {
        remaining -= 1;
        if (remaining <= 0) {
          clearInterval(timer);
          errNode.textContent = "You can try again now.";
        } else {
          errNode.textContent = `${err.message} (${remaining}s)`;
        }
      }, 1000);
    } else {
      errNode.textContent = err.message || "Invalid password.";
    }
  } finally {
    btn.disabled = false;
  }
});

el("btn-logout").addEventListener("click", async () => {
  try { await api("/api/logout", { method: "POST" }); } catch (_) {}
  state.csrfToken = null;
  showLogin();
});

async function checkSession() {
  try {
    const data = await api("/api/session");
    // /api/session mints a CSRF token into the session the moment it's
    // called, even before login — capture it now so the login POST (and
    // anything after) can echo it back, or the CSRF guard will 403 it.
    if (data.csrf_token) state.csrfToken = data.csrf_token;
    if (data.authenticated) {
      showApp();
      await loadWorkbooks();
      return;
    }
  } catch (_) {}
  showLogin();
}

// ---------------- Workbooks / sheets ----------------

async function loadWorkbooks() {
  const data = await api("/api/workbooks");
  const select = el("wb-select");
  select.innerHTML = "";
  (data.workbooks || []).forEach((wb) => {
    const opt = document.createElement("option");
    opt.value = wb.name;
    opt.textContent = wb.name;
    select.appendChild(opt);
  });
  if (data.workbooks && data.workbooks.length) {
    let remembered = null;
    try { remembered = await api("/api/last-opened"); } catch (_) {}
    const wbNames = data.workbooks.map((w) => w.name);
    const rememberedWb = remembered && wbNames.includes(remembered.workbook)
      ? remembered.workbook
      : null;
    state.workbook = rememberedWb || data.workbooks[0].name;
    select.value = state.workbook;
    await loadSheets(rememberedWb ? remembered.sheet : null);
  } else {
    renderEmptyTable("No workbooks found. Upload a .xlsx file to get started.");
  }
}

el("wb-select").addEventListener("change", async (e) => {
  state.workbook = e.target.value;
  await loadSheets();
  saveLastOpened();
});

async function loadSheets(preferredSheet = null) {
  if (!state.workbook) return;
  const data = await api(`/api/sheets?wb=${encodeURIComponent(state.workbook)}`);
  const select = el("sheet-select");
  select.innerHTML = "";
  const sheets = data.sheets || [];
  sheets.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    select.appendChild(opt);
  });
  if (sheets.length) {
    const target = preferredSheet && sheets.includes(preferredSheet)
      ? preferredSheet
      : sheets[0];
    state.sheet = target;
    select.value = target;
    await loadSettings();
    await loadSheetData();
  }
}

el("sheet-select").addEventListener("change", async (e) => {
  state.sheet = e.target.value;
  state.selected.clear();
  await loadSettings();
  await loadSheetData();
  saveLastOpened();
});

function saveLastOpened() {
  if (!state.workbook || !state.sheet) return;
  api("/api/last-opened", {
    method: "PUT",
    body: { workbook: state.workbook, sheet: state.sheet },
  }).catch(() => {});
}

el("btn-add-sheet").addEventListener("click", openNewSheetModal);

function sheetColRow() {
  const row = document.createElement("div");
  row.className = "sheet-col-row";
  row.innerHTML = `
    <input class="sheet-col-name" type="text" placeholder="Column name, e.g. AMOUNT">
    <select class="sheet-col-type" aria-label="Column type">
      <option value="text">Text</option>
      <option value="number">Number</option>
      <option value="date">Date</option>
    </select>
    <button type="button" class="btn btn-quiet sheet-col-remove" aria-label="Remove column">×</button>`;
  row.querySelector(".sheet-col-remove").addEventListener("click", () => row.remove());
  return row;
}

async function openNewSheetModal() {
  openModal(`
    <h2>Add sheet</h2>
    <div class="modal-field">
      <label for="new-sheet-name">Sheet name</label>
      <input id="new-sheet-name" type="text" placeholder="e.g. January 2026">
    </div>
    <div class="modal-field" style="margin-top:14px;">
      <label>Columns <span class="hint" style="display:inline; margin-left:6px;">optional — defaults to DATE / CUSTOMER NAME</span></label>
      <div id="new-sheet-cols"></div>
      <button type="button" class="btn btn-quiet add-sheet-col-btn" id="add-sheet-col">+ Add column</button>
    </div>
    <div class="modal-actions">
      <button class="btn" data-close>Cancel</button>
      <button class="btn btn-primary" id="modal-save-sheet">Create</button>
    </div>
  `);

  const colsRoot = el("new-sheet-cols");
  el("add-sheet-col").addEventListener("click", () => colsRoot.appendChild(sheetColRow()));
  colsRoot.appendChild(sheetColRow());

  const saveBtn = el("modal-save-sheet");
  saveBtn.addEventListener("click", async () => {
    if (saveBtn.disabled) return;
    const name = el("new-sheet-name").value.trim();
    if (!name) {
      toast("Sheet name is required.", "error");
      return;
    }
    const columns = [];
    colsRoot.querySelectorAll(".sheet-col-row").forEach((row) => {
      const colName = row.querySelector(".sheet-col-name").value.trim();
      if (!colName) return;
      columns.push({ name: colName, type: row.querySelector(".sheet-col-type").value });
    });
    saveBtn.disabled = true;
    try {
      const body = { wb: state.workbook, name };
      if (columns.length) body.columns = columns;
      await api("/api/sheets", { method: "POST", body });
      closeModal();
      toast(`Sheet "${name}" added.`, "success");
      await loadSheets();
    } catch (err) {
      toast(err.message, "error");
      saveBtn.disabled = false;
    }
  });
}

// ---------------- Settings (append direction) ----------------

async function loadSettings() {
  try {
    const data = await api(`/api/settings?wb=${encodeURIComponent(state.workbook)}`);
    state.appendDirection = data.append_direction || "bottom";
  } catch (_) {
    state.appendDirection = "bottom";
  }
}

// ---------------- Sheet data ----------------

async function loadSheetData() {
  if (!state.workbook || !state.sheet) return;
  const data = await api(
    `/api/sheet-data?wb=${encodeURIComponent(state.workbook)}&sheet=${encodeURIComponent(state.sheet)}`
  );
  state.headers = data.headers || [];
  state.rows = data.rows || [];
  state.formulaCols = new Set(data.formula_cols || []);
  state.numericCols = new Set(data.numeric_cols || []);
  state.dateCols = new Set(data.date_cols || []);
  state.amountCols = new Set(data.amount_cols || []);
  state.appendDirection = data.append_direction || state.appendDirection;
  state.totals = data.totals || {};
  state.selected.clear();
  renderWarnings(data.warnings || []);
  renderTable();
}

function renderWarnings(warnings) {
  if (!warnings.length) {
    warningsBanner.classList.add("hidden");
    warningsBanner.innerHTML = "";
    return;
  }
  warningsBanner.classList.remove("hidden");
  warningsBanner.innerHTML =
    `<strong>${warnings.length} row${warnings.length > 1 ? "s" : ""} need attention:</strong>` +
    `<ul>${warnings.map((w) => `<li>${escapeHtml(w.column || "")}: ${escapeHtml(w.reason || "could not calculate")}</li>`).join("")}</ul>`;
}

// ---------------- Table rendering ----------------

function formatDate(value) {
  if (!value) return "";
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(value));
  if (!m) return value;
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const month = months[Number(m[2]) - 1];
  if (!month) return value;
  return `${month} ${String(Number(m[3])).padStart(2, "0")}`;
}

function formatINR(value) {
  if (value === "" || value == null) return "";
  const n = Number(value);
  if (!Number.isFinite(n)) return String(value);
  const [int, frac] = String(Math.abs(n)).split(".");
  const last3 = int.slice(-3);
  let rest = int.slice(0, -3);
  let grouped = last3;
  if (rest) {
    const parts = [];
    while (rest.length > 2) {
      parts.unshift(rest.slice(-2));
      rest = rest.slice(0, -2);
    }
    parts.unshift(rest);
    grouped = parts.join(",") + "," + last3;
  }
  return (n < 0 ? "-" : "") + grouped + (frac !== undefined ? "." + frac : "");
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function getFilteredSortedRows() {
  let rows = state.rows.slice();
  if (state.flaggedOnly) rows = rows.filter((r) => r.flagged);
  if (state.search) {
    const q = state.search.toLowerCase();
    rows = rows.filter((r) =>
      state.headers.some((h) => String(r.values[h] ?? "").toLowerCase().includes(q))
    );
  }
  if (state.sortCol) {
    const col = state.sortCol;
    const isNum = state.numericCols.has(col);
    const isDate = state.dateCols.has(col);
    rows.sort((a, b) => {
      let va = a.values[col] ?? "";
      let vb = b.values[col] ?? "";
      if (isNum) {
        va = parseFloat(va) || 0;
        vb = parseFloat(vb) || 0;
      } else if (isDate) {
        va = new Date(va).getTime() || 0;
        vb = new Date(vb).getTime() || 0;
      } else {
        va = String(va).toLowerCase();
        vb = String(vb).toLowerCase();
      }
      if (va < vb) return -1 * state.sortDir;
      if (va > vb) return 1 * state.sortDir;
      return 0;
    });
  }
  return rows;
}

function renderTable() {
  if (!state.headers.length) {
    renderEmptyTable("This sheet has no columns yet.");
    return;
  }

  tableHead.innerHTML = "";
  const trh = document.createElement("tr");

  const thSelect = document.createElement("th");
  thSelect.className = "col-select";
  const selectAll = document.createElement("input");
  selectAll.type = "checkbox";
  selectAll.className = "row-check";
  selectAll.setAttribute("aria-label", "Select all rows");
  selectAll.addEventListener("change", (e) => {
    const rows = getFilteredSortedRows();
    if (e.target.checked) rows.forEach((r) => state.selected.add(r.excel_row));
    else state.selected.clear();
    renderTableBody();
  });
  thSelect.appendChild(selectAll);
  trh.appendChild(thSelect);

  const thSerial = document.createElement("th");
  thSerial.className = "col-serial";
  thSerial.textContent = "#";
  trh.appendChild(thSerial);

  state.headers.forEach((h) => {
    if (h === "FLAGGED") return;
    const th = document.createElement("th");
    const isNum = state.numericCols.has(h);
    if (isNum) th.classList.add("numeric");
    if (state.sortCol === h) th.classList.add("sorted");
    th.innerHTML =
      escapeHtml(h) +
      (state.formulaCols.has(h) ? '<span class="warn-badge" title="Calculated field">fx</span>' : "") +
      (state.sortCol === h ? `<span class="sort-arrow">${state.sortDir === 1 ? "▲" : "▼"}</span>` : "");
    th.addEventListener("click", () => {
      if (state.sortCol === h) state.sortDir *= -1;
      else { state.sortCol = h; state.sortDir = 1; }
      renderTable();
    });
    trh.appendChild(th);
  });

  const thFlag = document.createElement("th");
  thFlag.className = "col-flag";
  thFlag.innerHTML = '<span aria-hidden="true">&#9873;</span><span class="sr-only">Flag</span>';
  trh.appendChild(thFlag);

  tableHead.appendChild(trh);
  renderTableBody();
}

function renderTableFoot(visible) {
  tableFoot.innerHTML = "";
  const activeCols = state.headers.filter(
    (h) => h !== "FLAGGED" && state.totals[h] && state.totals[h] !== "off"
  );
  if (!activeCols.length) return;

  if (!visible) visible = getFilteredSortedRows();
  const pool = (mode) => {
    if (mode === "all") return state.rows;
    if (mode === "flagged") return state.rows.filter((r) => r.flagged);
    if (mode === "visible") return visible;
    if (mode === "visible_flagged") return visible.filter((r) => r.flagged);
    return [];
  };

  const tr = document.createElement("tr");
  tr.className = "totals-row";
  const tdSelect = document.createElement("td");
  tdSelect.className = "col-select";
  tr.appendChild(tdSelect);
  const tdSerial = document.createElement("td");
  tdSerial.className = "col-serial";
  tdSerial.textContent = "Totals";
  tr.appendChild(tdSerial);

  state.headers.forEach((h) => {
    if (h === "FLAGGED") return;
    const td = document.createElement("td");
    const mode = state.totals[h];
    if (mode && mode !== "off") {
      const rows = pool(mode);
      if (state.numericCols.has(h)) {
        const nums = rows.map((r) => Number(r.values[h])).filter((n) => Number.isFinite(n));
        if (nums.length) {
          td.classList.add("numeric");
          const sum = nums.reduce((a, b) => a + b, 0);
          td.textContent = state.amountCols.has(h) ? formatINR(sum) : sum.toLocaleString();
        }
      } else if (rows.length) {
        td.textContent = String(rows.length);
      }
    }
    tr.appendChild(td);
  });

  const tdFlag = document.createElement("td");
  tdFlag.className = "col-flag";
  tr.appendChild(tdFlag);
  tableFoot.appendChild(tr);
}

function renderTableBody() {
  const rows = getFilteredSortedRows();
  tableBody.innerHTML = "";

  if (!rows.length) {
    renderEmptyBody(
      state.search || state.flaggedOnly ? "No matching records." : "No records yet — add your first one."
    );
    renderTableFoot(rows);
    return;
  }

  rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    if (row.flagged) tr.classList.add("flagged");
    if (state.selected.has(row.excel_row)) tr.classList.add("selected");

    const tdSelect = document.createElement("td");
    tdSelect.className = "col-select";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "row-check";
    cb.checked = state.selected.has(row.excel_row);
    cb.addEventListener("change", () => {
      if (cb.checked) state.selected.add(row.excel_row);
      else state.selected.delete(row.excel_row);
      tr.classList.toggle("selected", cb.checked);
    });
    tdSelect.appendChild(cb);
    tr.appendChild(tdSelect);

    const tdSerial = document.createElement("td");
    tdSerial.className = "col-serial";
    tdSerial.textContent = String(rowIndex + 1);
    tr.appendChild(tdSerial);

    state.headers.forEach((h) => {
      if (h === "FLAGGED") return;
      const td = document.createElement("td");
      const isFormula = state.formulaCols.has(h);
      const isNum = state.numericCols.has(h);
      const isDate = state.dateCols.has(h);
      const isAmount = state.amountCols.has(h);
      if (isNum) td.classList.add("numeric");
      if (isFormula) {
        td.classList.add("formula");
        td.textContent = isAmount ? formatINR(row.values[h]) : (row.values[h] ?? "");
      } else {
        td.classList.add("editable");
        td.textContent = isDate ? formatDate(row.values[h]) : isAmount ? formatINR(row.values[h]) : (row.values[h] ?? "");
        td.title = "Click to edit";
        td.addEventListener("click", () => editCell(td, row, h, isDate));
      }
      tr.appendChild(td);
    });

    const tdFlag = document.createElement("td");
    tdFlag.className = "col-flag";
    const flagBtn = document.createElement("button");
    flagBtn.className = "flag-toggle" + (row.flagged ? " flagged" : "");
    flagBtn.innerHTML = '<span aria-hidden="true">&#9873;</span>';
    flagBtn.setAttribute("aria-label", row.flagged ? "Unflag this row" : "Flag this row");
    flagBtn.addEventListener("click", () => toggleFlag(row));
    tdFlag.appendChild(flagBtn);
    tr.appendChild(tdFlag);

    tableBody.appendChild(tr);
  });

  renderTableFoot(rows);
}

function renderEmptyBody(message) {
  tableBody.innerHTML = "";
  const tr = document.createElement("tr");
  const td = document.createElement("td");
  td.className = "empty-state";
  td.colSpan = state.headers.filter((h) => h !== "FLAGGED").length + 2;
  td.innerHTML = `<div class="empty-title">${escapeHtml(message)}</div>`;
  tr.appendChild(td);
  tableBody.appendChild(tr);
}

function renderEmptyTable(message) {
  tableHead.innerHTML = "";
  renderEmptyBody(message);
  renderTableFoot([]);
}

// ---------------- Inline cell edit ----------------

function editCell(td, row, header, isDate) {
  if (td.querySelector("input")) return;
  const original = row.values[header] ?? "";
  const input = document.createElement("input");
  input.className = "cell-input";
  input.type = isDate ? "date" : "text";
  input.value = original;
  td.textContent = "";
  td.appendChild(input);
  input.focus();
  input.select();

  let done = false;
  const commit = async () => {
    if (done) return;
    done = true;
    const newValue = input.value;
    if (newValue === original) {
      td.textContent = isDate ? formatDate(original) : original;
      return;
    }
    try {
      const values = { ...row.values, [header]: newValue };
      await saveRow(row.excel_row, values, false);
      // Reload rather than merge the response in-place: the exact shape
      // of PUT /api/rows/<id>'s response isn't guaranteed to mirror
      // {values}, and a formula column may depend on this edit, so a
      // fresh load is the only way to be sure computed cells are correct.
      await loadSheetData();
      toast("Saved.", "success");
    } catch (err) {
      if (err.status === 409) {
        await handleDuplicate(err.data, row.excel_row, { ...row.values, [header]: newValue });
      } else {
        toast(err.message, "error");
        renderTableBody();
      }
    }
  };
  input.addEventListener("blur", commit);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") input.blur();
    if (e.key === "Escape") { done = true; renderTableBody(); }
  });
}

async function saveRow(excelRow, values, force) {
  return api(`/api/rows/${excelRow}`, {
    method: "PUT",
    body: { wb: state.workbook, sheet: state.sheet, values, force },
  });
}

async function handleDuplicate(data, excelRow, values) {
  const dupes = (data && data.duplicates) || [];
  const proceed = await confirmDialog(
    "Possible duplicate",
    `This looks similar to ${dupes.length} existing record${dupes.length > 1 ? "s" : ""}. Save anyway?`,
    dupes.map((d) => `Row ${d.excel_row}: ${Object.values(d.values || {}).slice(0, 3).join(" · ")}`)
  );
  if (!proceed) { await loadSheetData(); return; }
  try {
    if (excelRow) await saveRow(excelRow, values, true);
    else await api("/api/rows", { method: "POST", body: { wb: state.workbook, sheet: state.sheet, values, force: true } });
    await loadSheetData();
    toast("Saved.", "success");
  } catch (err) {
    toast(err.message, "error");
  }
}

// ---------------- Flag ----------------

async function toggleFlag(row) {
  try {
    const data = await api(`/api/rows/${row.excel_row}/flag`, {
      method: "POST",
      body: { wb: state.workbook, sheet: state.sheet },
    });
    row.flagged = data.flagged;
    renderTableBody();
    if (data.flagged_column_added) {
      toast('A "Flagged" column was added to this sheet to remember flags.', "warn");
    }
    return true;
  } catch (err) {
    toast(err.message, "error");
    return false;
  }
}

el("btn-flag").addEventListener("click", async () => {
  if (!state.selected.size) { toast("Select one or more rows first.", "warn"); return; }
  const rowsByExcel = new Map(state.rows.map((r) => [r.excel_row, r]));
  let allOk = true;
  for (const excelRow of state.selected) {
    const row = rowsByExcel.get(excelRow);
    if (row) {
      if (!(await toggleFlag(row))) allOk = false;
    }
  }
  if (allOk) {
    state.selected.clear();
    renderTableBody();
  }
});

// ---------------- Delete ----------------

el("btn-delete").addEventListener("click", async () => {
  if (!state.selected.size) { toast("Select one or more rows first.", "warn"); return; }
  const count = state.selected.size;
  const ok = await confirmDialog(
    "Delete records",
    `Delete ${count} selected record${count > 1 ? "s" : ""}? This can't be undone from here — a backup is kept automatically.`
  );
  if (!ok) return;
  try {
    for (const excelRow of Array.from(state.selected)) {
      await api(`/api/rows/${excelRow}`, {
        method: "DELETE",
        body: { wb: state.workbook, sheet: state.sheet },
      });
    }
    toast(`${count} record${count > 1 ? "s" : ""} deleted.`, "success");
    state.selected.clear();
    await loadSheetData();
  } catch (err) {
    toast(err.message, "error");
  }
});

// ---------------- Search / filter ----------------

let searchTimer = null;
el("search-input").addEventListener("input", (e) => {
  clearTimeout(searchTimer);
  const value = e.target.value;
  searchTimer = setTimeout(() => {
    state.search = value;
    renderTable();
  }, 300);
});

el("flagged-only").addEventListener("change", (e) => {
  state.flaggedOnly = e.target.checked;
  renderTable();
});

// ---------------- Add record modal ----------------

el("btn-add").addEventListener("click", openAddModal);

function openAddModal() {
  const fields = state.headers.filter((h) => h !== "FLAGGED");
  const formHtml = fields.map((h) => {
    const isFormula = state.formulaCols.has(h);
    const isDate = state.dateCols.has(h);
    const isNum = state.numericCols.has(h);
    const type = isDate ? "date" : isNum ? "number" : "text";
    const defaultVal = isNum ? "0" : "";
    return `
      <div class="modal-field">
        <label for="add-${cssId(h)}">${escapeHtml(h)}${isFormula ? " (calculated)" : ""}</label>
        <input id="add-${cssId(h)}" type="${type}" data-col="${escapeHtml(h)}"
               value="${escapeHtml(defaultVal)}" ${isFormula ? "disabled" : ""} ${isNum ? "step=\"any\"" : ""}>
      </div>`;
  }).join("");

  openModal(`
    <h2>Add record</h2>
    <div class="modal-grid">${formHtml}</div>
    <div class="modal-field" style="margin-top:16px;">
      <label for="add-append">New records go to</label>
      <select id="add-append" aria-label="Append direction">
        <option value="bottom"${state.appendDirection === "bottom" ? " selected" : ""}>Bottom</option>
        <option value="top"${state.appendDirection === "top" ? " selected" : ""}>Top</option>
      </select>
    </div>
    <div class="modal-actions">
      <button class="btn" data-close>Cancel</button>
      <button class="btn btn-primary" id="modal-save">Save</button>
    </div>
  `);

  const addSaveBtn = el("modal-save");
  addSaveBtn.addEventListener("click", async () => {
    if (addSaveBtn.disabled) return;
    addSaveBtn.disabled = true;
    const inputs = modalRoot.querySelectorAll("[data-col]");
    const values = {};
    inputs.forEach((input) => { values[input.dataset.col] = input.value; });
    const appendDir = el("add-append").value;
    try {
      if (appendDir !== state.appendDirection) {
        await api(`/api/settings?wb=${encodeURIComponent(state.workbook)}&sheet=${encodeURIComponent(state.sheet)}`, {
          method: "PUT",
          body: { append_direction: appendDir },
        });
        state.appendDirection = appendDir;
      }
      await api("/api/rows", { method: "POST", body: { wb: state.workbook, sheet: state.sheet, values } });
      closeModal();
      toast("Record added.", "success");
      await loadSheetData();
    } catch (err) {
      if (err.status === 409) {
        closeModal();
        await handleDuplicate(err.data, null, values);
      } else {
        toast(err.message, "error");
        addSaveBtn.disabled = false;
      }
    }
  });
}

function cssId(str) {
  return String(str).toLowerCase().replace(/[^a-z0-9]+/g, "-");
}

// ---------------- Changes (recent save points + raw backups) ----------------

el("btn-changes").addEventListener("click", openChangesModal);

async function openChangesModal() {
  try {
    const data = await api(`/api/backups?wb=${encodeURIComponent(state.workbook)}`);
    const backups = data.backups || [];
    const recentHtml = backups.length
      ? `<ul class="backup-list">${backups.map((b, i) => `
          <li class="change-row">
            <div class="change-info">
              <div class="change-title">${escapeHtml(b.description || `Version from ${b.created_at}`)}</div>
              <div class="backup-meta">${escapeHtml(b.created_at || "unknown time")}${i === 0 ? " · newest" : ""}</div>
            </div>
            <button type="button" class="btn btn-quiet change-revert" data-filename="${encodeURIComponent(b.filename)}">Revert to this version</button>
          </li>`).join("")}</ul>`
      : `<p class="modal-sub">No changes yet — a snapshot is saved automatically every time you make a change.</p>`;
    const rawHtml = backups.length
      ? `<ul class="backup-list">${backups.map((b) => `
          <li>
            <div>
              <div>${escapeHtml(b.filename)}</div>
              <div class="backup-meta">${escapeHtml(b.created_at || "")} · ${Math.round((b.size_bytes || 0) / 1024)} KB</div>
            </div>
            <a class="btn btn-quiet" href="/api/backups/${encodeURIComponent(b.filename)}/download">Download</a>
          </li>`).join("")}</ul>`
      : `<p class="modal-sub">No backup files yet.</p>`;
    openModal(`
      <h2>Changes — ${escapeHtml(state.sheet)}</h2>
      <p class="modal-sub">Recent save points for this workbook. Reverting restores an earlier version — the current version is saved as a new snapshot first.</p>
      ${recentHtml}

      <details class="format-section">
        <summary>Raw backup files</summary>
        <p class="hint">The underlying .bak files, one per save point. Download them directly.</p>
        ${rawHtml}
      </details>

      <div class="modal-actions"><button class="btn" data-close>Close</button></div>
    `);

    modalRoot.querySelectorAll(".change-revert").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const filename = decodeURIComponent(btn.dataset.filename);
        const label = btn.closest("li").querySelector(".change-title").textContent;
        const ok = await confirmDialog(
          "Revert to this version",
          `Restore this workbook to "${label}"? The current version is saved as a new snapshot first, so this can be undone.`
        );
        if (!ok) return;
        btn.disabled = true;
        try {
          await api(`/api/backups/${encodeURIComponent(filename)}/revert`, {
            method: "POST",
            body: { wb: state.workbook },
          });
          closeModal();
          toast("Restored to an earlier version.", "success");
          await loadSheets();
        } catch (err) {
          toast(err.message, "error");
          btn.disabled = false;
        }
      });
    });
  } catch (err) {
    toast(err.message, "error");
  }
}

// ---------------- Download / upload ----------------

el("btn-download").addEventListener("click", () => {
  window.location.href = `/api/download?wb=${encodeURIComponent(state.workbook)}`;
});

el("btn-upload").addEventListener("click", () => {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".xlsx";
  input.addEventListener("change", async () => {
    if (!input.files.length) return;
    const formData = new FormData();
    formData.append("file", input.files[0]);
    try {
      await api("/api/upload", { method: "POST", body: formData });
      toast("Workbook uploaded.", "success");
      await loadWorkbooks();
    } catch (err) {
      toast(err.message, "error");
    }
  });
  input.click();
});

// ---------------- Format panel (formulas + duplicate-check columns) ----------------

el("btn-formulas").addEventListener("click", openFormatModal);

async function openFormatModal() {
  let formulaData, settingsData;
  try {
    [formulaData, settingsData] = await Promise.all([
      api(`/api/formulas?wb=${encodeURIComponent(state.workbook)}&sheet=${encodeURIComponent(state.sheet)}`),
      api(`/api/settings?wb=${encodeURIComponent(state.workbook)}&sheet=${encodeURIComponent(state.sheet)}`),
    ]);
  } catch (err) {
    toast(err.message, "error");
    return;
  }

  const formulas = formulaData.formulas || {}; // {header: {formula, ref}}
  const dupColumns = new Set(settingsData.duplicate_check_columns || []);
  const columnHeaders = state.headers.filter((h) => h !== "FLAGGED");

  const formulaRows = columnHeaders.map((h) => {
    const entry = formulas[h] || { formula: "", ref: "" };
    return `
      <div class="modal-field">
        <label for="formula-${cssId(h)}">${escapeHtml(h)}${entry.ref ? `<span class="ref-badge">${escapeHtml(entry.ref)}</span>` : ""}</label>
        <input id="formula-${cssId(h)}" type="text" data-formula-col="${escapeHtml(h)}"
               value="${escapeHtml(entry.formula || "")}" placeholder="e.g. =D2*1.18">
      </div>`;
  }).join("");

  const dupRows = columnHeaders.map((h) => `
    <label>
      <input type="checkbox" data-dup-col="${escapeHtml(h)}" ${dupColumns.has(h) ? "checked" : ""}>
      <span>${escapeHtml(h)}</span>
    </label>`).join("");

  const totalsRows = columnHeaders.map((h) => {
    const mode = (settingsData.totals || {})[h] || "off";
    return `
      <label class="totals-col-row">
        <span class="totals-col-name">${escapeHtml(h)}</span>
        <select data-total-col="${escapeHtml(h)}" aria-label="Totals for ${escapeHtml(h)}">
          <option value="off" ${mode === "off" ? "selected" : ""}>Off</option>
          <option value="all" ${mode === "all" ? "selected" : ""}>All rows</option>
          <option value="visible" ${mode === "visible" ? "selected" : ""}>Visible</option>
          <option value="flagged" ${mode === "flagged" ? "selected" : ""}>Flagged</option>
          <option value="visible_flagged" ${mode === "visible_flagged" ? "selected" : ""}>Visible + Flagged</option>
        </select>
      </label>`;
  }).join("");

  const declaredType = (h) =>
    state.amountCols.has(h) ? "amount"
      : state.dateCols.has(h) ? "date"
      : state.numericCols.has(h) ? "number"
      : "text";

  const typeOptions = (current) =>
    ["text", "number", "date", "amount"]
      .map((t) => `<option value="${t}"${t === current ? " selected" : ""}>${t[0].toUpperCase()}${t.slice(1)}</option>`)
      .join("");

  const manageRows = columnHeaders.map((h) => `
    <div class="manage-col-row" data-col="${escapeHtml(h)}">
      <span class="manage-col-name">${escapeHtml(h)}</span>
      <select class="manage-col-type" aria-label="Type of ${escapeHtml(h)}">${typeOptions(declaredType(h))}</select>
      <button type="button" class="btn btn-quiet manage-rename">Rename</button>
      <button type="button" class="btn btn-quiet manage-delete">Delete</button>
    </div>`).join("");

  openModal(`
    <h2>Format — ${escapeHtml(state.sheet)}</h2>

    <details class="format-section">
      <summary>Formulas</summary>
      <p class="modal-sub">Formulas apply to every new row added to this sheet. Column letters are shown next to each name — use them like =D2*1.18.</p>
      <div class="modal-grid">${formulaRows}</div>
      <div class="modal-field" style="margin-top:16px;">
        <label>Add a new column</label>
        <div class="inline-pair">
          <input id="format-new-col" type="text" placeholder="Column name, e.g. GST TOTAL">
          <input id="format-new-formula" type="text" placeholder="Formula, e.g. =D2*1.18">
        </div>
      </div>
    </details>

    <details class="format-section">
      <summary>Advanced: duplicate detection columns</summary>
      <p class="hint">Leave all unchecked to use automatic detection based on common field names. Check specific columns to only flag duplicates on those.</p>
      <div class="dup-col-list">${dupRows}</div>
    </details>

    <details class="format-section">
      <summary>Column totals</summary>
      <p class="hint">Numeric columns show a sum, other columns show a count. "Visible" respects the current search and flag filter; "Flagged" only counts flagged rows.</p>
      <div class="totals-col-list">${totalsRows}</div>
    </details>

    <details class="format-section">
      <summary>Manage columns</summary>
      <p class="hint">Add, rename or delete columns. Deleting a column removes its data from every row — a backup is created first.</p>
      ${manageRows}
      <div class="manage-add">
        <input id="manage-new-col" type="text" placeholder="New column name">
        <select id="manage-new-type" aria-label="Column type">
          <option value="text">Text</option>
          <option value="number">Number</option>
          <option value="date">Date</option>
          <option value="amount">Amount</option>
        </select>
        <button type="button" class="btn" id="manage-add-col">Add</button>
      </div>
    </details>

    <div class="modal-actions">
      <button class="btn" data-close>Cancel</button>
      <button class="btn btn-primary" id="modal-save-format">Save</button>
    </div>
  `);

  modalRoot.querySelectorAll(".manage-delete").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const col = btn.closest(".manage-col-row").dataset.col;
      const ok = await confirmDialog(
        "Delete column",
        `Delete "${col}" and its data from every row? A backup is created first so you can restore it.`
      );
      if (!ok) return;
      try {
        await api("/api/columns", {
          method: "DELETE",
          body: { wb: state.workbook, sheet: state.sheet, name: col },
        });
        toast(`Column "${col}" deleted.`, "success");
        await loadSheetData();
        openFormatModal();
      } catch (err) {
        toast(err.message, "error");
      }
    });
  });

  modalRoot.querySelectorAll(".manage-col-type").forEach((sel) => {
    const original = sel.value;
    sel.addEventListener("change", async () => {
      const col = sel.closest(".manage-col-row").dataset.col;
      const nextType = sel.value;
      if (nextType === original) return;
      sel.disabled = true;
      try {
        await api("/api/columns/type", {
          method: "PUT",
          body: { wb: state.workbook, sheet: state.sheet, name: col, type: nextType },
        });
        toast(`Column "${col}" is now ${nextType[0].toUpperCase()}${nextType.slice(1)}.`, "success");
        await loadSheetData();
        openFormatModal();
      } catch (err) {
        toast(err.message, "error");
        sel.value = original;
        sel.disabled = false;
      }
    });
  });

  modalRoot.querySelectorAll(".manage-rename").forEach((btn) => {
    btn.addEventListener("click", () => {
      const row = btn.closest(".manage-col-row");
      if (row.dataset.editing) return;
      const col = row.dataset.col;
      row.dataset.editing = "1";
      row.innerHTML = `
        <input type="text" class="manage-rename-input" value="${escapeHtml(col)}">
        <button type="button" class="btn btn-primary manage-save">Save</button>
        <button type="button" class="btn btn-quiet manage-cancel">Cancel</button>`;
      const input = row.querySelector(".manage-rename-input");
      const saveBtn = row.querySelector(".manage-save");
      const cancelBtn = row.querySelector(".manage-cancel");
      input.focus();
      input.select();
      cancelBtn.addEventListener("click", () => openFormatModal());
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") saveBtn.click();
        else if (e.key === "Escape") openFormatModal();
      });
      saveBtn.addEventListener("click", async () => {
        if (saveBtn.disabled) return;
        const newName = input.value.trim();
        if (!newName || newName === col) { openFormatModal(); return; }
        saveBtn.disabled = true;
        try {
          await api("/api/columns", {
            method: "PUT",
            body: { wb: state.workbook, sheet: state.sheet, name: col, new_name: newName },
          });
          toast(`Column renamed to "${newName}".`, "success");
          await loadSheetData();
          openFormatModal();
        } catch (err) {
          toast(err.message, "error");
          saveBtn.disabled = false;
        }
      });
    });
  });

  el("manage-add-col").addEventListener("click", async () => {
    const addBtn = el("manage-add-col");
    const name = el("manage-new-col").value.trim();
    if (!name) {
      toast("Column name is required.", "error");
      return;
    }
    addBtn.disabled = true;
    try {
      await api("/api/columns", {
        method: "POST",
        body: { wb: state.workbook, sheet: state.sheet, name, type: el("manage-new-type").value },
      });
      toast(`Column "${name}" added.`, "success");
      await loadSheetData();
      openFormatModal();
    } catch (err) {
      toast(err.message, "error");
      addBtn.disabled = false;
    }
  });

  const formatSaveBtn = el("modal-save-format");
  formatSaveBtn.addEventListener("click", async () => {
    if (formatSaveBtn.disabled) return;
    formatSaveBtn.disabled = true;
    const formulaUpdates = {};
    modalRoot.querySelectorAll("[data-formula-col]").forEach((input) => {
      formulaUpdates[input.dataset.formulaCol] = input.value;
    });
    const newCol = el("format-new-col");
    const newFormula = el("format-new-formula");
    if (newCol && newFormula && newCol.value.trim() && newFormula.value.trim()) {
      formulaUpdates[newCol.value.trim()] = newFormula.value.trim();
    }

    const dupSelected = [];
    modalRoot.querySelectorAll("[data-dup-col]").forEach((cb) => {
      if (cb.checked) dupSelected.push(cb.dataset.dupCol);
    });

    const totals = {};
    modalRoot.querySelectorAll("[data-total-col]").forEach((sel) => {
      if (sel.value !== "off") totals[sel.dataset.totalCol] = sel.value;
    });

    try {
      await api(`/api/formulas?wb=${encodeURIComponent(state.workbook)}&sheet=${encodeURIComponent(state.sheet)}`, {
        method: "PUT",
        body: { formulas: formulaUpdates },
      });
      await api(`/api/settings?wb=${encodeURIComponent(state.workbook)}&sheet=${encodeURIComponent(state.sheet)}`, {
        method: "PUT",
        body: { append_direction: state.appendDirection, duplicate_check_columns: dupSelected, totals },
      });
      closeModal();
      toast("Format saved.", "success");
      await loadSheetData();
    } catch (err) {
      if (err.status === 400 && err.data && err.data.invalid) {
        const lines = Object.entries(err.data.invalid).map(([col, reason]) => `${col}: ${reason}`).join("; ");
        toast(`Some formulas couldn't be saved — ${lines}`, "error");
        formatSaveBtn.disabled = false;
      } else {
        toast(err.message, "error");
        formatSaveBtn.disabled = false;
      }
    }
  });
}

// ---------------- Modal helpers ----------------

function openModal(innerHtml) {
  modalRoot.innerHTML = `<div class="modal-backdrop"><div class="modal">${innerHtml}</div></div>`;
  modalRoot.querySelectorAll("[data-close]").forEach((btn) => btn.addEventListener("click", closeModal));
  modalRoot.querySelector(".modal-backdrop").addEventListener("click", (e) => {
    if (e.target.classList.contains("modal-backdrop")) closeModal();
  });
  document.addEventListener("keydown", escCloseOnce);
}

function escCloseOnce(e) {
  if (e.key === "Escape") { closeModal(); }
}

function closeModal() {
  modalRoot.innerHTML = "";
  document.removeEventListener("keydown", escCloseOnce);
}

function confirmDialog(title, message, listItems) {
  return new Promise((resolve) => {
    const listHtml = listItems && listItems.length
      ? `<ul class="duplicate-list">${listItems.map((i) => `<li>${escapeHtml(i)}</li>`).join("")}</ul>`
      : "";
    openModal(`
      <h2>${escapeHtml(title)}</h2>
      <p class="modal-sub">${escapeHtml(message)}</p>
      ${listHtml}
      <div class="modal-actions">
        <button class="btn" id="confirm-no">Cancel</button>
        <button class="btn btn-primary" id="confirm-yes">Continue</button>
      </div>
    `);
    el("confirm-no").addEventListener("click", () => { closeModal(); resolve(false); });
    el("confirm-yes").addEventListener("click", () => { closeModal(); resolve(true); });
  });
}

// ---------------- Boot ----------------

checkSession();
