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
};

const el = (id) => document.getElementById(id);
const loginView = el("login-view");
const appView = el("app-view");
const tableHead = el("table-head");
const tableBody = el("table-body");
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
    state.workbook = data.workbooks[0].name;
    select.value = state.workbook;
    await loadSheets();
  } else {
    renderEmptyTable("No workbooks found. Upload a .xlsx file to get started.");
  }
}

el("wb-select").addEventListener("change", async (e) => {
  state.workbook = e.target.value;
  await loadSheets();
});

async function loadSheets() {
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
    state.sheet = sheets[0];
    select.value = state.sheet;
    await loadSettings();
    await loadSheetData();
  }
}

el("sheet-select").addEventListener("change", async (e) => {
  state.sheet = e.target.value;
  state.selected.clear();
  await loadSettings();
  await loadSheetData();
});

el("btn-add-sheet").addEventListener("click", async () => {
  const name = prompt("New sheet name:");
  if (!name) return;
  try {
    await api("/api/sheets", { method: "POST", body: { wb: state.workbook, name } });
    toast(`Sheet "${name}" added.`, "success");
    await loadSheets();
  } catch (err) {
    toast(err.message, "error");
  }
});

// ---------------- Settings (append direction) ----------------

async function loadSettings() {
  try {
    const data = await api(`/api/settings?wb=${encodeURIComponent(state.workbook)}`);
    state.appendDirection = data.append_direction || "bottom";
    el("append-select").value = state.appendDirection;
  } catch (_) {
    state.appendDirection = "bottom";
  }
}

el("append-select").addEventListener("change", async (e) => {
  state.appendDirection = e.target.value;
  try {
    await api(`/api/settings?wb=${encodeURIComponent(state.workbook)}`, {
      method: "PUT",
      body: { append_direction: state.appendDirection },
    });
  } catch (err) {
    toast(err.message, "error");
  }
});

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
  state.appendDirection = data.append_direction || state.appendDirection;
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
  const d = new Date(value);
  if (isNaN(d.getTime())) return value;
  return d.toLocaleDateString("en-US", { month: "short", day: "2-digit" });
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

function renderTableBody() {
  const rows = getFilteredSortedRows();
  tableBody.innerHTML = "";

  if (!rows.length) {
    renderEmptyTable(
      state.search || state.flaggedOnly ? "No matching records." : "No records yet — add your first one."
    );
    return;
  }

  rows.forEach((row) => {
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

    state.headers.forEach((h) => {
      if (h === "FLAGGED") return;
      const td = document.createElement("td");
      const isFormula = state.formulaCols.has(h);
      const isNum = state.numericCols.has(h);
      const isDate = state.dateCols.has(h);
      if (isNum) td.classList.add("numeric");
      if (isFormula) {
        td.classList.add("formula");
        td.textContent = row.values[h] ?? "";
      } else {
        td.classList.add("editable");
        td.textContent = isDate ? formatDate(row.values[h]) : (row.values[h] ?? "");
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
}

function renderEmptyTable(message) {
  tableHead.innerHTML = "";
  tableBody.innerHTML = "";
  const tr = document.createElement("tr");
  const td = document.createElement("td");
  td.className = "empty-state";
  td.innerHTML = `<div class="empty-title">${escapeHtml(message)}</div>`;
  tr.appendChild(td);
  tableBody.appendChild(tr);
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
  } catch (err) {
    toast(err.message, "error");
  }
}

el("btn-flag").addEventListener("click", async () => {
  if (!state.selected.size) { toast("Select one or more rows first.", "warn"); return; }
  const rowsByExcel = new Map(state.rows.map((r) => [r.excel_row, r]));
  for (const excelRow of state.selected) {
    const row = rowsByExcel.get(excelRow);
    if (row) await toggleFlag(row);
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
    renderTableBody();
  }, 300);
});

el("flagged-only").addEventListener("change", (e) => {
  state.flaggedOnly = e.target.checked;
  renderTableBody();
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
    <p class="modal-sub">New records are added to the ${state.appendDirection} of this sheet.</p>
    <div class="modal-grid">${formHtml}</div>
    <div class="modal-actions">
      <button class="btn" data-close>Cancel</button>
      <button class="btn btn-primary" id="modal-save">Save</button>
    </div>
  `);

  el("modal-save").addEventListener("click", async () => {
    const inputs = modalRoot.querySelectorAll("[data-col]");
    const values = {};
    inputs.forEach((input) => { values[input.dataset.col] = input.value; });
    try {
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
      }
    }
  });
}

function cssId(str) {
  return String(str).toLowerCase().replace(/[^a-z0-9]+/g, "-");
}

// ---------------- Backups ----------------

el("btn-backups").addEventListener("click", async () => {
  try {
    const data = await api(`/api/backups?wb=${encodeURIComponent(state.workbook)}`);
    const backups = data.backups || [];
    const listHtml = backups.length
      ? `<ul class="backup-list">${backups.map((b) => `
          <li>
            <div>
              <div>${escapeHtml(b.filename)}</div>
              <div class="backup-meta">${new Date(b.created_at).toLocaleString()} · ${Math.round((b.size_bytes || 0) / 1024)} KB</div>
            </div>
            <a class="btn btn-quiet" href="/api/backups/${encodeURIComponent(b.filename)}/download">Download</a>
          </li>`).join("")}</ul>`
      : `<p class="modal-sub">No backups yet — one is created automatically every time you save.</p>`;
    openModal(`
      <h2>Backups</h2>
      <p class="modal-sub">The last 5 saved versions of this workbook.</p>
      ${listHtml}
      <div class="modal-actions"><button class="btn" data-close>Close</button></div>
    `);
  } catch (err) {
    toast(err.message, "error");
  }
});

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

// ---------------- Formulas panel ----------------
// NOTE: depends on GET/PUT /api/formulas, which does not exist in the
// backend yet (see handover prompt). Until OpenCode adds it, this shows
// a clear "not available yet" message instead of failing silently.

el("btn-formulas").addEventListener("click", async () => {
  try {
    const data = await api(
      `/api/formulas?wb=${encodeURIComponent(state.workbook)}&sheet=${encodeURIComponent(state.sheet)}`
    );
    const cols = data.formulas || {};
    const rowsHtml = Object.keys(cols).length
      ? Object.entries(cols).map(([col, formula]) => `
          <div class="modal-field">
            <label for="formula-${cssId(col)}">${escapeHtml(col)}</label>
            <input id="formula-${cssId(col)}" type="text" data-formula-col="${escapeHtml(col)}"
                   value="${escapeHtml(formula || "")}" placeholder="e.g. =D2*1.18">
          </div>`).join("")
      : `<p class="modal-sub">No formula columns on this sheet yet. Add one below.</p>
         <div class="modal-field">
           <label for="formula-new-col">Column name</label>
           <input id="formula-new-col" type="text" placeholder="e.g. GST TOTAL">
         </div>
         <div class="modal-field">
           <label for="formula-new-value">Formula</label>
           <input id="formula-new-value" type="text" placeholder="e.g. =D2*1.18">
         </div>`;
    openModal(`
      <h2>Formulas — ${escapeHtml(state.sheet)}</h2>
      <p class="modal-sub">These apply to every new row added to this sheet. Edit and save to update the template.</p>
      <div class="modal-grid">${rowsHtml}</div>
      <div class="modal-actions">
        <button class="btn" data-close>Cancel</button>
        <button class="btn btn-primary" id="modal-save-formulas">Save</button>
      </div>
    `);
    el("modal-save-formulas").addEventListener("click", async () => {
      const updates = {};
      modalRoot.querySelectorAll("[data-formula-col]").forEach((input) => {
        updates[input.dataset.formulaCol] = input.value;
      });
      const newCol = el("formula-new-col");
      const newVal = el("formula-new-value");
      if (newCol && newVal && newCol.value && newVal.value) {
        updates[newCol.value] = newVal.value;
      }
      try {
        await api(`/api/formulas?wb=${encodeURIComponent(state.workbook)}&sheet=${encodeURIComponent(state.sheet)}`, {
          method: "PUT",
          body: { formulas: updates },
        });
        closeModal();
        toast("Formulas updated.", "success");
        await loadSheetData();
      } catch (err) {
        toast(err.message, "error");
      }
    });
  } catch (err) {
    toast("Formula editing isn't available in this version yet.", "warn");
  }
});

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
