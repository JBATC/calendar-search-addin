"use strict";
// =====================================================================
// Calendar Search (Tabbed) — Excel task pane add-in
//   Tab 1  Campaign Search      find text on the schedule
//   Tab 2  Destination Summary  slots / sold / open / rate card per placement
//   Tab 3  Open Placements      every unsold slot for a destination
// Everything renders in this panel only (private to each user).
// Nothing is written to the workbook.
// =====================================================================
// ---- Sheet name ----
const SCHEDULE_SHEET_NAME = "2027 Schedule";
// ---- Schedule layout (0-based indexes) ----
const SCHEDULE_RATE_COL_Q1 = 0; // column A
const SCHEDULE_RATE_COL_Q2 = 1; // column B
const SCHEDULE_PLACEMENT_COL = 2; // column C
const SCHEDULE_DATE_ROW = 0; // row 1
const SCHEDULE_Q2_START_COL = 16; // column Q onward uses the Q2 rate (colIndex < 16 => Q1)
const SCHEDULE_FIRST_ROW = 13; // row 14 — nothing above this is searched or summarised
// ---- Destinations: the cell style names that mark a destination slot ----
const DESTINATION_STYLES = [
    "Asia",
    "Cab",
    "Central America",
    "Cruise",
    "Europe",
    "Hawaii",
    "Mexico",
    "Open",
    "SoPac",
    "USA",
];
// ---- "Find me" animation when jumping to a cell (only changes YOUR selection) ----
const FLASH_PULSES = 2; // times the box shrinks onto the cell (0 = off)
const FLASH_STEP_MS = 90; // pause between frames
// ---- Scan limits ----
const SCAN_CHUNK_ROWS = 300; // rows of styles read per round trip
const SCAN_MAX_ROWS = 5000; // safety caps
const SCAN_MAX_COLS = 1000;
const MERGE_BAND_ROWS = 100; // rows per merged-cell lookup (auto-splits if Excel says it's too many)
const COLUMNS = ["Placement", "Date", "Cost", "Supplier"];
let flashToken = 0; // cancels an in-progress flash when another row is clicked
let mergeRowsUnread = 0; // rows whose merged cells couldn't be read
let scanCache = null;
const $ = (id) => document.getElementById(id);
// =====================================================================
// PLACEMENT FILTER — one dropdown per tab ("<key>-placement").
// Options are the placement names in the tab's current data, in sheet order.
// =====================================================================
function placementFilter(key) {
    return $(`${key}-placement`);
}
/** Selected placement, or "" for all. */
function placementValue(key) {
    return placementFilter(key).value;
}
/** Back to "All placements" — used when a new search or destination is chosen. */
function resetPlacementFilter(key) {
    placementFilter(key).value = "";
}
/** Rebuilds the dropdown from `names`, keeping the current choice if it still exists. */
function fillPlacementFilter(key, names) {
    const sel = placementFilter(key);
    const current = sel.value;
    const unique = [];
    const seen = new Set();
    names.forEach((n) => {
        if (!seen.has(n)) {
            seen.add(n);
            unique.push(n);
        }
    });
    sel.innerHTML =
        `<option value="">All placements${unique.length ? ` (${unique.length})` : ""}</option>` +
            unique.map((n) => `<option value="${esc(n)}">${esc(n)}</option>`).join("");
    sel.value = seen.has(current) ? current : "";
    sel.disabled = unique.length === 0;
}
/** " · filtered to “X”" when a placement is selected. */
function filterNote(key) {
    const p = placementValue(key);
    return p ? ` · filtered to “${p}”` : "";
}
// =====================================================================
// TAB 1 — CAMPAIGN SEARCH
// =====================================================================
const searchView = {
    key: "cs",
    results: [],
    sortCol: null,
    sortDir: 1,
    activeIndex: -1,
    highlightTerm: "",
    statusText: () => {
        const n = searchView.results.length;
        return n
            ? `${n} result${n === 1 ? "" : "s"} for “${searchView.highlightTerm}”`
            : `No results for “${searchView.highlightTerm}”.`;
    },
};
async function runSearch() {
    const term = $("search-term").value.trim();
    if (!term) {
        $("cs-status").textContent = "Type something to search for.";
        return;
    }
    const btn = $("search-btn");
    btn.disabled = true;
    $("cs-status").textContent = "Searching…";
    try {
        await Excel.run(async (context) => {
            const scheduleSheet = context.workbook.worksheets.getItem(SCHEDULE_SHEET_NAME);
            const scheduleUsed = scheduleSheet.getUsedRangeOrNullObject(true);
            scheduleUsed.load("isNullObject");
            await context.sync();
            const results = [];
            if (!scheduleUsed.isNullObject) {
                // Read from A1 so grid[r][c] lines up with 0-based sheet indexes.
                const scheduleGrid = scheduleSheet.getRange("A1").getBoundingRect(scheduleUsed);
                scheduleGrid.load("text, values");
                await context.sync();
                const t = scheduleGrid.text;
                const v = scheduleGrid.values;
                forEachMatch(t, term.toLowerCase(), (r, c, cellText) => {
                    if (r < SCHEDULE_FIRST_ROW)
                        return; // nothing above row 14
                    const rateCol = c < SCHEDULE_Q2_START_COL ? SCHEDULE_RATE_COL_Q1 : SCHEDULE_RATE_COL_Q2;
                    results.push({
                        row: r,
                        col: c,
                        placement: cell(t, r, SCHEDULE_PLACEMENT_COL),
                        date: cell(t, SCHEDULE_DATE_ROW, c),
                        rate: cell(t, r, rateCol),
                        match: cellText,
                        dateVal: val(v, SCHEDULE_DATE_ROW, c),
                        rateVal: val(v, r, rateCol),
                    });
                });
            }
            searchView.results = results;
            searchView.activeIndex = -1;
            searchView.highlightTerm = term;
            resetPlacementFilter("cs");
            renderView(searchView);
        });
    }
    finally {
        btn.disabled = false;
    }
}
/** Calls fn for every cell whose displayed text contains needle (case-insensitive, partial match). */
function forEachMatch(grid, needle, fn) {
    for (let r = 0; r < grid.length; r++) {
        const row = grid[r];
        for (let c = 0; c < row.length; c++) {
            const v = row[c];
            if (v && v.toLowerCase().indexOf(needle) !== -1)
                fn(r, c, v);
        }
    }
}
// =====================================================================
// DESTINATION SCAN — shared by Destination Summary and Open Placements.
// A styled slot with visible text = SOLD; blank or only invisible characters = OPEN.
// Rows with no placement name are ignored. Merged blocks count once.
// =====================================================================
async function ensureScan(statusId) {
    if (scanCache)
        return;
    const controls = ["sum-select", "sum-refresh", "op-select", "op-refresh"].map((id) => $(id));
    controls.forEach((c) => (c.disabled = true));
    try {
        scanCache = { scannedAt: new Date(), slots: await scanDestinations(statusId) };
    }
    finally {
        controls.forEach((c) => (c.disabled = false));
    }
}
async function scanDestinations(statusId) {
    const destSet = new Set(DESTINATION_STYLES);
    const slots = [];
    const status = (msg) => ($(statusId).textContent = msg);
    mergeRowsUnread = 0;
    await Excel.run(async (context) => {
        status("Reading calendar…");
        const sheet = context.workbook.worksheets.getItem(SCHEDULE_SHEET_NAME);
        // Include formatting-only cells: an open slot is a styled cell with no text.
        const used = sheet.getUsedRangeOrNullObject(false);
        used.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
        await context.sync();
        if (used.isNullObject)
            return;
        const nRows = Math.min(used.rowIndex + used.rowCount, SCAN_MAX_ROWS);
        const nCols = Math.min(used.columnIndex + used.columnCount, SCAN_MAX_COLS);
        const full = sheet.getRangeByIndexes(0, 0, nRows, nCols); // from A1
        full.load("text, values");
        await context.sync();
        status("Reading merged cells…");
        const skip = new Set();
        await collectMerges(context, sheet, 0, nRows, nCols, skip);
        const styles = [];
        for (let r0 = 0; r0 < nRows; r0 += SCAN_CHUNK_ROWS) {
            const n = Math.min(SCAN_CHUNK_ROWS, nRows - r0);
            const props = sheet.getRangeByIndexes(r0, 0, n, nCols).getCellProperties({ style: true });
            await context.sync();
            props.value.forEach((row) => styles.push(row.map((p) => p.style || "")));
            status(`Reading styles… ${Math.min(r0 + n, nRows)} / ${nRows} rows`);
        }
        const text = full.text;
        const values = full.values;
        for (let r = Math.max(SCHEDULE_FIRST_ROW, SCHEDULE_DATE_ROW + 1); r < nRows; r++) {
            const placement = cell(text, r, SCHEDULE_PLACEMENT_COL);
            if (isBlank(placement))
                continue; // no placement name -> ignore
            for (let c = SCHEDULE_PLACEMENT_COL + 1; c < nCols; c++) {
                const style = styles[r] ? styles[r][c] : "";
                if (!destSet.has(style) || skip.has(r + "," + c))
                    continue;
                const rateCol = c < SCHEDULE_Q2_START_COL ? SCHEDULE_RATE_COL_Q1 : SCHEDULE_RATE_COL_Q2;
                const cellText = cell(text, r, c);
                slots.push({
                    row: r,
                    col: c,
                    dest: style,
                    placement: placement.trim(),
                    date: cell(text, SCHEDULE_DATE_ROW, c),
                    rate: cell(text, r, rateCol),
                    match: cellText,
                    dateVal: val(values, SCHEDULE_DATE_ROW, c),
                    rateVal: val(values, r, rateCol),
                    sold: !isBlank(cellText),
                });
            }
        }
    });
    return slots;
}
/** Adds every non-top-left cell of each merged block in rows [r0, r0+n) to `skip`. */
async function collectMerges(context, sheet, r0, n, nCols, skip) {
    for (let start = r0; start < r0 + n; start += MERGE_BAND_ROWS) {
        await collectMergeBand(context, sheet, start, Math.min(MERGE_BAND_ROWS, r0 + n - start), nCols, skip);
    }
}
async function collectMergeBand(context, sheet, r0, n, nCols, skip) {
    try {
        const merged = sheet.getRangeByIndexes(r0, 0, n, nCols).getMergedAreasOrNullObject();
        merged.load("isNullObject");
        await context.sync();
        if (merged.isNullObject)
            return;
        merged.areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
        await context.sync();
        merged.areas.items.forEach((a) => {
            for (let r = a.rowIndex; r < a.rowIndex + a.rowCount; r++)
                for (let c = a.columnIndex; c < a.columnIndex + a.columnCount; c++)
                    if (r !== a.rowIndex || c !== a.columnIndex)
                        skip.add(r + "," + c);
        });
    }
    catch (e) {
        // Excel refuses when a range holds too many merged areas (~512): split and retry.
        if (n <= 1) {
            mergeRowsUnread++;
            return;
        }
        const half = Math.floor(n / 2);
        await collectMergeBand(context, sheet, r0, half, nCols, skip);
        await collectMergeBand(context, sheet, r0 + half, n - half, nCols, skip);
    }
}
function scanNote() {
    if (!scanCache)
        return "";
    const time = scanCache.scannedAt.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    return ` · scanned ${time}` + (mergeRowsUnread ? ` · ⚠ merged cells unreadable on ${mergeRowsUnread} row(s)` : "");
}
const SUMMARY_COLUMNS = ["Placement", "Slots", "Sold", "Open", "Rate card", "Open value"];
let summaryRows = [];
async function showSummary() {
    const dest = $("sum-select").value;
    if (!dest) {
        summaryRows = [];
        fillPlacementFilter("sum", []);
        $("sum-results").innerHTML = "";
        $("sum-status").textContent = "Pick a destination to summarise its inventory.";
        $("sum-copy").disabled = true;
        return;
    }
    await ensureScan("sum-status");
    resetPlacementFilter("sum");
    const map = new Map();
    scanCache.slots
        .filter((s) => s.dest === dest)
        .forEach((s) => {
        let p = map.get(s.placement);
        if (!p) {
            p = { placement: s.placement, order: s.row, slots: 0, sold: 0, open: 0, rateTotal: 0, openTotal: 0, unpriced: 0 };
            map.set(s.placement, p);
        }
        const rate = numericRate(s);
        p.slots++;
        if (s.sold)
            p.sold++;
        else
            p.open++;
        if (rate === null)
            p.unpriced++;
        else {
            p.rateTotal += rate;
            if (!s.sold)
                p.openTotal += rate;
        }
    });
    summaryRows = Array.from(map.values()).sort((a, b) => a.order - b.order);
    renderSummary(dest);
}
function summaryTotals(rows) {
    return rows.reduce((t, r) => ({
        slots: t.slots + r.slots,
        sold: t.sold + r.sold,
        open: t.open + r.open,
        rateTotal: t.rateTotal + r.rateTotal,
        openTotal: t.openTotal + r.openTotal,
        unpriced: t.unpriced + r.unpriced,
    }), { slots: 0, sold: 0, open: 0, rateTotal: 0, openTotal: 0, unpriced: 0 });
}
/** Summary rows that pass the placement filter. Copy uses this too. */
function visibleSummaryRows() {
    const p = placementValue("sum");
    return p ? summaryRows.filter((r) => r.placement === p) : summaryRows;
}
function renderSummary(dest) {
    fillPlacementFilter("sum", summaryRows.map((r) => r.placement));
    const rows = visibleSummaryRows();
    const t = summaryTotals(rows);
    $("sum-copy").disabled = rows.length === 0;
    $("sum-status").textContent =
        (rows.length
            ? `${dest}: ${t.slots.toLocaleString()} slots across ${rows.length} placement${rows.length === 1 ? "" : "s"} · ${t.sold.toLocaleString()} sold · ${t.open.toLocaleString()} open`
            : `No slots styled “${dest}” on the calendar`) +
            filterNote("sum") +
            scanNote();
    if (!rows.length) {
        $("sum-results").innerHTML = "";
        return;
    }
    const body = rows
        .map((r) => `
      <tr${r.unpriced ? ` title="${r.unpriced} slot(s) without a numeric rate"` : ""}>
        <td>${esc(r.placement)}</td>
        <td class="num">${r.slots.toLocaleString()}</td>
        <td class="num">${r.sold.toLocaleString()}</td>
        <td class="num">${r.open.toLocaleString()}</td>
        <td class="num">${money(r.rateTotal)}</td>
        <td class="num">${money(r.openTotal)}</td>
      </tr>`)
        .join("");
    $("sum-results").innerHTML = `
    <table class="nosort">
      <thead><tr>${SUMMARY_COLUMNS.map((c, i) => `<th class="${i ? "num" : ""}">${c}</th>`).join("")}</tr></thead>
      <tbody>${body}</tbody>
      <tfoot><tr>
        <td>Total</td>
        <td class="num">${t.slots.toLocaleString()}</td>
        <td class="num">${t.sold.toLocaleString()}</td>
        <td class="num">${t.open.toLocaleString()}</td>
        <td class="num">${money(t.rateTotal)}</td>
        <td class="num">${money(t.openTotal)}</td>
      </tr></tfoot>
    </table>
    ${t.unpriced ? `<p class="muted note">${t.unpriced} slot(s) have no numeric rate and aren't included in the dollar totals.</p>` : ""}`;
}
async function copySummary() {
    const rows = visibleSummaryRows();
    const t = summaryTotals(rows);
    const lines = [tsvLine(SUMMARY_COLUMNS)]
        .concat(rows.map((r) => tsvLine([r.placement, r.slots, r.sold, r.open, money(r.rateTotal), money(r.openTotal)])))
        .concat([tsvLine(["Total", t.slots, t.sold, t.open, money(t.rateTotal), money(t.openTotal)])]);
    await copyText(lines.join("\r\n"), $("sum-copy"), "Copy summary");
}
// =====================================================================
// TAB 3 — OPEN PLACEMENTS (same table as Campaign Search, unsold slots only)
// =====================================================================
const openView = {
    key: "op",
    results: [],
    sortCol: null,
    sortDir: 1,
    activeIndex: -1,
    highlightTerm: "",
    statusText: () => {
        const dest = $("op-select").value;
        const n = openView.results.length;
        return (n ? `${dest}: ${n} open slot${n === 1 ? "" : "s"}` : `${dest}: no open slots`) + scanNote();
    },
};
async function showOpen() {
    const dest = $("op-select").value;
    if (!dest) {
        openView.results = [];
        fillPlacementFilter("op", []);
        $("op-results").innerHTML = "";
        $("op-status").textContent = "Pick a destination to list its open (unsold) slots. Click a row to jump to it.";
        $("op-copy").disabled = true;
        return;
    }
    await ensureScan("op-status");
    openView.results = scanCache.slots.filter((s) => s.dest === dest && !s.sold);
    openView.activeIndex = -1;
    resetPlacementFilter("op");
    renderView(openView);
}
// =====================================================================
// SHARED RESULTS TABLE (Campaign Search + Open Placements)
// =====================================================================
function initView(v) {
    $(`${v.key}-results`).addEventListener("click", (e) => {
        const target = e.target;
        // Column header -> sort
        const th = target.closest("th[data-col]");
        if (th) {
            const col = Number(th.dataset.col);
            if (v.sortCol === col)
                v.sortDir = v.sortDir === 1 ? -1 : 1;
            else {
                v.sortCol = col;
                v.sortDir = 1;
            }
            renderView(v);
            return;
        }
        // Row -> jump to the cell (unless the user is drag-selecting text to copy)
        const sel = window.getSelection();
        if (sel && sel.toString().length > 0)
            return;
        const row = target.closest("tr[data-i]");
        if (row) {
            v.activeIndex = Number(row.dataset.i);
            $(`${v.key}-results`)
                .querySelectorAll("tr.active")
                .forEach((el) => el.classList.remove("active"));
            row.classList.add("active");
            tryCatch(() => goTo(v.results[v.activeIndex]), `${v.key}-status`);
        }
    });
    $(`${v.key}-copy`).addEventListener("click", () => tryCatch(() => copyView(v), `${v.key}-status`));
    placementFilter(v.key).addEventListener("change", () => renderView(v));
}
/** Results that pass the placement filter, in the current sort order. Copy uses this too. */
function visibleRows(v) {
    const p = placementValue(v.key);
    const rows = v.results.map((r, i) => ({ r, i })).filter(({ r }) => !p || r.placement === p);
    if (v.sortCol !== null) {
        const col = v.sortCol;
        rows.sort((a, b) => compareValues(sortKey(a.r, col), sortKey(b.r, col), v.sortDir) || a.i - b.i);
    }
    return rows;
}
function renderView(v) {
    fillPlacementFilter(v.key, v.results.map((r) => r.placement));
    const rows = visibleRows(v);
    $(`${v.key}-copy`).disabled = rows.length === 0;
    const box = $(`${v.key}-results`);
    if (rows.length === 0) {
        box.innerHTML = v.results.length
            ? `<p class="empty">No rows for this placement.</p>`
            : `<p class="empty">No matches.</p>`;
    }
    else {
        const body = rows
            .map(({ r, i }) => `
        <tr class="${i === v.activeIndex ? "active" : ""}" data-i="${i}" title="Click to go to ${esc(cellAddress(r.row, r.col))}">
          <td>${esc(r.placement)}</td>
          <td class="nowrap">${esc(r.date)}</td>
          <td class="nowrap">${esc(r.rate)}</td>
          <td>${highlight(r.match, v.highlightTerm)}</td>
        </tr>`)
            .join("");
        const total = costTotal(rows.map((x) => x.r));
        const foot = total.counted === 0
            ? ""
            : `
        <tfoot><tr>
          <td>Total</td><td></td><td class="nowrap">${esc(total.text)}</td><td>${total.skipped ? `<span class="muted">${total.skipped} not counted</span>` : ""}</td>
        </tr></tfoot>`;
        box.innerHTML = `
      <table>
        <thead><tr>${COLUMNS.map((c, ci) => {
            const cls = v.sortCol === ci ? (v.sortDir === 1 ? "sorted asc" : "sorted desc") : "";
            const tip = v.sortCol === ci
                ? v.sortDir === 1
                    ? "Sorted ascending — click for descending"
                    : "Sorted descending — click for ascending"
                : "Click to sort";
            return `<th data-col="${ci}" class="${cls}" title="${tip}">${c}</th>`;
        }).join("")}</tr></thead>
        <tbody>${body}</tbody>${foot}
      </table>`;
    }
    const p = placementValue(v.key);
    $(`${v.key}-status`).textContent =
        v.statusText() + (p ? ` · showing ${rows.length} in “${p}”` : "");
}
async function copyView(v) {
    const rows = visibleRows(v);
    const lines = [tsvLine(COLUMNS)].concat(rows.map(({ r }) => tsvLine([r.placement, r.date, r.rate, r.match])));
    const total = costTotal(rows.map((x) => x.r));
    if (total.counted > 0)
        lines.push(tsvLine(["Total", "", total.text, ""]));
    await copyText(lines.join("\r\n"), $(`${v.key}-copy`), "Copy results");
}
/**
 * Jump to a cell, then draw the eye to it: the selection box shrinks
 * 7x7 -> 5x5 -> 3x3 -> cell, FLASH_PULSES times. Only the user's own selection
 * changes — nothing in the sheet is modified.
 */
async function goTo(r) {
    const token = ++flashToken;
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(SCHEDULE_SHEET_NAME);
        sheet.activate();
        sheet.getCell(r.row, r.col).select();
        await context.sync();
        for (let p = 0; p < FLASH_PULSES; p++) {
            for (const size of [3, 2, 1, 0]) {
                if (token !== flashToken)
                    return; // another row was clicked
                const top = Math.max(0, r.row - size);
                const left = Math.max(0, r.col - size);
                sheet.getRangeByIndexes(top, left, r.row + size - top + 1, r.col + size - left + 1).select();
                await context.sync();
                await sleep(FLASH_STEP_MS);
            }
        }
    });
}
// =====================================================================
// Helpers
// =====================================================================
/** True when text is empty or only spaces / non-breaking spaces / zero-width characters. */
function isBlank(s) {
    return String(s == null ? "" : s).replace(/[\s ᠎​-‍⁠﻿]/g, "") === "";
}
function numericRate(r) {
    const n = typeof r.rateVal === "number" ? r.rateVal : parseNumber(r.rate);
    return typeof n === "number" ? n : null;
}
/** Placement & Supplier sort as text; Date & Cost sort by their real underlying values. */
function sortKey(r, col) {
    switch (col) {
        case 0:
            return r.placement;
        case 1:
            return typeof r.dateVal === "number" ? r.dateVal : r.date;
        case 2:
            return typeof r.rateVal === "number" ? r.rateVal : parseNumber(r.rate);
        default:
            return r.match;
    }
}
/** Blanks last; numbers before text; natural text order ("Item 2" < "Item 10"). */
function compareValues(a, b, dir) {
    const aBlank = a === "" || a === null || a === undefined;
    const bBlank = b === "" || b === null || b === undefined;
    if (aBlank || bBlank)
        return aBlank === bBlank ? 0 : aBlank ? 1 : -1;
    if (typeof a === "number" && typeof b === "number")
        return (a - b) * dir;
    if (typeof a === "number")
        return -1 * dir;
    if (typeof b === "number")
        return 1 * dir;
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }) * dir;
}
/** Sums Cost; non-numeric costs (e.g. "TBD") are skipped and counted separately. */
function costTotal(results) {
    let sum = 0, counted = 0, skipped = 0, currency = false;
    for (const r of results) {
        const n = numericRate(r);
        if (n !== null) {
            sum += n;
            counted++;
        }
        else if (r.rate.trim() !== "")
            skipped++;
        if (r.rate.indexOf("$") !== -1)
            currency = true;
    }
    const text = currency ? money(sum) : sum.toLocaleString("en-US", { maximumFractionDigits: 2 });
    return { text, counted, skipped };
}
/** "$1,250.00" -> 1250; returns the original text if it isn't a number. */
function parseNumber(s) {
    const n = Number(s.replace(/[^0-9.\-]/g, ""));
    return s.trim() !== "" && !isNaN(n) && /\d/.test(s) ? n : s;
}
function money(n) {
    return n.toLocaleString("en-US", { style: "currency", currency: "USD" });
}
function tsvLine(cells) {
    return cells.map((c) => String(c).replace(/[\t\r\n]+/g, " ").trim()).join("\t");
}
async function copyText(text, btn, label) {
    let ok = false;
    try {
        await navigator.clipboard.writeText(text);
        ok = true;
    }
    catch (_a) {
        // Clipboard API can be blocked inside Office panels; fall back to execCommand
        const ta = document.createElement("textarea");
        ta.value = text;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        ok = document.execCommand("copy");
        document.body.removeChild(ta);
    }
    btn.textContent = ok ? "Copied ✓" : "Copy failed";
    setTimeout(() => (btn.textContent = label), 1500);
}
function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}
function cell(grid, r, c) {
    return grid[r] && grid[r][c] != null ? String(grid[r][c]) : "";
}
function val(grid, r, c) {
    return grid[r] ? grid[r][c] : "";
}
function cellAddress(r, c) {
    let s = "";
    let n = c + 1;
    while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
    }
    return `${s}${r + 1}`;
}
function esc(s) {
    return String(s)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}
function highlight(text, term) {
    if (!term)
        return esc(text);
    const lower = text.toLowerCase();
    const t = term.toLowerCase();
    let out = "";
    let i = 0;
    let j;
    while ((j = lower.indexOf(t, i)) !== -1) {
        out += esc(text.slice(i, j)) + "<mark>" + esc(text.slice(j, j + t.length)) + "</mark>";
        i = j + t.length;
    }
    return out + esc(text.slice(i));
}
async function tryCatch(callback, statusId) {
    try {
        await callback();
    }
    catch (error) {
        console.error(error);
        $(statusId).textContent =
            error && error.code === "ItemNotFound"
                ? `Couldn't find a sheet named "${SCHEDULE_SHEET_NAME}". Check the sheet name at the top of the script.`
                : `Error: ${(error === null || error === void 0 ? void 0 : error.message) || error}`;
    }
}
// =====================================================================
// STARTUP — kept at the bottom so every setting above is defined first.
// =====================================================================
Office.onReady((info) => {
    // Guard: this pane only works in Excel with ExcelApi 1.13+ (merged-area lookups).
    if (info.host !== Office.HostType.Excel) {
        $("cs-status").textContent = "Calendar Search only runs inside Excel.";
        return;
    }
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
        $("cs-status").textContent =
            "This version of Excel is too old for Calendar Search (needs ExcelApi 1.13 / Microsoft 365). Please update Office.";
        return;
    }
    // Tabs
    document.querySelectorAll(".tab").forEach((tab) => {
        tab.addEventListener("click", () => {
            document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("on", t === tab));
            document.querySelectorAll(".panel").forEach((p) => (p.hidden = p.id !== tab.dataset.panel));
        });
    });
    // Destination dropdowns
    const options = `<option value="">Choose a destination…</option>` +
        DESTINATION_STYLES.map((d) => `<option value="${esc(d)}">${esc(d)}</option>`).join("");
    $("sum-select").innerHTML = options;
    $("op-select").innerHTML = options;
    // Campaign Search
    $("search-btn").addEventListener("click", () => tryCatch(runSearch, "cs-status"));
    $("search-term").addEventListener("keydown", (e) => {
        if (e.key === "Enter")
            tryCatch(runSearch, "cs-status");
    });
    initView(searchView);
    // Destination Summary
    $("sum-select").addEventListener("change", () => tryCatch(showSummary, "sum-status"));
    $("sum-refresh").addEventListener("click", () => tryCatch(async () => {
        scanCache = null;
        await showSummary();
    }, "sum-status"));
    $("sum-copy").addEventListener("click", () => tryCatch(copySummary, "sum-status"));
    placementFilter("sum").addEventListener("change", () => renderSummary($("sum-select").value));
    // Placement filters start empty until there's data
    ["cs", "sum", "op"].forEach((k) => fillPlacementFilter(k, []));
    // Open Placements
    initView(openView);
    $("op-select").addEventListener("change", () => tryCatch(showOpen, "op-status"));
    $("op-refresh").addEventListener("click", () => tryCatch(async () => {
        scanCache = null;
        await showOpen();
    }, "op-status"));
    $("search-term").focus();
});
