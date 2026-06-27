import { TabulatorFull as Tabulator } from "tabulator-tables";
import "tabulator-tables/dist/css/tabulator.min.css";
import * as XLSX from "xlsx";
import "./style.css";

// Tabulator's xlsx download looks for a global SheetJS instance.
window.XLSX = XLSX;

const DATA_BASE = "data";

// Each source file maps to a certification status.
const SOURCES = {
  effective: { url: `${DATA_BASE}/effective.json`, status: "active" },
  noneffective: { url: `${DATA_BASE}/noneffective.json`, status: "expired" },
};

// Status scope = which source(s) feed the table.
const SCOPES = {
  active: { sources: ["effective"], file: "bioreg-deystvashti" },
  expired: { sources: ["noneffective"], file: "bioreg-iztekli" },
  all: { sources: ["effective", "noneffective"], file: "bioreg-vsichki" },
};

const STATUS_LABEL = { active: "Действащ", expired: "Изтекъл" };

const el = (id) => document.getElementById(id);
const ui = {
  table: el("table"),
  status: el("status"),
  search: el("globalSearch"),
  clearSearch: el("clearSearch"),
  district: el("filterDistrict"),
  activity: el("filterActivity"),
  clearFilters: el("clearFilters"),
  resultCount: el("resultCount"),
  exportXlsx: el("exportXlsx"),
  exportCsv: el("exportCsv"),
  segs: Array.from(document.querySelectorAll(".seg")),
  statActive: el("stat-active"),
  statExpired: el("stat-expired"),
  statDistricts: el("stat-districts"),
  // detail modal
  detail: el("detail"),
  detailStatus: el("detailStatus"),
  detailName: el("detailName"),
  detailContract: el("detailContract"),
  detailController: el("detailController"),
  detailDistrict: el("detailDistrict"),
  detailActivities: el("detailActivities"),
  detailCerts: el("detailCerts"),
};

const sourceCache = new Map(); // source key -> tagged rows
let table = null;
let scope = "active";

const nf = new Intl.NumberFormat("bg-BG");

const SEARCH_FIELDS = [
  "contractCode",
  "companyName",
  "controllerName",
  "districtName",
  "activitiesText",
  "certificateNumbers",
];

// ---------- activities ----------
// Source separates activities with <br/>; strip the leading "- " bullet.
function splitActivities(text) {
  if (!text) return [];
  return text
    .split(/<br\s*\/?>/i)
    .map((p) => p.replace(/^[\s\-–—]+/, "").trim())
    .filter(Boolean);
}

const escapeHtml = (s) =>
  String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);

// Certificate links are rebuilt from validated data — never from upstream HTML.
const CERT_HOST = "https://bioreg.mzh.government.bg/";
function renderCertificates(certs) {
  if (!Array.isArray(certs) || !certs.length) return "—";
  const links = certs
    .filter((c) => typeof c?.url === "string" && c.url.startsWith(CERT_HOST))
    .map(
      (c) =>
        `<a class="cert-link" href="${escapeHtml(encodeURI(c.url))}" target="_blank" rel="noopener noreferrer">${escapeHtml(c.number)}</a>`,
    );
  return links.length ? links.join("<br>") : "—";
}

// ---------- formatters ----------
function statusPill(cell) {
  const s = cell.getValue();
  return `<span class="pill pill--${s}">${STATUS_LABEL[s] || s}</span>`;
}

function activitiesFormatter(cell) {
  const items = splitActivities(cell.getValue());
  if (!items.length) return "";
  return `<ul class="act-list">${items
    .map((a) => `<li>${escapeHtml(a)}</li>`)
    .join("")}</ul>`;
}

const columns = [
  {
    title: "Статус",
    field: "status",
    width: 130,
    formatter: statusPill,
    tooltip: false,
    download: false,
  },
  { title: "Договор", field: "contractCode", width: 190, cssClass: "cell-mono" },
  { title: "Оператор", field: "companyName", minWidth: 200, widthGrow: 3, cssClass: "cell-name" },
  { title: "Контролиращо лице", field: "controllerName", minWidth: 190, widthGrow: 3, cssClass: "cell-muted" },
  { title: "Област", field: "districtName", width: 150 },
  {
    title: "Дейности",
    field: "activitiesText",
    minWidth: 180,
    widthGrow: 2,
    headerSort: false,
    formatter: activitiesFormatter,
    cssClass: "cell-muted",
    tooltip: (e, cell) => splitActivities(cell.getValue()).join(", "),
    // Export as a clean comma-separated list instead of raw <br/> markup.
    accessorDownload: (value) => splitActivities(value).join(", "),
  },
  {
    title: "Сертификат",
    field: "certificates",
    minWidth: 200,
    widthGrow: 2,
    headerSort: false,
    tooltip: false,
    formatter: (cell) => renderCertificates(cell.getValue()),
    download: false,
  },
  // Export-only columns.
  { title: "Статус", field: "statusLabel", visible: false, download: true },
  { title: "Сертификати", field: "certificateNumbers", visible: false, download: true },
  { title: "Връзка към сертификат", field: "certificateUrl", visible: false, download: true },
];

// ---------- data ----------
async function loadSource(key) {
  if (sourceCache.has(key)) return sourceCache.get(key);
  const src = SOURCES[key];
  const res = await fetch(src.url);
  if (!res.ok) throw new Error(`HTTP ${res.status} при зареждане на ${src.url}`);
  const rows = await res.json();
  const tagged = rows.map((r) => ({
    ...r,
    status: src.status,
    statusLabel: STATUS_LABEL[src.status],
  }));
  sourceCache.set(key, tagged);
  return tagged;
}

async function rowsForScope(s) {
  const parts = await Promise.all(SCOPES[s].sources.map(loadSource));
  return parts.flat();
}

// ---------- filters ----------
function activeFilter() {
  const term = ui.search.value.trim().toLowerCase();
  const district = ui.district.value;
  const activity = ui.activity.value;
  return (row) => {
    if (district && row.districtName !== district) return false;
    if (activity && !(row.activitiesText || "").includes(activity)) return false;
    if (term) {
      let hit = false;
      for (const f of SEARCH_FIELDS) {
        const v = row[f];
        if (v && v.toLowerCase().includes(term)) {
          hit = true;
          break;
        }
      }
      if (!hit) return false;
    }
    return true;
  };
}

function applyFilter() {
  if (!table) return;
  table.setFilter(activeFilter());
  const anyFilter = ui.search.value.trim() || ui.district.value || ui.activity.value;
  ui.clearFilters.hidden = !anyFilter;
  ui.clearSearch.hidden = !ui.search.value;
  updateResultCount();
  writeState();
}

// ---------- shareable URL state ----------
let pendingState = null;

function readUrlState() {
  const p = new URLSearchParams(location.search);
  const s = p.get("scope");
  return {
    scope: SCOPES[s] ? s : "active",
    q: p.get("q") || "",
    district: p.get("district") || "",
    activity: p.get("activity") || "",
  };
}

function writeState() {
  const p = new URLSearchParams();
  if (scope !== "active") p.set("scope", scope);
  if (ui.search.value.trim()) p.set("q", ui.search.value.trim());
  if (ui.district.value) p.set("district", ui.district.value);
  if (ui.activity.value) p.set("activity", ui.activity.value);
  const qs = p.toString();
  history.replaceState(null, "", qs ? `?${qs}` : location.pathname);
}

function applyPendingState() {
  if (!pendingState) return;
  ui.search.value = pendingState.q;
  const has = (sel, val) => Array.from(sel.options).some((o) => o.value === val);
  ui.district.value = has(ui.district, pendingState.district) ? pendingState.district : "";
  ui.activity.value = has(ui.activity, pendingState.activity) ? pendingState.activity : "";
  pendingState = null;
}

function updateResultCount() {
  if (!table) return;
  const shown = table.getDataCount("active");
  const total = table.getDataCount();
  const txt =
    shown === total
      ? `<strong>${nf.format(total)}</strong> оператора`
      : `<strong>${nf.format(shown)}</strong> от ${nf.format(total)} оператора`;
  ui.resultCount.innerHTML = txt;
}

// ---------- option lists ----------
function uniqueSorted(values) {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "bg"),
  );
}

function parseActivities(rows) {
  const set = new Set();
  for (const r of rows) {
    for (const token of splitActivities(r.activitiesText)) set.add(token);
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b, "bg"));
}

function fillSelect(select, values, keepFirst = true) {
  const current = select.value;
  const first = keepFirst ? select.options[0] : null;
  select.innerHTML = "";
  if (first) select.appendChild(first);
  for (const v of values) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  }
  // Restore prior selection if still valid.
  select.value = values.includes(current) ? current : "";
}

function rebuildOptions(rows) {
  const districts = uniqueSorted(rows.map((r) => r.districtName));
  fillSelect(ui.district, districts);
  fillSelect(ui.activity, parseActivities(rows));
  ui.statDistricts.textContent = nf.format(districts.length);
}

// ---------- detail panel ----------
function openDetail(data) {
  ui.detailStatus.className = `pill pill--${data.status}`;
  ui.detailStatus.textContent = STATUS_LABEL[data.status] || data.status;
  ui.detailName.textContent = data.companyName || "—";
  ui.detailContract.textContent = data.contractCode || "—";
  ui.detailController.textContent = data.controllerName || "—";
  ui.detailDistrict.textContent = data.districtName || "—";

  const acts = splitActivities(data.activitiesText);
  ui.detailActivities.innerHTML = acts.length
    ? `<ul class="act-list">${acts.map((a) => `<li>${escapeHtml(a)}</li>`).join("")}</ul>`
    : "—";

  ui.detailCerts.innerHTML = renderCertificates(data.certificates);

  ui.detail.hidden = false;
  document.body.style.overflow = "hidden";
}

function closeDetail() {
  ui.detail.hidden = true;
  document.body.style.overflow = "";
}

// ---------- table ----------
function buildTable(rows) {
  table = new Tabulator(ui.table, {
    data: rows,
    columns,
    layout: "fitColumns",
    height: "68vh",
    minHeight: 420,
    placeholder: "Няма оператори, отговарящи на търсенето.",
    pagination: true,
    paginationSize: 50,
    paginationSizeSelector: [25, 50, 100, 250],
    paginationCounter: "rows",
    initialSort: [{ column: "companyName", dir: "asc" }],
    columnDefaults: {
      resizable: "header",
      tooltip: true,
    },
  });
  table.on("dataFiltered", updateResultCount);
  table.on("tableBuilt", () => {
    applyFilter();
    ui.status.hidden = true;
  });
  table.on("rowClick", (e, row) => {
    // Let certificate links work normally instead of opening the panel.
    if (e.target.closest("a")) return;
    openDetail(row.getData());
  });
}

async function setScope(s) {
  scope = s;
  ui.segs.forEach((b) => b.classList.toggle("is-active", b.dataset.scope === s));

  ui.status.hidden = false;
  ui.status.textContent = "Зареждане на данните…";
  try {
    const rows = await rowsForScope(s);
    rebuildOptions(rows);
    applyPendingState();
    if (!table) {
      buildTable(rows);
    } else {
      await table.replaceData(rows);
      applyFilter();
      ui.status.hidden = true;
    }
  } catch (err) {
    ui.status.hidden = false;
    ui.status.textContent = `Грешка при зареждане: ${err.message}`;
    console.error(err);
  }
}

// ---------- events ----------
let searchTimer = null;
function onSearch() {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(applyFilter, 160);
}

function exportName(ext) {
  return `${SCOPES[scope].file}.${ext}`;
}

function wire() {
  // Logo resets to the default view in-place — no full reload (avoids a FOUC).
  document.querySelector(".brand").addEventListener("click", (e) => {
    e.preventDefault();
    ui.search.value = "";
    ui.district.value = "";
    ui.activity.value = "";
    setScope("active");
  });

  ui.segs.forEach((b) => b.addEventListener("click", () => setScope(b.dataset.scope)));
  ui.search.addEventListener("input", onSearch);
  ui.clearSearch.addEventListener("click", () => {
    ui.search.value = "";
    applyFilter();
    ui.search.focus();
  });
  ui.district.addEventListener("change", applyFilter);
  ui.activity.addEventListener("change", applyFilter);
  ui.clearFilters.addEventListener("click", () => {
    ui.search.value = "";
    ui.district.value = "";
    ui.activity.value = "";
    applyFilter();
  });
  // "active" range = all rows matching the current filters (every page), not the
  // whole dataset and not just the visible page.
  ui.exportXlsx.addEventListener("click", () => {
    table?.download("xlsx", exportName("xlsx"), { sheetName: "Регистър" }, "active");
  });
  ui.exportCsv.addEventListener("click", () => {
    table?.download("csv", exportName("csv"), { bom: true }, "active");
  });

  // Detail modal: close on backdrop, ✕, or Escape.
  ui.detail.querySelectorAll("[data-close]").forEach((b) =>
    b.addEventListener("click", closeDetail),
  );
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !ui.detail.hidden) closeDetail();
  });
}

async function loadMeta() {
  try {
    const res = await fetch(`${DATA_BASE}/meta.json`);
    if (!res.ok) return;
    const meta = await res.json();
    ui.statActive.textContent = nf.format(meta.effectiveCount ?? 0);
    ui.statExpired.textContent = nf.format(meta.nonEffectiveCount ?? 0);
    if (meta.updatedAt) {
      const d = new Date(meta.updatedAt);
      el("updatedAt").textContent = `Обновено: ${d.toLocaleDateString("bg-BG", {
        day: "numeric",
        month: "long",
        year: "numeric",
      })}`;
    }
  } catch {
    /* non-fatal */
  }
}

pendingState = readUrlState();
wire();
loadMeta();
setScope(pendingState.scope);
