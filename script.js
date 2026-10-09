/* =====================================================================
   PALM SEEDLING DASHBOARD
   ---------------------------------------------------------------------
   TO UPDATE THE DATA: replace  data/palm-data.xlsx  with the new Excel
   file (same file name) on GitHub. Vercel redeploys; no code changes.

   Columns needed in the first sheet (any order):
   Year | Date | District | Department Type | Gov Department Type | Seedlings
   ===================================================================== */

const CONFIG = {
  dataFile: 'data/palm-data.xlsx',
  title: 'GTM - Palm Dibbling Dashboard',
  updatedOn: '09 Oct 2026',   // <-- CHANGE THIS DATE each time you upload a new Excel file
  timeZone: 'Asia/Kolkata'    // "Today" is decided in this time zone
};

let ALL = [];
let YEAR = '';                      // taken from the "Year" column of the Excel, e.g. 2026-27
const state = { district: '', dept: '', period: 'all' };
const LIMIT = { listDistrict: 12, listDept: 6, matrix: 10 };   // rows shown before "Show all"
const showAll = {};                 // which lists are expanded
const openD = new Set();            // districts expanded in the table
let lastDistricts = [];

const $ = id => document.getElementById(id);
const nf = new Intl.NumberFormat('en-IN');
const fmt = n => nf.format(Math.round(n));
const pad = n => String(n).padStart(2, '0');
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fmtDate = iso => iso ? `${iso.slice(8, 10)} ${MON[+iso.slice(5, 7) - 1]} ${iso.slice(0, 4)}` : '—';
const fmtMonth = ym => `${MON[+ym.slice(5, 7) - 1]} ${ym.slice(0, 4)}`;
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/* ---------- 1. read the Excel file ---------- */
const mkIso = (y, m, d) => (y > 1900 && m >= 1 && m <= 12 && d >= 1 && d <= 31) ? `${y}-${pad(m)}-${pad(d)}` : '';

function toIso(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') { const o = XLSX.SSF.parse_date_code(v); return o ? mkIso(o.y, o.m, o.d) : ''; }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);            // 2026-10-06
  if (m) return mkIso(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/);              // 06-10-2026 (day first)
  if (m) return mkIso(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
  return '';
}

const SMALL = new Set(['of', 'and', 'the', 'for', 'in']);
const tidy = v => (v == null ? '' : String(v)).replace(/\s+/g, ' ').trim().toLowerCase().split(' ')
  .map((w, i) => (i && SMALL.has(w)) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(' ');

function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: 'array' });
  const grid = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
  const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const h = grid.findIndex(r => r.some(c => norm(c) === 'district'));
  if (h < 0) throw new Error('No "District" column found in the Excel file.');
  const head = grid[h].map(norm);
  const col = (...names) => { for (const n of names) { const i = head.indexOf(n); if (i >= 0) return i; } return -1; };
  const ix = {
    date: col('date', 'reportdate', 'plantingdate'),
    year: col('year', 'financialyear', 'fy'),
    district: col('district'),
    category: col('departmenttype', 'category', 'type'),
    dept: col('govdepartmenttype', 'govtdepartmenttype', 'governmentdepartmenttype', 'department'),
    seedlings: col('seedlings', 'seedling', 'dibbling', 'palmdibbling', 'dibbled', 'noofseedlings', 'numberofseedlings', 'count')
  };
  if (ix.seedlings < 0) throw new Error('No "Seedlings" (or "Dibbling") column found in the Excel file.');

  const rows = [], years = new Set();
  for (let i = h + 1; i < grid.length; i++) {
    const r = grid[i];
    const district = tidy(r[ix.district]);
    const raw = r[ix.seedlings];
    const seedlings = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(/,/g, ''));
    if (!district || !isFinite(seedlings) || seedlings < 0) continue;      // skip blank / invalid lines
    const category = ix.category >= 0 ? tidy(r[ix.category]) : '';
    const gov = ix.dept >= 0 ? tidy(r[ix.dept]) : '';
    const iso = ix.date >= 0 ? toIso(r[ix.date]) : '';
    if (ix.year >= 0 && String(r[ix.year]).trim()) years.add(String(r[ix.year]).trim());
    rows.push({ iso, ym: iso.slice(0, 7), district, dept: gov || category || 'Not specified', seedlings });
  }
  YEAR = [...years].sort().join(', ');
  return rows;
}

async function load() {
  try {
    const res = await fetch(CONFIG.dataFile, { cache: 'no-cache' });
    if (!res.ok) throw new Error(`Could not load ${CONFIG.dataFile} (HTTP ${res.status}).`);
    ALL = parseWorkbook(await res.arrayBuffer());
    if (!ALL.length) throw new Error('The Excel file has no usable rows.');
    init();
  } catch (e) {
    $('status').className = 'status error';
    $('status').innerHTML = `<b>Data could not be loaded.</b><br>${esc(e.message)}`;
    console.error(e);
  }
}

/* ---------- 2. periods: today / this month / this year ---------- */
const P = {};
function setPeriods() {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: CONFIG.timeZone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  const today = parts.slice(0, 10);
  P.day = { label: 'Today', sub: fmtDate(today), from: today, to: today };
  P.month = { label: 'This month', sub: fmtMonth(today), from: today.slice(0, 7) + '-01', to: today.slice(0, 7) + '-31' };
  P.all = { label: 'Overall', sub: YEAR || 'All reports', from: '', to: '9999' };
}
const inPeriod = (r, p) => p === 'all' || (r.iso && r.iso >= P[p].from && r.iso <= P[p].to);

/* ---------- 3. filter + group ---------- */
function rowsFor({ district = true, period = true } = {}) {
  return ALL.filter(r =>
    (!district || !state.district || r.district === state.district) &&
    (!state.dept || r.dept === state.dept) &&
    (!period || inPeriod(r, state.period)));
}
function group(rows, key) {
  const m = new Map();
  for (const r of rows) {
    let g = m.get(r[key]);
    if (!g) m.set(r[key], g = { name: r[key], reports: 0, seedlings: 0, rows: [] });
    g.reports++; g.seedlings += r.seedlings; g.rows.push(r);
  }
  return [...m.values()];
}
const total = rows => rows.reduce((s, r) => s + r.seedlings, 0);
const bySeed = (a, b) => b.seedlings - a.seedlings || a.name.localeCompare(b.name);

/* ---------- 4. draw ---------- */
function render() {
  // tiles (respond to district + department; tapping one sets the period)
  const base = rowsFor({ period: false });
  $('tiles').innerHTML = ['day', 'month', 'all'].map(p => {
    const rs = base.filter(r => inPeriod(r, p));
    return `<button type="button" class="tile ${state.period === p ? 'on' : ''}" data-p="${p}" aria-pressed="${state.period === p}">
      <div class="tile-label">${ICON[p]}${P[p].label}</div>
      <div class="tile-value" data-p="${p}" data-v="${total(rs)}">${fmt(total(rs))}</div>
      <div class="tile-sub">${esc(P[p].sub)} · ${fmt(rs.length)} report${rs.length === 1 ? '' : 's'}</div>
    </button>`;
  }).join('');

  document.querySelectorAll('.tile-value').forEach(countUp);

  const rows = rowsFor();
  cells('listDistrict', group(rowsFor({ district: false }), 'district').sort(bySeed), true);
  lines('listDept', group(rows, 'dept').sort(bySeed));
  matrix(rows);
  chips();
  syncSheet();
}

// small line icons for the four tiles
const svg = d => `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
const ICON = {
  day: svg('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M5 5l1.5 1.5M17.5 17.5L19 19M19 5l-1.5 1.5M6.5 17.5L5 19"/>'),
  month: svg('<rect x="3" y="5" width="18" height="16" rx="3"/><path d="M3 10h18M8 3v4M16 3v4"/>'),
  all: svg('<path d="M12 21V11"/><path d="M12 11C12 7 9 4 4 4c0 5 3 8 8 7z"/><path d="M12 14c0-3 3-6 8-6 0 5-3 7-8 6z"/>')
};

// numbers roll from the previous value to the new one
const shown = {};
function countUp(el) {
  const to = +el.dataset.v, from = shown[el.dataset.p] ?? 0;
  shown[el.dataset.p] = to;
  if (from === to || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const t0 = performance.now(), ms = 650;
  const step = t => {
    if (!el.isConnected) return;
    const k = Math.min(1, (t - t0) / ms), e = 1 - Math.pow(1 - k, 3);
    el.textContent = fmt(from + (to - from) * e);
    if (k < 1) requestAnimationFrame(step);
  };
  el.textContent = fmt(from);
  requestAnimationFrame(step);
}

const EMPTY = '<p class="empty">No palm dibbling geotagged for this selection.</p>';

// keeps long lists short: first N items + a "Show all" button
function clip(id, items) {
  const n = LIMIT[id];
  const selAt = id === 'listDistrict' ? items.findIndex(x => x.name === state.district) : -1;
  const open = showAll[id] || selAt >= n;
  return {
    view: open ? items : items.slice(0, n),
    more: items.length > n ? `<button type="button" class="more" data-more="${id}">${open ? 'Show less' : `Show all ${items.length}`}</button>` : ''
  };
}

function cells(id, items, pick) {
  const { view, more } = clip(id, items);
  const tag = pick ? 'button' : 'div';
  $(id).innerHTML = items.length ? view.map((x, i) =>
    `<${tag} class="cell r${i + 1} ${pick && x.name === state.district ? 'sel' : ''}" ${pick ? `type="button" data-d="${esc(x.name)}"` : ''} title="${fmt(x.reports)} report(s)">
      <span class="cell-name"><span class="rank">${i + 1}</span>${esc(x.name)}</span><span class="cell-val">${fmt(x.seedlings)}</span>
    </${tag}>`).join('') + more : EMPTY;
}

function lines(id, items) {
  const { view, more } = clip(id, items);
  $(id).innerHTML = items.length ? view.map((x, i) =>
    `<div class="line r${i + 1}" title="${fmt(x.reports)} report(s)"><span class="line-name"><span class="rank">${i + 1}</span>${esc(x.name)}</span><span class="line-val">${fmt(x.seedlings)}</span></div>`).join('') + more : EMPTY;
}

// districts are collapsed; tap one to see its departments
function matrix(rows) {
  const d = group(rows, 'district').sort(bySeed);
  lastDistricts = d.map(x => x.name);
  const isOpen = x => openD.has(x.name) || d.length === 1;
  const cut = clip('matrix', d);
  $('matrixMore').innerHTML = cut.more;
  $('matrix').tBodies[0].innerHTML = d.length ? cut.view.map((x, i) =>
    `<tr class="d r${i + 1} ${isOpen(x) ? 'open' : ''}" data-d="${esc(x.name)}" tabindex="0" role="button" aria-expanded="${isOpen(x)}"><td class="sn"><span class="rank">${i + 1}</span></td><td>${esc(x.name)}</td><td class="num">${fmt(x.reports)}</td><td class="num">${fmt(x.seedlings)}</td></tr>` +
    (isOpen(x) ? group(x.rows, 'dept').sort(bySeed).map(s =>
      `<tr class="s"><td></td><td>${esc(s.name)}</td><td class="num">${fmt(s.reports)}</td><td class="num">${fmt(s.seedlings)}</td></tr>`).join('') : '')
  ).join('') : '<tr><td colspan="4" class="empty">No palm dibbling geotagged for this selection.</td></tr>';
  $('matrix').tFoot.innerHTML = d.length ? `<tr><td></td><td>Total</td><td class="num">${fmt(rows.length)}</td><td class="num">${fmt(total(rows))}</td></tr>` : '';
  const all = d.length > 0 && d.every(isOpen);
  $('toggleAll').textContent = all ? 'Collapse all' : 'Expand all';
  $('toggleAll').hidden = d.length < 2;
}

function chips() {
  const c = [];
  if (state.period !== 'all') c.push(['period', `${P[state.period].label}: ${P[state.period].sub}`]);
  if (state.district) c.push(['district', state.district]);
  if (state.dept) c.push(['dept', state.dept]);
  $('chips').innerHTML = c.length
    ? c.map(([k, t]) => `<button type="button" class="chip" data-k="${k}" aria-label="Remove filter ${esc(t)}">${esc(t)}<span aria-hidden="true">×</span></button>`).join('') +
      (c.length > 1 ? '<button type="button" class="chip clear" data-k="all">Clear all</button>' : '')
    : '<span class="chip muted">Showing all districts and departments</span>';
  $('filterCount').hidden = !c.length;
  $('filterCount').textContent = c.length;
}

/* ---------- 5. filter sheet ---------- */
function syncSheet() {
  $('fDistrict').value = state.district;
  $('fDept').value = state.dept;
  document.querySelectorAll('#fPeriod button').forEach(b => b.classList.toggle('on', b.dataset.p === state.period));
}
function openSheet(open) {
  $('sheet').classList.toggle('open', open);
  $('sheet').setAttribute('aria-hidden', !open);
  document.body.classList.toggle('locked', open);
  const b = $('backdrop');
  if (open) { b.hidden = false; requestAnimationFrame(() => b.classList.add('show')); $('closeFilters').focus({ preventScroll: true }); }
  else { b.classList.remove('show'); setTimeout(() => { b.hidden = true; }, 250); $('openFilters').focus({ preventScroll: true }); }
}
const set = patch => { Object.assign(state, patch); render(); };

function init() {
  setPeriods();
  $('pageTitle').textContent = document.title = CONFIG.title;
  const last = ALL.reduce((a, r) => r.iso > a ? r.iso : a, '');
  $('asOn').textContent = `Updated on ${CONFIG.updatedOn || fmtDate(last)}`;

  const opts = (key, all) => `<option value="">${all}</option>` +
    [...new Set(ALL.map(r => r[key]))].sort().map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  $('fDistrict').innerHTML = opts('district', 'All districts');
  $('fDept').innerHTML = opts('dept', 'All departments');

  $('fDistrict').onchange = e => set({ district: e.target.value });
  $('fDept').onchange = e => set({ dept: e.target.value });
  $('fPeriod').onclick = e => { const b = e.target.closest('button'); if (b) set({ period: b.dataset.p }); };
  $('tiles').onclick = e => { const b = e.target.closest('.tile'); if (b) set({ period: state.period === b.dataset.p ? 'all' : b.dataset.p }); };
  $('listDistrict').onclick = e => { const b = e.target.closest('[data-d]'); if (b) set({ district: state.district === b.dataset.d ? '' : b.dataset.d }); };
  document.querySelector('main').addEventListener('click', e => { const b = e.target.closest('[data-more]'); if (b) { showAll[b.dataset.more] = !showAll[b.dataset.more]; render(); } });
  const toggleRow = e => { const tr = e.target.closest('tr.d'); if (!tr) return; const n = tr.dataset.d; openD.has(n) ? openD.delete(n) : openD.add(n); render(); };
  $('matrix').onclick = toggleRow;
  $('matrix').onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleRow(e); } };
  $('toggleAll').onclick = () => { const all = lastDistricts.every(n => openD.has(n)); lastDistricts.forEach(n => all ? openD.delete(n) : openD.add(n)); render(); };
  $('chips').onclick = e => {
    const b = e.target.closest('[data-k]'); if (!b) return;
    if (b.dataset.k === 'all') set({ district: '', dept: '', period: 'all' });
    else set({ [b.dataset.k]: b.dataset.k === 'period' ? 'all' : '' });
  };
  $('clearFilters').onclick = () => set({ district: '', dept: '', period: 'all' });
  $('openFilters').onclick = () => openSheet(true);
  $('closeFilters').onclick = $('backdrop').onclick = () => openSheet(false);
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('sheet').classList.contains('open')) openSheet(false); });

  $('status').hidden = true;
  $('dash').hidden = false;
  $('dash').classList.add('intro');
  setTimeout(() => $('dash').classList.remove('intro'), 1200);
  render();
}

load();
