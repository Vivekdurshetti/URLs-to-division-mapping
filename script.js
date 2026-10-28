// ========================================
// CONFIG
// ========================================
const FIELD_ALIASES = {
    id:            ['id', 'item id'],
    language:      ['language', 'lang'],
    path:          ['path', 'item path'],
    url:           ['page urls', 'page url', 'url', 'urls'],
    division:      ['division names', 'division name', 'divisions', 'division'],
    subDivision:   ['sub-division', 'sub division', 'sub divisions', 'sub-divisions', 'subdivision'],
    pageType:      ['page types', 'page type'],
    templateName:  ['template name', 'template'],
    createdDate:   ['page cretaed date', 'page created date', 'created date', 'created'],
    pageName:      ['page name', 'name'],
    branchTemplate:['branch template'],
    status:        ['page status', 'status', 'edge status'],
    businessArea:  ['business areas', 'business area'],
    createdBy:     ['page created by', 'created by'],
    hasRedirect:   ['has redirect'],
    redirectType:  ['redirect type'],
    neverPublish:  ['never publish flag', 'never publish']
};
const DEFAULT_VISIBLE_KEYS = ['id', 'language', 'path', 'url', 'division', 'subDivision', 'pageType', 'pageName', 'status', 'businessArea', 'hasRedirect', 'redirectType', 'neverPublish'];
const SOURCE_ORDER = ['XP', 'XM'];
const SOURCE_LABEL = { XP: 'XP pages', XM: 'XM pages' };
const LOOKUP_PAGE_SIZE = 15;
const BROWSE_PAGE_SIZE = 25;
const NONE_LABEL = '(No value)';

// ========================================
// STATE
// ========================================
let records = [];                 // every row from every loaded source
let sourceMeta = {};              // { XP: {file, sheet, count, headers:[{lower,label,key}]} }
let urlIndex = new Map();         // normalized url -> [records]
let headerList = [];              // union of headers across sources [{lower,label,key}]
let visibleCols = new Set();      // lower-case header names shown in Browse
let knownHeaderSet = new Set();   // headers already offered, so new ones get defaults

let lookupResults = [];
let lookupVisible = [];
let lookupPage = 1;

let browseVisible = [];
let browsePage = 1;
const browseFilters = { source: '', division: '', subDivision: '', language: '', pageType: '', quality: '', search: '' };

// ========================================
// HELPERS
// ========================================
const $ = id => document.getElementById(id);

function esc(value) {
    return String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function cleanValue(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
}

function normalizeUrl(url) {
    return cleanValue(url).replace(/\/$/, '').toLowerCase();
}

function splitMulti(value) {
    return cleanValue(value).split(/[|;]/).map(s => s.trim()).filter(Boolean);
}

function fmt(n) { return Number(n).toLocaleString(); }

function csvEscape(v) {
    return '"' + String(v ?? '').replace(/"/g, '""') + '"';
}

function downloadCsv(filename, headers, rows) {
    const lines = [headers.map(csvEscape).join(',')];
    rows.forEach(r => lines.push(r.map(csvEscape).join(',')));
    const blob = new Blob(['﻿' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function downloadXlsx(filename, sheetName, headers, rows) {
    const data = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, filename);
}

function showToast(message) {
    const toast = $('toast');
    toast.innerText = message;
    toast.classList.add('show');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => toast.classList.remove('show'), 2500);
}

function copyToClipboard(text, okMessage) {
    navigator.clipboard.writeText(text)
        .then(() => showToast(okMessage))
        .catch(err => { console.error('Copy failed:', err); showToast('Copy failed'); });
}

function setStatus(message, warn) {
    const el = $('uploadStatus');
    el.textContent = message;
    el.classList.toggle('warn', !!warn);
}

function sourceClass(src) { return src === 'XP' ? 'xp' : src === 'XM' ? 'xm' : 'other'; }
function sourceBadge(src) { return `<span class="src-badge ${sourceClass(src)}">${esc(src)}</span>`; }

// ========================================
// FILE LOADING (XP + XM, one workbook or several files)
// ========================================
const fileInput = $('fileInput');
const dropZone = $('dropZone');

fileInput.addEventListener('change', e => { handleFiles(e.target.files); fileInput.value = ''; });
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
});

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        const isExcel = /\.xlsx?$/i.test(file.name);
        reader.onerror = () => reject(reader.error);
        reader.onload = e => resolve({ file, isExcel, data: e.target.result });
        if (isExcel) reader.readAsArrayBuffer(file);
        else reader.readAsText(file);
    });
}

async function handleFiles(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;

    const messages = [];
    let warn = false;

    for (const file of files) {
        try {
            const { isExcel, data } = await readFile(file);
            if (isExcel) {
                const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
                const names = workbook.SheetNames;
                names.forEach(sheetName => {
                    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: '' });
                    let src = detectSource(sheetName);
                    if (!src && names.length === 1) src = detectSource(file.name) || $('defaultSource').value;
                    if (!src) { messages.push(`Skipped tab "${sheetName}" (not XP or XM)`); warn = true; return; }
                    const result = ingestRows(rows, src, file.name, sheetName);
                    if (result.error) { messages.push(`${file.name} › ${sheetName}: ${result.error}`); warn = true; }
                    else {
                        messages.push(`${SOURCE_LABEL[src]}: ${fmt(result.count)} rows` + (result.warnNoUrl ? ' (no "Page URLs" column, so URL lookup can\'t use it)' : ''));
                        if (result.warnNoUrl) warn = true;
                    }
                });
            } else {
                const rows = parseCSV(data);
                const src = detectSource(file.name) || $('defaultSource').value;
                const result = ingestRows(rows, src, file.name, 'CSV');
                if (result.error) { messages.push(`${file.name}: ${result.error}`); warn = true; }
                else {
                    messages.push(`${SOURCE_LABEL[src]}: ${fmt(result.count)} rows` + (result.warnNoUrl ? ' (no "Page URLs" column, so URL lookup can\'t use it)' : ''));
                    if (result.warnNoUrl) warn = true;
                }
            }
        } catch (err) {
            console.error(err);
            messages.push(`${file.name}: could not be read`);
            warn = true;
        }
    }

    rebuild();
    setStatus(messages.join(' · '), warn);
}

function detectSource(name) {
    const n = String(name || '').toLowerCase();
    const xm = /(^|[^a-z])xm([^a-z]|$)/.test(n);
    const xp = /(^|[^a-z])xp([^a-z]|$)/.test(n);
    if (xm && !xp) return 'XM';
    if (xp && !xm) return 'XP';
    return null;
}

// Minimal RFC-4180 style CSV parser (handles quoted commas, quotes and newlines)
function parseCSV(text) {
    const rows = [];
    let row = [], cell = '', inQuotes = false;
    text = text.replace(/^﻿/, '');
    for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (inQuotes) {
            if (c === '"') {
                if (text[i + 1] === '"') { cell += '"'; i++; }
                else inQuotes = false;
            } else cell += c;
        } else if (c === '"') inQuotes = true;
        else if (c === ',') { row.push(cell); cell = ''; }
        else if (c === '\n' || c === '\r') {
            if (c === '\r' && text[i + 1] === '\n') i++;
            row.push(cell); cell = '';
            rows.push(row); row = [];
        } else cell += c;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
}

function aliasKey(lowerHeader) {
    for (const [key, aliases] of Object.entries(FIELD_ALIASES)) {
        if (aliases.includes(lowerHeader)) return key;
    }
    return null;
}

function findHeaderRow(rows) {
    const limit = Math.min(rows.length, 10);
    for (let i = 0; i < limit; i++) {
        const known = (rows[i] || []).filter(h => aliasKey(cleanValue(h).toLowerCase())).length;
        if (known >= 2) return i;
    }
    return 0;
}

function ingestRows(rows, src, fileName, sheetName) {
    if (!rows || rows.length < 2) return { error: 'empty or invalid data' };

    const headerRowIdx = findHeaderRow(rows);
    const rawHeaders = (rows[headerRowIdx] || []).map(h => cleanValue(h));
    if (!rawHeaders.some(Boolean)) return { error: 'no header row found' };

    // Unique lower-case keys per column (duplicate header names get a suffix)
    const seen = {};
    const cols = rawHeaders.map((label, i) => {
        if (!label) return null;
        let lower = label.toLowerCase();
        if (seen[lower]) { seen[lower]++; lower = `${lower} (${seen[lower]})`; label = `${label} (${seen[lower]})`; }
        else seen[lower] = 1;
        return { index: i, lower, label, key: aliasKey(lower) };
    }).filter(Boolean);

    const hasUrlCol = cols.some(c => c.key === 'url');
    const added = [];

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || !row.some(v => cleanValue(v) !== '')) continue;

        const raw = {};
        const rec = { source: src, raw };
        cols.forEach(c => {
            const v = cleanValue(row[c.index]);
            raw[c.lower] = v;
            if (c.key && !rec[c.key]) rec[c.key] = v;   // first non-empty column wins per field
        });
        Object.keys(FIELD_ALIASES).forEach(k => { if (rec[k] === undefined) rec[k] = ''; });

        rec.rawUrl = rec.url;
        rec.url = normalizeUrl(rec.url);
        rec.divisions = splitMulti(rec.division);
        rec.subDivisions = splitMulti(rec.subDivision);
        rec.businessAreas = splitMulti(rec.businessArea);
        added.push(rec);
    }

    // Re-loading a source replaces its previous rows
    records = records.filter(r => r.source !== src);
    records.push(...added);
    sourceMeta[src] = { file: fileName, sheet: sheetName, count: added.length, headers: cols.map(c => ({ lower: c.lower, label: c.label, key: c.key })), hasUrlCol };

    if (!hasUrlCol) return { count: added.length, error: null, warnNoUrl: true };
    return { count: added.length, error: null };
}

function removeSource(src) {
    records = records.filter(r => r.source !== src);
    delete sourceMeta[src];
    rebuild();
    setStatus(`Removed ${SOURCE_LABEL[src] || src}.`, false);
}

// ========================================
// REBUILD DERIVED DATA
// ========================================
function rebuild() {
    // URL index
    urlIndex = new Map();
    records.forEach(r => {
        if (!r.url) return;
        if (!urlIndex.has(r.url)) urlIndex.set(r.url, []);
        urlIndex.get(r.url).push(r);
    });

    // Quality flags
    records.forEach(r => {
        const matches = r.url ? urlIndex.get(r.url) : [];
        r.flags = {
            missingDivision: r.divisions.length === 0,
            missingUrl: !r.url,
            dupUrl: !!r.url && matches.filter(m => m.source === r.source).length > 1,
            both: !!r.url && new Set(matches.map(m => m.source)).size > 1
        };
    });

    // Union of headers across sources, in source order
    const known = new Map();
    sourcesInOrder().forEach(src => {
        (sourceMeta[src].headers || []).forEach(h => { if (!known.has(h.lower)) known.set(h.lower, h); });
    });
    headerList = Array.from(known.values());

    // Visible columns: keep the user's choices; new headers get sensible defaults
    if (!headerList.length) knownHeaderSet = new Set();
    const previous = visibleCols;
    visibleCols = new Set();
    headerList.forEach(h => {
        const isNew = !knownHeaderSet.has(h.lower);
        const show = isNew ? !!(h.key && DEFAULT_VISIBLE_KEYS.includes(h.key)) : previous.has(h.lower);
        if (show) visibleCols.add(h.lower);
    });
    knownHeaderSet = new Set(headerList.map(h => h.lower));

    renderSourceChips();
    populateBrowseFilters();
    renderColumnPicker();
    enableControls();
    renderDashboard();

    lookupResults = [];
    lookupVisible = [];
    renderLookupEmpty();

    browsePage = 1;
    renderBrowse();
}

function sourcesInOrder() {
    const keys = Object.keys(sourceMeta);
    return keys.sort((a, b) => {
        const ia = SOURCE_ORDER.indexOf(a), ib = SOURCE_ORDER.indexOf(b);
        return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
    });
}

function renderSourceChips() {
    const html = sourcesInOrder().map(src => {
        const m = sourceMeta[src];
        const note = m.hasUrlCol ? '' : ' · no "Page URLs" column';
        return `<span class="chip">
            <span class="dot ${sourceClass(src)}"></span>
            <strong>${esc(SOURCE_LABEL[src] || src)}</strong>
            <span class="meta">${fmt(m.count)} rows · ${esc(m.file)}${m.sheet !== 'CSV' ? ' › ' + esc(m.sheet) : ''}${esc(note)}</span>
            <button data-remove="${esc(src)}" title="Remove ${esc(SOURCE_LABEL[src] || src)}" aria-label="Remove ${esc(SOURCE_LABEL[src] || src)}">×</button>
        </span>`;
    }).join('');
    $('sourceChips').innerHTML = html;
}

$('sourceChips').addEventListener('click', e => {
    const btn = e.target.closest('[data-remove]');
    if (btn) removeSource(btn.dataset.remove);
});

function enableControls() {
    const has = records.length > 0;
    ['lookupBtn', 'clearBtn', 'downloadExcelBtn', 'downloadCsvBtn', 'filterInput', 'searchInput', 'browseExcelBtn', 'browseCsvBtn']
        .forEach(id => { $(id).disabled = !has; });
}

// ========================================
// TABS
// ========================================
function switchTab(name) {
    document.querySelectorAll('.tab').forEach(t => t.setAttribute('aria-selected', String(t.dataset.tab === name)));
    document.querySelectorAll('.tab-panel').forEach(p => { p.hidden = p.id !== 'tab-' + name; });
}
document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => switchTab(t.dataset.tab)));

// ========================================
// DASHBOARD
// ========================================
function countBy(getValues, limit) {
    const map = new Map();
    records.forEach(r => {
        let values = getValues(r);
        if (!values.length) values = [NONE_LABEL];
        values.forEach(v => {
            if (!map.has(v)) map.set(v, { label: v, XP: 0, XM: 0, other: 0, total: 0 });
            const e = map.get(v);
            if (r.source === 'XP' || r.source === 'XM') e[r.source]++; else e.other++;
            e.total++;
        });
    });
    const all = Array.from(map.values()).sort((a, b) => b.total - a.total || a.label.localeCompare(b.label));
    return { items: limit ? all.slice(0, limit) : all, totalDistinct: all.length };
}

function barChartHTML(title, subtitle, data, filterKey) {
    if (!data.items.length) return '';
    const max = Math.max(...data.items.map(i => i.total), 1);
    const rows = data.items.map(i => {
        const isNone = i.label === NONE_LABEL;
        const seg = (n, cls) => n ? `<span class="seg ${cls}" style="width:${(n / max * 100).toFixed(2)}%" title="${esc(cls === 'xp' ? 'XP' : cls === 'xm' ? 'XM' : 'Other')}: ${fmt(n)}"></span>` : '';
        const inner = `
            <span class="bar-label${isNone ? ' none' : ''}" title="${esc(i.label)}">${esc(i.label)}</span>
            <span class="bar-track">${seg(i.XP, 'xp')}${seg(i.XM, 'xm')}${seg(i.other, 'other')}</span>
            <span class="bar-count">${fmt(i.total)}</span>`;
        return filterKey && !isNone
            ? `<button class="bar-row" data-filter="${esc(filterKey)}" data-value="${esc(i.label)}" title="Show these pages in Browse">${inner}</button>`
            : `<div class="bar-row">${inner}</div>`;
    }).join('');
    const more = data.totalDistinct > data.items.length
        ? `<div class="more-note">+ ${fmt(data.totalDistinct - data.items.length)} more (use the Browse tab filters for the full list)</div>` : '';
    return `<div class="card"><h2 class="card-title">${esc(title)}<span class="card-sub">${esc(subtitle || '')}</span></h2><div class="bars">${rows}</div>${more}</div>`;
}

function yearOf(value) {
    const m = String(value || '').match(/(19|20)\d{2}/);
    return m ? [m[0]] : [];
}

function renderDashboard() {
    const host = $('dashboardContent');
    if (!records.length) {
        host.innerHTML = '<div class="card empty-state">Upload your mapping file to see the dashboard.</div>';
        return;
    }

    const total = records.length;
    const bySource = s => records.filter(r => r.source === s).length;
    const distinct = fn => new Set(records.flatMap(fn).filter(Boolean)).size;
    const withDivision = records.filter(r => !r.flags.missingDivision).length;
    const coverage = Math.round(withDivision / total * 100);
    const overlapUrls = Array.from(urlIndex.values()).filter(list => new Set(list.map(r => r.source)).size > 1).length;
    const sources = sourcesInOrder();

    const kpi = (label, value, sub) => `<div class="kpi"><h3>${esc(label)}</h3><p>${value}</p>${sub ? `<small>${esc(sub)}</small>` : ''}</div>`;
    let html = '<div class="kpis">'
        + kpi('Total pages', fmt(total))
        + sources.map(s => kpi(SOURCE_LABEL[s] || s, fmt(bySource(s)))).join('')
        + kpi('Divisions', fmt(distinct(r => r.divisions)))
        + kpi('Languages', fmt(distinct(r => [r.language])))
        + kpi('Page types', fmt(distinct(r => [r.pageType])))
        + kpi('Division coverage', coverage + '%', `${fmt(total - withDivision)} without division`)
        + (sources.length > 1 ? kpi('URLs in both', fmt(overlapUrls), 'XP and XM') : '')
        + '</div>';

    html += '<div class="legend">'
        + sources.map(s => `<span><i class="dot ${sourceClass(s)}" style="display:inline-block;width:10px;height:10px;border-radius:50%"></i>${esc(SOURCE_LABEL[s] || s)}</span>`).join('')
        + '<span>Click a bar to open those pages in the Browse tab</span></div>';

    html += '<div class="chart-grid">';
    html += barChartHTML('Pages by division', 'top 20', countBy(r => r.divisions, 20), 'division');
    html += barChartHTML('Pages by page type', 'top 12', countBy(r => r.pageType ? [r.pageType] : [], 12), 'pageType');
    html += barChartHTML('Pages by sub-division', 'top 12', countBy(r => r.subDivisions, 12), 'subDivision');
    html += barChartHTML('Pages by language', '', countBy(r => r.language ? [r.language] : [], 12), 'language');
    if (records.some(r => r.businessAreas.length)) html += barChartHTML('Pages by business area', 'top 12', countBy(r => r.businessAreas, 12), null);
    if (records.some(r => r.status)) html += barChartHTML('Pages by status', '', countBy(r => r.status ? [r.status] : [], 12), null);
    if (records.some(r => r.neverPublish)) html += barChartHTML('Never Publish flag', '', countBy(r => r.neverPublish ? [r.neverPublish] : [], 12), null);
    if (records.some(r => r.hasRedirect)) html += barChartHTML('Pages with a redirect', '', countBy(r => r.hasRedirect ? [r.hasRedirect] : [], 12), null);
    if (records.some(r => r.redirectType)) html += barChartHTML('Redirect type', '', countBy(r => r.redirectType ? [r.redirectType] : [], 12), null);
    if (records.some(r => r.createdBy)) html += barChartHTML('Pages by creator', 'top 12', countBy(r => r.createdBy ? [r.createdBy] : [], 12), null);
    if (records.some(r => r.createdDate)) {
        const years = countBy(r => yearOf(r.createdDate), 0);
        years.items.sort((a, b) => a.label.localeCompare(b.label));
        html += barChartHTML('Pages by created year', '', years, null);
    }
    if (records.some(r => r.templateName)) html += barChartHTML('Pages by template', 'top 12', countBy(r => r.templateName ? [r.templateName] : [], 12), null);
    html += '</div>';

    // Data quality
    const q = k => records.filter(r => r.flags[k]).length;
    const qItem = (label, count, filter) => `
        <div class="q-item ${count ? 'warn' : 'ok'}">
            <div><strong>${fmt(count)}</strong><span>${esc(label)}</span></div>
            ${count ? `<button data-quality="${esc(filter)}">View</button>` : ''}
        </div>`;
    html += `<div class="card" style="margin-bottom:1.5rem"><h2 class="card-title">Data quality</h2><div class="quality-grid">`
        + qItem('Pages without a division', q('missingDivision'), 'missingDivision')
        + qItem('Pages without a URL', q('missingUrl'), 'missingUrl')
        + qItem('Duplicate URLs within a sheet', q('dupUrl'), 'dupUrl')
        + (sources.length > 1 ? qItem('Pages whose URL is in both XP and XM', q('both'), 'both') : '')
        + '</div></div>';

    // Field completeness
    const rowsHtml = headerList.map(h => {
        const cells = sources.map(s => {
            const list = records.filter(r => r.source === s);
            if (!sourceMeta[s].headers.some(x => x.lower === h.lower)) return `<td style="color:#94A3B8">not in sheet</td>`;
            const filled = list.filter(r => (r.raw[h.lower] || '') !== '').length;
            const pct = list.length ? Math.round(filled / list.length * 100) : 0;
            return `<td><span class="fill-bar"><i class="${sourceClass(s)}" style="width:${pct}%"></i></span>${pct}% <span style="color:#94A3B8">(${fmt(filled)})</span></td>`;
        }).join('');
        return `<tr><td>${esc(h.label)}</td>${cells}</tr>`;
    }).join('');
    html += `<div class="card"><h2 class="card-title">Field completeness<span class="card-sub">share of rows with a value, per field</span></h2>
        <div class="table-container"><table class="fill-table"><thead><tr><th>Field</th>${sources.map(s => `<th>${esc(SOURCE_LABEL[s] || s)}</th>`).join('')}</tr></thead><tbody>${rowsHtml}</tbody></table></div></div>`;

    host.innerHTML = html;
}

$('dashboardContent').addEventListener('click', e => {
    const bar = e.target.closest('[data-filter]');
    if (bar) { openBrowse({ [bar.dataset.filter]: bar.dataset.value }); return; }
    const q = e.target.closest('[data-quality]');
    if (q) openBrowse({ quality: q.dataset.quality });
});

// ========================================
// LOOKUP
// ========================================
function renderLookupEmpty() {
    $('resultsBody').innerHTML = `<tr><td colspan="9" style="text-align:center;padding:3rem;color:#94A3B8">${records.length ? 'Paste URLs and click "Process URLs".' : 'Awaiting file upload and URL processing...'}</td></tr>`;
    $('stats').innerHTML = statBoxes(0, 0, 0, 0, 0, 0);
    $('pagination').innerHTML = '';
}

function statBoxes(total, found, missing, dups, inXp, inXm) {
    const box = (h, v, color) => `<div class="stat-box"><h3>${h}</h3><p${color ? ` style="color:${color}"` : ''}>${fmt(v)}</p></div>`;
    return box('Total Queried', total) + box('Matches Found', found, '#10B981') + box('Missing', missing, '#EF4444')
        + box('Duplicates', dups, '#F59E0B') + box('Matched in XP', inXp, 'var(--xp)') + box('Matched in XM', inXm, 'var(--xm)');
}

$('lookupBtn').addEventListener('click', () => {
    const urls = $('urlInput').value.split(/\n/).map(normalizeUrl).filter(Boolean);
    if (!urls.length) { alert('Please paste URLs to look up.'); return; }

    const onlySource = $('lookupSource').value;
    const seen = new Set();
    lookupResults = [];

    urls.forEach(url => {
        const duplicate = seen.has(url);
        seen.add(url);
        let matches = urlIndex.get(url) || [];
        if (onlySource) matches = matches.filter(m => m.source === onlySource);

        if (!matches.length) {
            lookupResults.push({ originalUrl: url, duplicate, notFound: true, source: '', id: '', language: '', path: '', division: '', subDivision: '', pageType: '' });
        } else {
            matches.forEach(m => lookupResults.push({
                originalUrl: url, duplicate, notFound: false, source: m.source,
                id: m.id, language: m.language, path: m.path, division: m.division, subDivision: m.subDivision, pageType: m.pageType
            }));
        }
    });

    lookupPage = 1;
    renderLookup();
    renderLookupStats(urls.length);
});

function renderLookupStats(totalQueried) {
    const foundUrls = new Set(lookupResults.filter(r => !r.notFound).map(r => r.originalUrl));
    const missingUrls = new Set(lookupResults.filter(r => r.notFound).map(r => r.originalUrl));
    const distinctQueried = new Set(lookupResults.map(r => r.originalUrl)).size;
    const inSrc = s => new Set(lookupResults.filter(r => r.source === s).map(r => r.originalUrl)).size;
    // Duplicates = queried lines beyond the first occurrence of each URL
    $('stats').innerHTML = statBoxes(totalQueried, foundUrls.size, missingUrls.size, totalQueried - distinctQueried, inSrc('XP'), inSrc('XM'));
}

$('filterInput').addEventListener('input', () => { lookupPage = 1; renderLookup(); });
$('searchInput').addEventListener('input', () => { lookupPage = 1; renderLookup(); });

function renderLookup() {
    const filterValue = $('filterInput').value.toLowerCase();
    const searchValue = $('searchInput').value.toLowerCase();

    lookupVisible = lookupResults.filter(r =>
        (r.division || '').toLowerCase().includes(filterValue) &&
        ((r.originalUrl || '').toLowerCase().includes(searchValue) || (r.path || '').toLowerCase().includes(searchValue))
    );

    const start = (lookupPage - 1) * LOOKUP_PAGE_SIZE;
    const paged = lookupVisible.slice(start, start + LOOKUP_PAGE_SIZE);

    if (!paged.length) {
        $('resultsBody').innerHTML = `<tr><td colspan="9" style="text-align:center;padding:2rem;color:#6B7280">No records found.</td></tr>`;
    } else {
        $('resultsBody').innerHTML = paged.map(r => {
            let badges = '';
            if (r.notFound) badges += '<span class="status-badge status-missing">Missing</span> ';
            if (r.duplicate) badges += '<span class="status-badge status-dup">Duplicate</span>';
            if (!r.notFound && !r.duplicate) badges += '<span class="status-badge status-found">Found</span>';
            const cell = (v, extra) => `<td class="copy-cell cell"${extra || ''}>${esc(v) || '-'}</td>`;
            const safeHref = /^https?:\/\//i.test(r.originalUrl) ? r.originalUrl : '';
            return `<tr>
                <td class="copy-cell cell">${safeHref ? `<a href="${esc(safeHref)}" target="_blank" rel="noopener" data-nocopy>${esc(r.originalUrl)}</a>` : esc(r.originalUrl)}</td>
                <td>${r.source ? sourceBadge(r.source) : '-'}</td>
                <td class="copy-cell cell" style="font-family:monospace;font-size:0.8rem;color:#64748B">${esc(r.id) || '-'}</td>
                ${cell(r.language)}
                ${cell(r.path)}
                <td class="copy-cell cell" style="${r.notFound ? 'color:#EF4444;font-style:italic' : 'font-weight:500'}">${r.notFound ? 'Not Found' : (esc(r.division) || '-')}</td>
                ${cell(r.subDivision)}
                ${cell(r.pageType)}
                <td>${badges}</td>
            </tr>`;
        }).join('');
    }
    renderPager('pagination', lookupVisible.length, LOOKUP_PAGE_SIZE, lookupPage, p => { lookupPage = p; renderLookup(); });
}

function copyColumn(key) {
    if (!lookupVisible.length) { alert('No data available to copy.'); return; }
    const data = lookupVisible.map(r => r[key]).filter(v => v && v !== '-').join('\n');
    if (!data) { alert('No valid data found in this column to copy.'); return; }
    copyToClipboard(data, `Copied ${lookupVisible.length} items to clipboard`);
}
document.querySelector('#resultsTable thead').addEventListener('click', e => {
    const btn = e.target.closest('[data-copy]');
    if (btn) copyColumn(btn.dataset.copy);
});

function copyCellHandler(e) {
    if (e.target.closest('[data-nocopy]')) return;
    const td = e.target.closest('td.copy-cell');
    if (!td) return;
    const text = td.innerText.trim();
    if (!text || text === '-' || text === 'Not Found') return;
    copyToClipboard(text, `Copied: ${text.length > 30 ? text.slice(0, 30) + '...' : text}`);
}
$('resultsBody').addEventListener('click', copyCellHandler);

$('clearBtn').addEventListener('click', () => {
    $('urlInput').value = '';
    $('filterInput').value = '';
    $('searchInput').value = '';
    lookupResults = [];
    lookupVisible = [];
    renderLookupEmpty();
});

function lookupExportRows() {
    return lookupResults.map(r => [
        r.id, r.language, r.path, r.notFound ? '' : r.division, r.subDivision, r.originalUrl, r.pageType,
        r.source ? SOURCE_LABEL[r.source] || r.source : '', r.notFound ? 'Missing' : 'Found', r.duplicate ? 'Yes' : 'No'
    ]);
}
const LOOKUP_EXPORT_HEADERS = ['ID', 'Language', 'Path', 'Division Names', 'SUB-Division', 'Page URLs', 'Page Types', 'Source', 'Match Status', 'Is Duplicate'];

$('downloadExcelBtn').addEventListener('click', () => {
    if (!lookupResults.length) return alert('No results to export.');
    downloadXlsx('URL_Division_Mapping.xlsx', 'Mapped Results', LOOKUP_EXPORT_HEADERS, lookupExportRows());
});
$('downloadCsvBtn').addEventListener('click', () => {
    if (!lookupResults.length) return alert('No results to export.');
    downloadCsv('URL_Division_Mapping.csv', LOOKUP_EXPORT_HEADERS, lookupExportRows());
});

// ========================================
// BROWSE ALL FIELDS
// ========================================
const FILTER_SELECTS = [
    { key: 'source', label: 'All sources' },
    { key: 'division', label: 'All divisions' },
    { key: 'subDivision', label: 'All sub-divisions' },
    { key: 'language', label: 'All languages' },
    { key: 'pageType', label: 'All page types' }
];

function populateBrowseFilters() {
    const distinctSorted = fn => Array.from(new Set(records.flatMap(fn).filter(Boolean))).sort((a, b) => a.localeCompare(b));
    const options = {
        source: sourcesInOrder().map(s => [s, SOURCE_LABEL[s] || s]),
        division: distinctSorted(r => r.divisions).map(v => [v, v]),
        subDivision: distinctSorted(r => r.subDivisions).map(v => [v, v]),
        language: distinctSorted(r => [r.language]).map(v => [v, v]),
        pageType: distinctSorted(r => [r.pageType]).map(v => [v, v])
    };
    FILTER_SELECTS.forEach(f => {
        const sel = $('bf-' + f.key);
        const list = options[f.key];
        // Drop a stale filter value that no longer exists
        if (browseFilters[f.key] && !list.some(o => o[0] === browseFilters[f.key])) browseFilters[f.key] = '';
        sel.innerHTML = `<option value="">${esc(f.label)}</option>` + list.map(([v, l]) => `<option value="${esc(v)}">${esc(l)}</option>`).join('');
        sel.value = browseFilters[f.key];
    });
    $('bf-quality').value = browseFilters.quality;
}

function renderColumnPicker() {
    $('colPickerPanel').innerHTML = headerList.map(h =>
        `<label><input type="checkbox" data-col="${esc(h.lower)}" ${visibleCols.has(h.lower) ? 'checked' : ''}> ${esc(h.label)}</label>`
    ).join('') || '<span style="color:#94A3B8;font-size:0.85rem">No columns yet</span>';
}

$('colPickerPanel').addEventListener('change', e => {
    const cb = e.target.closest('[data-col]');
    if (!cb) return;
    if (cb.checked) visibleCols.add(cb.dataset.col); else visibleCols.delete(cb.dataset.col);
    renderBrowse();
});

function openBrowse(filters) {
    Object.keys(browseFilters).forEach(k => { browseFilters[k] = ''; });
    Object.assign(browseFilters, filters);
    FILTER_SELECTS.forEach(f => { $('bf-' + f.key).value = browseFilters[f.key]; });
    $('bf-quality').value = browseFilters.quality;
    $('bf-search').value = browseFilters.search;
    browsePage = 1;
    renderBrowse();
    switchTab('browse');
}

FILTER_SELECTS.forEach(f => $('bf-' + f.key).addEventListener('change', e => { browseFilters[f.key] = e.target.value; browsePage = 1; renderBrowse(); }));
$('bf-quality').addEventListener('change', e => { browseFilters.quality = e.target.value; browsePage = 1; renderBrowse(); });
$('bf-search').addEventListener('input', e => { browseFilters.search = e.target.value; browsePage = 1; renderBrowse(); });
$('bf-reset').addEventListener('click', () => openBrowse({}));

function applyBrowseFilters() {
    const f = browseFilters;
    const needle = f.search.trim().toLowerCase();
    return records.filter(r => {
        if (f.source && r.source !== f.source) return false;
        if (f.division && !r.divisions.includes(f.division)) return false;
        if (f.subDivision && !r.subDivisions.includes(f.subDivision)) return false;
        if (f.language && r.language !== f.language) return false;
        if (f.pageType && r.pageType !== f.pageType) return false;
        if (f.quality && !r.flags[f.quality]) return false;
        if (needle && !Object.values(r.raw).some(v => v.toLowerCase().includes(needle))) return false;
        return true;
    });
}

function shownHeaders() { return headerList.filter(h => visibleCols.has(h.lower)); }

function renderBrowse() {
    browseVisible = applyBrowseFilters();
    const cols = shownHeaders();

    $('browseCount').textContent = records.length ? `${fmt(browseVisible.length)} of ${fmt(records.length)} pages` : '';

    if (!records.length) {
        $('browseHead').innerHTML = '';
        $('browseBody').innerHTML = `<tr><td style="text-align:center;padding:3rem;color:#94A3B8">Awaiting file upload...</td></tr>`;
        $('browsePagination').innerHTML = '';
        return;
    }

    $('browseHead').innerHTML = '<tr><th>Source</th>' + cols.map(h => `<th>${esc(h.label)}</th>`).join('') + '</tr>';

    const start = (browsePage - 1) * BROWSE_PAGE_SIZE;
    const paged = browseVisible.slice(start, start + BROWSE_PAGE_SIZE);
    if (!paged.length) {
        $('browseBody').innerHTML = `<tr><td colspan="${cols.length + 1}" style="text-align:center;padding:2rem;color:#6B7280">No records match these filters.</td></tr>`;
    } else {
        $('browseBody').innerHTML = paged.map(r => {
            const cells = cols.map(h => {
                const v = r.raw[h.lower] ?? '';
                const content = /^https?:\/\//i.test(v)
                    ? `<a href="${esc(v)}" target="_blank" rel="noopener" data-nocopy>${esc(v)}</a>` : esc(v);
                return `<td class="copy-cell cell" title="${esc(v)}">${content || '-'}</td>`;
            }).join('');
            return `<tr><td>${sourceBadge(r.source)}</td>${cells}</tr>`;
        }).join('');
    }
    renderPager('browsePagination', browseVisible.length, BROWSE_PAGE_SIZE, browsePage, p => { browsePage = p; renderBrowse(); });
}

$('browseBody').addEventListener('click', copyCellHandler);

function browseExport() {
    const cols = shownHeaders();
    return {
        headers: ['Source', ...cols.map(h => h.label)],
        rows: browseVisible.map(r => [SOURCE_LABEL[r.source] || r.source, ...cols.map(h => r.raw[h.lower] ?? '')])
    };
}
$('browseExcelBtn').addEventListener('click', () => {
    if (!browseVisible.length) return alert('No rows to export.');
    const { headers, rows } = browseExport();
    downloadXlsx('Pages_Export.xlsx', 'Pages', headers, rows);
});
$('browseCsvBtn').addEventListener('click', () => {
    if (!browseVisible.length) return alert('No rows to export.');
    const { headers, rows } = browseExport();
    downloadCsv('Pages_Export.csv', headers, rows);
});

// ========================================
// PAGINATION (shared)
// ========================================
function renderPager(containerId, total, pageSize, current, onPage) {
    const totalPages = Math.ceil(total / pageSize);
    const host = $(containerId);
    host.innerHTML = '';
    if (totalPages <= 1) return;

    const mkBtn = n => {
        const b = document.createElement('button');
        b.textContent = n;
        if (n === current) b.disabled = true;
        b.addEventListener('click', () => onPage(n));
        return b;
    };
    const dots = () => { const s = document.createElement('span'); s.textContent = '...'; s.style.padding = '0.5rem'; return s; };

    const start = Math.max(1, current - 3), end = Math.min(totalPages, current + 3);
    if (start > 1) { host.appendChild(mkBtn(1)); if (start > 2) host.appendChild(dots()); }
    for (let i = start; i <= end; i++) host.appendChild(mkBtn(i));
    if (end < totalPages) { if (end < totalPages - 1) host.appendChild(dots()); host.appendChild(mkBtn(totalPages)); }
}

// ========================================
// INIT
// ========================================
renderLookupEmpty();
renderBrowse();
