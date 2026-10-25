// Backend en producción (Render)
const BACKEND_URL = 'https://consulta-eep-backend.onrender.com';

const LS_ROWS = 'eep_excel_rows';
const LS_ROUTE = 'eep_selected_route';
const LS_RESULTS = 'eep_scan_results';
const LS_LINIERO = 'eep_liniero_nombre';

let excelRows = []; // cada fila trae ahora un campo `fecha` (YYYY-MM-DD) — el día del archivo del que vino
let scanResults = {}; // { [niu]: { pagado, ultimoPago, estadoPago, fechaRegistroPago, checkedAt, ... } } — global por NIU
let pagosManuales = {}; // { [niu]: { liniero, fecha } } — "cliente mostró que pagó" pero el portal aún no lo confirma
let cortes = []; // [{ niu, nombre, direccion, ruta, liniero, tipo, fecha, subidoCortes }]
let compartidos = []; // [{ niu, numeroFactura, fecha }]
let rutaLinieros = {}; // { [ruta]: linieroName }
let selectedRoute = null;
let selectedDay = null; // 'YYYY-MM-DD' | 'general' — día activo dentro de la ruta actual
let currentFilter = 'pendiente';
let searchTerm = '';
let pendingCorteRow = null; // row waiting for corte confirmation in the modal
let editingPhone = null; // { niu, field } cuando se está editando/agregando un número

function parseSpanishDateTime(str) {
  // Expects "DD-MES-YYYY hh:mmAM/PM" e.g. "08-JUL-2026 05:28PM"
  if (!str) return null;
  const meses = { ENE: 0, FEB: 1, MAR: 2, ABR: 3, MAY: 4, JUN: 5, JUL: 6, AGO: 7, SEP: 8, OCT: 9, NOV: 10, DIC: 11 };
  const match = str.trim().match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})\s+(\d{1,2}):(\d{2})(AM|PM)$/i);
  if (!match) return null;
  const [, day, mesTxt, year, hh, mm, ampm] = match;
  const mes = meses[mesTxt.toUpperCase()];
  if (mes === undefined) return null;
  let hour = parseInt(hh, 10) % 12;
  if (ampm.toUpperCase() === 'PM') hour += 12;
  return new Date(parseInt(year, 10), mes, parseInt(day, 10), hour, parseInt(mm, 10));
}

function todayIso() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// Extrae la fecha del nombre del archivo, ej. "VILLAGARZÓN_04082026.xlsx" -> "2026-08-04"
// Formato esperado: los últimos 8 dígitos del nombre (sin extensión) son DDMMYYYY.
function extractFechaFromFilename(filename) {
  const base = String(filename || '').replace(/\.[^.]+$/, '');
  const match = base.match(/(\d{2})(\d{2})(\d{4})$/);
  if (!match) return { fecha: todayIso(), detectada: false };
  const dd = parseInt(match[1], 10);
  const mm = parseInt(match[2], 10);
  const yyyy = match[3];
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return { fecha: todayIso(), detectada: false };
  return { fecha: `${yyyy}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`, detectada: true };
}

function formatFechaCorta(fechaIso) {
  // "2026-08-04" -> "04/08"
  const [, m, d] = fechaIso.split('-');
  return `${d}/${m}`;
}

// ---------- Column mapping (flexible — works regardless of column order/casing) ----------
const HEADER_MAP = {
  municipio: 'municipio',
  nombre: 'nombre',
  direccion: 'direccion',
  niu: 'niu',
  medidor: 'medidor',
  sector: 'sector',
  'consumo kwh': 'consumoKwh',
  'valor ap': 'valorAp',
  total: 'total',
  'total - ap': 'totalMenosAp',
  'total-ap': 'totalMenosAp',
  recaudo: 'recaudo',
  'meses atrasados': 'mesesAtrasados',
  ruta: 'ruta',
  'orden ruta': 'ordenRuta',
  pendiente: 'saldoPendiente',
  observacion: 'observacion',
};

function normalizeHeader(h) {
  return String(h)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/_/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeRows(rows2d) {
  if (!Array.isArray(rows2d) || rows2d.length === 0) return [];
  const headerRow = rows2d[0].map((h) => normalizeHeader(h));
  const dataRows = rows2d.slice(1);

  return dataRows
    .map((raw) => {
      const row = {};
      const phones = []; // acumula, en orden de aparición, cualquier columna "telefono" o "celular"
      headerRow.forEach((h, idx) => {
        const value = String(raw[idx] ?? '').trim();
        if (!value) return;
        if (h === 'telefono' || h === 'celular') {
          phones.push(value);
          return;
        }
        const mapped = HEADER_MAP[h];
        if (mapped) row[mapped] = value;
      });
      // Descarta "0", "00", "000...", vacíos, etc. — no son números reales,
      // así el campo queda disponible para registrar uno válido después.
      const validPhones = phones.filter((p) => !/^0+$/.test(p));
      if (validPhones[0]) row.telefono = validPhones[0];
      if (validPhones[1]) row.celular = validPhones[1];
      return row;
    })
    .filter((r) => r.niu); // descarta filas sin NIU — inutilizables
}

function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        // header: 1 => filas como arrays posicionales, no objetos — necesario
        // para no perder columnas cuando el Excel repite el mismo encabezado.
        const rows2d = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
        resolve(rows2d);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('No se pudo leer el archivo'));
    reader.readAsArrayBuffer(file);
  });
}

// ---------- Persistence ----------
function saveRows() { localStorage.setItem(LS_ROWS, JSON.stringify(excelRows)); }
function saveResults() { localStorage.setItem(LS_RESULTS, JSON.stringify(scanResults)); }
function saveRoute() { localStorage.setItem(LS_ROUTE, selectedRoute || ''); }
function saveLiniero(nombre) { localStorage.setItem(LS_LINIERO, nombre || ''); }
function getSavedLiniero() { return localStorage.getItem(LS_LINIERO) || ''; }

function loadAll() {
  try { excelRows = JSON.parse(localStorage.getItem(LS_ROWS) || '[]'); } catch (e) { excelRows = []; }
  try { scanResults = JSON.parse(localStorage.getItem(LS_RESULTS) || '{}'); } catch (e) { scanResults = {}; }
  selectedRoute = localStorage.getItem(LS_ROUTE) || null;
}

// ---------- Helpers: NIU único ("general") ----------
// Para una lista de filas (potencialmente con el mismo NIU en varios días),
// devuelve una fila por NIU: la más reciente por fecha.
function dedupeByNiuLatest(rows) {
  const map = {};
  rows.forEach((r) => {
    const prev = map[r.niu];
    if (!prev || (r.fecha || '') > (prev.fecha || '')) map[r.niu] = r;
  });
  return Object.values(map);
}

function rowsOfRoute(ruta) {
  return excelRows.filter((r) => (r.ruta || '(sin ruta)') === ruta);
}

function generalRowsForRoute(ruta) {
  return dedupeByNiuLatest(rowsOfRoute(ruta));
}

function generalRowsAll() {
  return dedupeByNiuLatest(excelRows);
}

function datesForRoute(ruta) {
  const set = new Set(rowsOfRoute(ruta).map((r) => r.fecha).filter(Boolean));
  return Array.from(set).sort(); // ascendente
}

function defaultDayForRoute(ruta) {
  const dates = datesForRoute(ruta);
  if (dates.length === 0) return 'general';
  const today = todayIso();
  if (dates.includes(today)) return today;
  return dates[dates.length - 1]; // el más reciente disponible
}

function currentRouteRowsForDay() {
  if (!selectedRoute) return [];
  if (selectedDay === 'general') return generalRowsForRoute(selectedRoute);
  return rowsOfRoute(selectedRoute).filter((r) => r.fecha === selectedDay);
}

// ---------- View switching ----------
function showView(name) {
  ['upload', 'routes', 'detail', 'stats'].forEach((v) => {
    document.getElementById(`view-${v}`).style.display = v === name ? '' : 'none';
  });
}

// ---------- Upload screen ----------
const fileInput = document.getElementById('fileInput');
const uploadStatus = document.getElementById('uploadStatus');
const subtitleText = document.getElementById('subtitleText');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  uploadStatus.textContent = 'Leyendo archivo...';
  uploadStatus.className = 'upload-status';
  try {
    const raw = await parseExcelFile(file);
    const normalized = normalizeRows(raw);
    if (normalized.length === 0) {
      uploadStatus.textContent = 'No se encontraron filas con columna "NIU". Revisa el archivo.';
      uploadStatus.className = 'upload-status error';
      return;
    }

    const { fecha, detectada } = extractFechaFromFilename(file.name);
    const fechaLabel = fecha === todayIso() ? 'Hoy' : formatFechaCorta(fecha);
    uploadStatus.textContent = detectada
      ? `Guardando el listado en este celular (día ${fechaLabel})...`
      : `No se detectó la fecha en el nombre del archivo — se usará hoy (${fechaLabel})...`;
    const m = mergeRowsLocal(normalized, fecha);
    uploadStatus.textContent = `Listo (día ${fechaLabel}): ${m.nuevos} nuevas, ${m.actualizados} actualizadas, ${m.total} NIU en total ✅`;
    uploadStatus.className = 'upload-status ok';
    setTimeout(() => goToRoutes(), 700);
  } catch (err) {
    uploadStatus.textContent = 'Error leyendo el archivo: ' + err.message;
    uploadStatus.className = 'upload-status error';
  }
});

document.getElementById('changeFileBtn').addEventListener('click', () => {
  showView('upload');
  subtitleText.textContent = 'Sube tu listado para empezar';
});

document.getElementById('resetMonthBtn').addEventListener('click', async () => {
  const clave = window.prompt('Esta acción requiere clave de autorización:');
  if (clave === null) return; // cancelled
  if (clave !== '033093') {
    alert('Clave incorrecta.');
    return;
  }
  const confirmado = window.confirm(
    '¿Seguro que quieres resetear el mes?\n\n' +
    'Esto borra el listado de matrículas y los resultados de pago actuales (útil si se subió un archivo equivocado).\n\n' +
    'Los reportes de cortes, notificaciones y comprobantes de este y otros meses NO se pierden — siguen disponibles en Indicadores.'
  );
  if (!confirmado) return;
  try {
    excelRows = [];
    scanResults = {};
    pagosManuales = {};
    lsSet(LS_PAGOS, pagosManuales);
    selectedRoute = null;
    selectedDay = null;
    saveRows();
    saveResults();
    saveRoute();
    showView('upload');
    subtitleText.textContent = 'Sube tu listado para empezar';
  } catch (e) {
    alert('No se pudo resetear — revisa tu conexión e intenta de nuevo.');
  }
});

// ---------- Datos locales (todo queda en este celular) ----------
const LS_PAGOS = 'eep_pagos_manuales';
const LS_CORTES = 'eep_cortes';
const LS_COMPARTIDOS = 'eep_compartidos';
const LS_RUTA_LINIEROS = 'eep_ruta_linieros';
const LS_MONTH = 'eep_data_month';

function lsGet(key, fallback) {
  try {
    const v = JSON.parse(localStorage.getItem(key));
    return v === null || v === undefined ? fallback : v;
  } catch (e) { return fallback; }
}
function lsSet(key, value) { localStorage.setItem(key, JSON.stringify(value)); }

function currentMonthKey() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

// El listado, los resultados y las marcas manuales se reinician cada mes; los
// cortes y comprobantes compartidos se conservan para los reportes.
function checkMonthRollover() {
  const now = currentMonthKey();
  const stored = localStorage.getItem(LS_MONTH);
  if (stored && stored !== now) {
    [LS_ROWS, LS_RESULTS, LS_PAGOS, LS_ROUTE].forEach((k) => localStorage.removeItem(k));
  }
  localStorage.setItem(LS_MONTH, now);
}

// Combina el listado subido con lo que ya hay: cada día es su propia "foto"; el
// teléfono/celular se hereda de la última versión conocida del NIU.
function mergeRowsLocal(rows, fecha) {
  const latestByNiu = {};
  excelRows.forEach((r) => {
    if (!latestByNiu[r.niu] || (r.fecha || '') > (latestByNiu[r.niu].fecha || '')) latestByNiu[r.niu] = r;
  });
  const existingForDate = new Set(excelRows.filter((r) => r.fecha === fecha).map((r) => r.niu));
  let nuevos = 0;
  let actualizados = 0;
  const merged = rows.map((r) => {
    if (existingForDate.has(r.niu)) actualizados++; else nuevos++;
    const out = { ...r, fecha };
    const carry = latestByNiu[r.niu];
    if (!out.telefono && carry && carry.telefono) out.telefono = carry.telefono;
    if (!out.celular && carry && carry.celular) out.celular = carry.celular;
    return out;
  });
  const replaced = new Set(merged.map((r) => r.niu));
  excelRows = excelRows.filter((r) => !(r.fecha === fecha && replaced.has(r.niu))).concat(merged);
  saveRows();
  return { nuevos, actualizados, total: new Set(excelRows.map((r) => r.niu)).size };
}

async function syncSharedRows() { excelRows = lsGet(LS_ROWS, []); }
async function syncSharedResults() { scanResults = lsGet(LS_RESULTS, {}); }
async function syncPagosManuales() { pagosManuales = lsGet(LS_PAGOS, {}); }
async function syncCortes() { cortes = lsGet(LS_CORTES, []); }
async function syncCompartidos() { compartidos = lsGet(LS_COMPARTIDOS, []); }
async function syncRutaLinieros() { rutaLinieros = lsGet(LS_RUTA_LINIEROS, {}); }

// Guarda en este celular lo que devolvió el servidor al terminar una consulta.
async function applyScanResults(jobId) {
  const res = await fetch(`${BACKEND_URL}/api/scan/results/${jobId}`);
  const data = await res.json();
  const checkedAt = new Date().toISOString();
  (data.results || []).forEach((r) => {
    if (String(r.ultimoPago || '').startsWith('ERROR:')) return;
    scanResults[r.niu] = {
      pagado: r.pagado,
      ultimoPago: r.ultimoPago,
      estadoPago: r.estadoPago,
      fechaRegistroPago: r.fechaRegistroPago,
      valorPagado: r.valorPagado,
      medioPago: r.medioPago,
      entidadRecaudo: r.entidadRecaudo,
      numeroCuenta: r.numeroCuenta,
      numeroFactura: r.numeroFactura,
      checkedAt,
    };
    // Si el portal ya confirmó el pago, la marca manual deja de hacer falta.
    if (r.pagado) delete pagosManuales[r.niu];
  });
  saveResults();
  lsSet(LS_PAGOS, pagosManuales);
}

// ---------- Routes screen ----------
async function goToRoutes() {
  showView('routes');
  await syncRutaLinieros();
  renderRoutesList();
  subtitleText.textContent = `${generalRowsAll().length} matrículas cargadas`;
}

function renderRoutesList() {
  const routesList = document.getElementById('routesList');
  const counts = {};
  generalRowsAll().forEach((r) => {
    const ruta = r.ruta || '(sin ruta)';
    counts[ruta] = (counts[ruta] || 0) + 1;
  });
  const rutas = Object.keys(counts).sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));

  routesList.innerHTML = rutas
    .map((r) => {
      const liniero = rutaLinieros[r] || '';
      return `<div class="route-card">
        <div class="route-click" data-ruta="${escapeHtml(r)}">
          <span class="route-name">Ruta ${escapeHtml(r)}</span>
          <span class="route-count">${counts[r]} matrículas</span>
        </div>
        <div class="route-liniero-row">
          <input type="text" class="route-liniero-input" data-ruta="${escapeHtml(r)}" placeholder="Liniero encargado..." value="${escapeHtml(liniero)}" />
          <button class="route-liniero-save" data-ruta="${escapeHtml(r)}">Guardar</button>
        </div>
      </div>`;
    })
    .join('');

  routesList.querySelectorAll('.route-click').forEach((card) => {
    card.addEventListener('click', () => {
      selectedRoute = card.dataset.ruta;
      saveRoute();
      goToDetail();
    });
  });

  routesList.querySelectorAll('.route-liniero-save').forEach((btn) => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const ruta = btn.dataset.ruta;
      const input = routesList.querySelector(`.route-liniero-input[data-ruta="${CSS.escape(ruta)}"]`);
      const liniero = input.value.trim();
      btn.textContent = '...';
      try {
        rutaLinieros[ruta] = liniero;
        lsSet(LS_RUTA_LINIEROS, rutaLinieros);
        btn.textContent = '✅';
        setTimeout(() => { btn.textContent = 'Guardar'; }, 1200);
      } catch (err) {
        btn.textContent = 'Error';
      }
    });
  });
  routesList.querySelectorAll('.route-liniero-input').forEach((input) => {
    input.addEventListener('click', (e) => e.stopPropagation());
  });
}

// ---------- Detail screen ----------
const detailRouteTitle = document.getElementById('detailRouteTitle');
const dayTabsEl = document.getElementById('dayTabs');
const listEl = document.getElementById('list');
const statusBar = document.getElementById('statusBar');
const lastUpdateEl = document.getElementById('lastUpdate');
const scanPendingBtn = document.getElementById('scanPendingBtn');
const searchInput = document.getElementById('search');

let detailRefreshTimer = null;

document.getElementById('backToRoutesBtn').addEventListener('click', () => {
  if (detailRefreshTimer) { clearInterval(detailRefreshTimer); detailRefreshTimer = null; }
  goToRoutes();
});

function cortadoRecord(niu) {
  // most recent corte for this NIU, if any
  const matches = cortes.filter((c) => c.niu === niu);
  if (matches.length === 0) return null;
  return matches.sort((a, b) => new Date(b.fecha) - new Date(a.fecha))[0];
}

// Returns { status: 'ok'|'no'|'unk'|'notificado'|'suspendido', pagado, corteInvalido, sinConfirmar }
// `sinConfirmar` = true cuando el "ok" viene de una marca manual (liniero dijo que
// el cliente ya pagó) y el portal todavía no lo ha confirmado con un escaneo real.
function evaluarEstado(niu) {
  const scan = scanResults[niu];
  const pagadoConfirmado = !!(scan && scan.pagado);
  const manual = pagosManuales[niu];
  const pagado = pagadoConfirmado || !!manual;
  const sinConfirmar = !!manual && !pagadoConfirmado;
  const corte = cortadoRecord(niu);

  if (!corte) {
    return { status: pagado ? 'ok' : (scan ? 'no' : 'unk'), pagado, corteInvalido: false, sinConfirmar };
  }

  if (corte.tipo === 'notificacion') {
    // A notification never physically cuts service — paying makes it "al día"
    // straight away; otherwise it's its own distinct category, not "pendiente".
    return { status: pagado ? 'ok' : 'notificado', pagado, corteInvalido: false, sinConfirmar };
  }

  // Physical cut (poste/pin) — check whether the payment happened BEFORE the cut,
  // which would mean the cut was a mistake (customer had already paid) and doesn't apply.
  // (Se usa solo el pago CONFIRMADO por el portal — la marca manual no trae hora exacta.)
  const corteTimestamp = new Date(corte.fecha);
  const pagoTimestamp = scan && scan.fechaRegistroPago ? parseSpanishDateTime(scan.fechaRegistroPago) : null;

  if (pagadoConfirmado && pagoTimestamp && pagoTimestamp < corteTimestamp) {
    return { status: 'ok', pagado: true, corteInvalido: true, sinConfirmar: false };
  }

  // Valid suspension — stays classified as "suspendido" even if later paid
  // (needs manual reconnection), so we surface `pagado` separately.
  return { status: 'suspendido', pagado, corteInvalido: false, sinConfirmar };
}

function statusOf(niu) {
  return evaluarEstado(niu).status;
}

// Un NIU necesita volver a revisarse si NO está en "al día" confirmado de verdad
// por el portal — un "al día" marcado manualmente (sin confirmar) sigue entrando.
function needsScan(niu) {
  const ev = evaluarEstado(niu);
  return !(ev.status === 'ok' && !ev.sinConfirmar);
}

function compartidoRecord(niu, numeroFactura) {
  if (!numeroFactura) return null;
  return compartidos.find((c) => c.niu === niu && c.numeroFactura === numeroFactura) || null;
}

// Fila más reciente conocida de un NIU (para modales/comprobantes: siempre el dato más actual)
function generalRowForNiu(niu) {
  const rows = excelRows.filter((r) => r.niu === niu);
  if (rows.length === 0) return null;
  return rows.reduce((best, r) => (!best || (r.fecha || '') > (best.fecha || '') ? r : best), null);
}

function goToDetail() {
  selectedDay = defaultDayForRoute(selectedRoute);
  showView('detail');
  Promise.all([syncSharedResults(), syncPagosManuales(), syncCortes(), syncCompartidos()]).then(renderDetail);
  renderDetail();

  if (detailRefreshTimer) clearInterval(detailRefreshTimer);
  detailRefreshTimer = setInterval(() => {
    Promise.all([syncSharedResults(), syncPagosManuales(), syncCortes(), syncCompartidos()]).then(renderDetail);
  }, 20000); // pick up progress from other devices every 20s
}

function renderDayTabs() {
  if (!dayTabsEl) return;
  const dates = datesForRoute(selectedRoute);
  const today = todayIso();
  const tabs = dates.map((d) => ({ key: d, label: d === today ? 'Hoy' : formatFechaCorta(d) }));
  tabs.push({ key: 'general', label: 'General' });

  dayTabsEl.innerHTML = tabs
    .map((t) => `<button class="day-tab ${selectedDay === t.key ? 'active' : ''}" data-day="${escapeHtml(t.key)}">${escapeHtml(t.label)}</button>`)
    .join('');

  dayTabsEl.querySelectorAll('.day-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      selectedDay = btn.dataset.day;
      renderDetail();
    });
  });
}

function renderDetail() {
  detailRouteTitle.textContent = `Ruta ${selectedRoute}`;
  const today = todayIso();
  const dayLabel = selectedDay === 'general' ? 'Listado general' : (selectedDay === today ? 'Hoy' : formatFechaCorta(selectedDay));
  subtitleText.textContent = `Ruta ${selectedRoute} · ${dayLabel}`;
  renderDayTabs();

  const term = searchTerm.trim().toLowerCase();
  const searching = term.length > 0;
  // A search looks across every route (NIU único), ignoring la ruta/día seleccionados.
  const baseRows = searching ? generalRowsAll() : currentRouteRowsForDay();

  const withStatus = baseRows.map((r) => ({ ...r, _eval: evaluarEstado(r.niu) }));

  let filtered = withStatus.filter((r) => {
    if (searching) {
      return `${r.niu} ${r.nombre}`.toLowerCase().includes(term);
    }
    const s = r._eval.status;
    if (currentFilter === 'pendiente' && s !== 'no' && s !== 'unk') return false;
    if (currentFilter === 'pagado' && s !== 'ok') return false;
    if (currentFilter === 'notificado' && s !== 'notificado') return false;
    if (currentFilter === 'suspendido' && s !== 'suspendido') return false;
    return true;
  });

  const okCount = withStatus.filter((r) => r._eval.status === 'ok').length;
  const noCount = withStatus.filter((r) => r._eval.status === 'no').length;
  const unkCount = withStatus.filter((r) => r._eval.status === 'unk').length;
  const notificadoCount = withStatus.filter((r) => r._eval.status === 'notificado').length;
  const suspendidoCount = withStatus.filter((r) => r._eval.status === 'suspendido').length;

  const chipCounts = {
    pendiente: noCount + unkCount,
    pagado: okCount,
    notificado: notificadoCount,
    suspendido: suspendidoCount,
    todos: withStatus.length,
  };
  document.querySelectorAll('#view-detail .chip').forEach((chip) => {
    const key = chip.dataset.filter;
    const label = chip.dataset.label || chip.textContent.replace(/\s*\(\d+\)$/, '');
    chip.dataset.label = label; // remember the plain label the first time
    chip.textContent = `${label} (${chipCounts[key] ?? 0})`;
  });

  if (filtered.length === 0) {
    listEl.innerHTML = `<div class="empty">Sin resultados</div>`;
    return;
  }

  listEl.innerHTML = filtered
    .map((r) => {
      const s = r._eval.status;
      const badgeClass = s === 'ok' ? (r._eval.sinConfirmar ? 'ok-unconfirmed' : 'ok') : s === 'no' ? 'no' : s === 'suspendido' ? 'suspendido' : s === 'notificado' ? 'notificado' : 'unk';
      const badgeLabel = s === 'ok' ? (r._eval.sinConfirmar ? '⏳ Al día (sin confirmar)' : '✅ Al día') : s === 'no' ? '❌ Pendiente' : s === 'suspendido' ? '✂️ Suspendido' : s === 'notificado' ? '📢 Notificado' : '⏳ Sin revisar';
      const scan = scanResults[r.niu];
      const manual = pagosManuales[r.niu];
      const corte = cortadoRecord(r.niu);
      let infoLine;
      if (r._eval.sinConfirmar) {
        infoLine = `Marcado como pagado por ${escapeHtml(manual.liniero || 'liniero')} — pendiente de confirmar en el portal`;
      } else if (s === 'suspendido') {
        infoLine = `Suspendido por ${escapeHtml(corte.liniero)} el ${new Date(corte.fecha).toLocaleString('es-CO')}`;
        if (r._eval.pagado) infoLine += ' · 💰 Ya pagó — falta reconexión';
      } else if (s === 'notificado') {
        infoLine = `Notificado por ${escapeHtml(corte.liniero)} el ${new Date(corte.fecha).toLocaleString('es-CO')}`;
      } else if (r._eval.corteInvalido) {
        infoLine = `El pago fue anterior al corte — no aplica suspensión`;
      } else if (scan) {
        const fechaConHora = scan.fechaRegistroPago || scan.ultimoPago || '—';
        infoLine = `Último pago: ${escapeHtml(fechaConHora)}${scan.estadoPago ? ' (' + escapeHtml(scan.estadoPago) + ')' : ''}`;
      } else {
        infoLine = 'Aún no revisado';
      }
      const rutaTag = searching ? ` · Ruta ${escapeHtml(r.ruta || '—')}` : '';
      const diaTag = !searching && selectedDay === 'general' && r.fecha ? ` · ${r.fecha === today ? 'Hoy' : formatFechaCorta(r.fecha)}` : '';
      const showCorteBtn = s === 'no' || s === 'unk' || s === 'notificado';
      const showPagadoBtn = (s === 'no' || s === 'unk' || s === 'notificado');
      const showQuitarMarcaBtn = r._eval.sinConfirmar;
      const showSubidoBtn = s === 'suspendido' && corte && !corte.subidoCortes;
      const subidoBadge = s === 'suspendido' && corte && corte.subidoCortes ? '<span class="uploaded-tag">📤 Reportado</span>' : '';

      const yaCompartido = s === 'ok' && !r._eval.sinConfirmar && scan && scan.numeroFactura ? compartidoRecord(r.niu, scan.numeroFactura) : null;
      const showCompartirBtn = s === 'ok' && !r._eval.sinConfirmar && scan && scan.numeroFactura && !yaCompartido;
      const compartidoBadge = yaCompartido ? '<span class="uploaded-tag">✅ Compartido</span>' : '';

      const atraso = parseInt(r.mesesAtrasados, 10) || 0;
      const nameClass = atraso === 0 ? 'name-atraso-0' : atraso === 1 ? 'name-atraso-1' : 'name-atraso-2';
      const corteBtnClass = showCorteBtn && atraso >= 2 ? 'corte-btn corte-btn-urgent' : 'corte-btn';

      return `
      <div class="item" data-niu="${escapeHtml(r.niu)}">
        <div class="info">
          <div class="name ${nameClass}">${escapeHtml(r.nombre || '(sin nombre)')}</div>
          <div class="meta">NIU ${escapeHtml(r.niu)} · ${escapeHtml(r.direccion || '')}${rutaTag}${diaTag}</div>
          <div class="meta2">Medidor ${escapeHtml(r.medidor || '—')} · Sector ${escapeHtml(r.sector || '—')} · Atraso: ${escapeHtml(r.mesesAtrasados || '0')} · Saldo: ${formatMoney(r.saldoPendiente)}</div>
          <div class="meta3">${infoLine} ${subidoBadge}${compartidoBadge}</div>
          <div class="meta4 phone-row">${renderPhoneField(r.niu, 'telefono', r.telefono)}${renderPhoneField(r.niu, 'celular', r.celular)}</div>
        </div>
        <div class="item-actions">
          <div class="badge ${badgeClass}">${badgeLabel}</div>
          <button class="recheck-btn" data-niu="${escapeHtml(r.niu)}">🔄 Revisar</button>
          ${showCorteBtn ? `<button class="${corteBtnClass}" data-niu="${escapeHtml(r.niu)}">✂️ Suspender</button>` : ''}
          ${showPagadoBtn ? `<button class="pagado-btn" data-niu="${escapeHtml(r.niu)}">💰 Pagó (sin confirmar)</button>` : ''}
          ${showQuitarMarcaBtn ? `<button class="quitar-marca-btn" data-niu="${escapeHtml(r.niu)}">✕ Quitar marca</button>` : ''}
          ${showSubidoBtn ? `<button class="subido-btn" data-id="${escapeHtml(corte.id)}">📤 Marcar subido a cortes</button>` : ''}
          ${showCompartirBtn ? `<button class="compartir-btn" data-niu="${escapeHtml(r.niu)}">📤 Compartir comprobante</button>` : ''}
        </div>
      </div>`;
    })
    .join('');

  listEl.querySelectorAll('.recheck-btn').forEach((btn) => {
    btn.addEventListener('click', () => recheckSingle(btn.dataset.niu));
  });
  listEl.querySelectorAll('.corte-btn').forEach((btn) => {
    btn.addEventListener('click', () => openCorteModal(btn.dataset.niu));
  });
  listEl.querySelectorAll('.pagado-btn').forEach((btn) => {
    btn.addEventListener('click', () => marcarPagadoSinConfirmar(btn.dataset.niu));
  });
  listEl.querySelectorAll('.quitar-marca-btn').forEach((btn) => {
    btn.addEventListener('click', () => quitarMarcaPagado(btn.dataset.niu));
  });
  listEl.querySelectorAll('.subido-btn').forEach((btn) => {
    btn.addEventListener('click', () => marcarSubido(btn.dataset.id));
  });
  listEl.querySelectorAll('.compartir-btn').forEach((btn) => {
    btn.addEventListener('click', () => compartirComprobante(btn.dataset.niu, btn));
  });
  // "+ Agregar" abre el campo de edición (input + botón ✓)
  listEl.querySelectorAll('.phone-add-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      editingPhone = { niu: btn.dataset.niu, field: btn.dataset.field };
      renderDetail();
      focusPhoneInput();
    });
  });

  // ✏️ Editar abre el mismo campo, pre-llenado con el valor actual
  listEl.querySelectorAll('.phone-edit-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      editingPhone = { niu: btn.dataset.niu, field: btn.dataset.field };
      renderDetail();
      focusPhoneInput();
    });
  });

  // ✕ Cancelar cierra el campo sin guardar
  listEl.querySelectorAll('.phone-cancel-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      editingPhone = null;
      renderDetail();
    });
  });

  // ✓ Guardar — solo aquí se envía al servidor, nunca automáticamente
  listEl.querySelectorAll('.phone-save-btn').forEach((btn) => {
    btn.addEventListener('click', () => confirmPhoneEdit(btn.dataset.niu, btn.dataset.field));
  });

  // Enter dentro del campo también guarda (equivalente a tocar ✓)
  listEl.querySelectorAll('.phone-input').forEach((input) => {
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') confirmPhoneEdit(input.dataset.niu, input.dataset.field);
      if (e.key === 'Escape') { editingPhone = null; renderDetail(); }
    });
  });

  // 🗑️ Eliminar — pide confirmación (ej. cliente dijo que el número está mal)
  listEl.querySelectorAll('.phone-delete-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      const ok = window.confirm('¿Eliminar este número? Podrás agregar uno nuevo después.');
      if (ok) savePhoneNumber(btn.dataset.niu, btn.dataset.field, '');
    });
  });
}

function focusPhoneInput() {
  // El input recién se creó al re-renderizar — hay que esperar al siguiente frame
  setTimeout(() => {
    const input = listEl.querySelector('.phone-input');
    if (input) { input.focus(); input.select(); }
  }, 0);
}

function confirmPhoneEdit(niu, field) {
  const input = listEl.querySelector(`.phone-input[data-niu="${CSS.escape(niu)}"][data-field="${field}"]`);
  if (!input) return;
  const value = input.value.trim();
  editingPhone = null;
  if (/^0+$/.test(value)) {
    alert('Ese número no es válido (solo ceros). Inténtalo de nuevo.');
    return;
  }
  savePhoneNumber(niu, field, value); // value puede ser '' si borró todo el texto y confirmó — equivale a eliminar
}

function parseMoneyValue(value) {
  if (value === null || value === undefined || value === '') return 0;
  let s = String(value).trim().replace(/[^\d.,-]/g, '');
  if (!s) return 0;
  if (s.includes(',') && s.includes('.')) {
    // "100.000,50" -> punto = miles, coma = decimales
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    // "100000,50" -> coma = decimales
    s = s.replace(',', '.');
  } else {
    // "100.000" -> punto = miles (grupos de 3 dígitos)
    s = s.replace(/\.(?=\d{3}(\D|$))/g, '');
  }
  const num = parseFloat(s);
  return isNaN(num) ? 0 : num;
}

function formatMoney(value) {
  if (value === null || value === undefined || value === '') return '—';
  return formatMoneyFromNumber(parseMoneyValue(value));
}

function formatMoneyFromNumber(num) {
  return '$ ' + Math.round(num || 0).toLocaleString('es-CO');
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function renderPhoneField(niu, field, value) {
  const label = field === 'telefono' ? '📞' : '📱';
  const isEditing = editingPhone && editingPhone.niu === niu && editingPhone.field === field;

  if (isEditing) {
    return `<span class="phone-editing">
      ${label}
      <input type="tel" class="phone-input" data-niu="${escapeHtml(niu)}" data-field="${field}" value="${escapeHtml(value || '')}" placeholder="Número..." />
      <button class="phone-save-btn" data-niu="${escapeHtml(niu)}" data-field="${field}" title="Guardar">✓</button>
      <button class="phone-cancel-btn" title="Cancelar">✕</button>
    </span>`;
  }

  if (value) {
    return `<span class="phone-filled">
      <a class="phone-link" href="tel:${escapeHtml(value)}">${label} ${escapeHtml(value)}</a>
      <button class="phone-edit-btn" data-niu="${escapeHtml(niu)}" data-field="${field}" title="Editar">✏️</button>
      <button class="phone-delete-btn" data-niu="${escapeHtml(niu)}" data-field="${field}" title="Eliminar">🗑️</button>
    </span>`;
  }

  return `<span class="phone-empty">
    ${label}
    <button class="phone-add-btn" data-niu="${escapeHtml(niu)}" data-field="${field}">+ Agregar</button>
  </span>`;
}

// ---------- Marca manual "pagado sin confirmar" ----------
async function marcarPagadoSinConfirmar(niu) {
  const row = generalRowForNiu(niu);
  const liniero = (row && rutaLinieros[row.ruta]) || getSavedLiniero() || null;
  try {
    pagosManuales[niu] = { liniero, fecha: new Date().toISOString() };
    lsSet(LS_PAGOS, pagosManuales);
    await syncPagosManuales();
    renderDetail();
  } catch (e) {
    alert('No se pudo guardar la marca — revisa tu conexión e intenta de nuevo.');
  }
}

async function quitarMarcaPagado(niu) {
  const ok = window.confirm('¿Quitar la marca de "pagado sin confirmar"? El NIU volverá a pendiente.');
  if (!ok) return;
  try {
    delete pagosManuales[niu];
    lsSet(LS_PAGOS, pagosManuales);
    await syncPagosManuales();
    renderDetail();
  } catch (e) {
    alert('No se pudo quitar la marca — revisa tu conexión e intenta de nuevo.');
  }
}

// ---------- Comprobante (voucher) generation + share ----------
function drawVoucherCanvas(row, scan) {
  const canvas = document.createElement('canvas');
  const W = 720;
  const PAD = 36;
  canvas.width = W;
  const ctx = canvas.getContext('2d');

  const lines = [
    ['Nombre', row.nombre || '—'],
    ['NIU', row.niu],
    ['Fecha registro pago', scan.fechaRegistroPago || '—'],
    ['Estado', scan.estadoPago || '—'],
    ['Fecha pago', scan.ultimoPago || '—'],
    ['Valor pagado', scan.valorPagado || '—'],
    ['Medio de pago', scan.medioPago || '—'],
    ['Entidad recaudo', (scan.entidadRecaudo || '—').replace(/^:\s*/, '')],
    ['Número cuenta', scan.numeroCuenta || '—'],
    ['Factura', scan.numeroFactura ? `Pago Factura: ${scan.numeroFactura}` : '—'],
  ];

  const lineHeight = 44;
  const headerHeight = 120;
  const footerHeight = 60;
  canvas.height = headerHeight + lines.length * lineHeight + footerHeight;

  // Background
  ctx.fillStyle = '#0b1220';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  // Header
  ctx.fillStyle = '#16a34a';
  ctx.fillRect(0, 0, canvas.width, headerHeight);
  ctx.fillStyle = '#ffffff';
  ctx.font = 'bold 30px Arial';
  ctx.fillText('Comprobante de Pago', PAD, 55);
  ctx.font = '16px Arial';
  ctx.fillText(String(row.ruta ? `Ruta ${row.ruta}` : 'Consulta EEP'), PAD, 85);

  // Body
  let y = headerHeight + 34;
  ctx.textBaseline = 'alphabetic';
  lines.forEach(([label, value]) => {
    ctx.fillStyle = '#8a94a6';
    ctx.font = '15px Arial';
    ctx.fillText(label, PAD, y);
    ctx.fillStyle = '#e6ebf5';
    ctx.font = 'bold 19px Arial';
    ctx.fillText(String(value), PAD, y + 22);
    y += lineHeight;
  });

  // Footer
  ctx.strokeStyle = '#263049';
  ctx.beginPath();
  ctx.moveTo(PAD, y);
  ctx.lineTo(canvas.width - PAD, y);
  ctx.stroke();
  ctx.fillStyle = '#8a94a6';
  ctx.font = '13px Arial';
  ctx.fillText(`Generado ${new Date().toLocaleString('es-CO')}`, PAD, y + 30);

  return canvas;
}

async function compartirComprobante(niu, btn) {
  const row = generalRowForNiu(niu);
  const scan = scanResults[niu];
  if (!row || !scan) return;

  const liniero = rutaLinieros[row.ruta] || getSavedLiniero() || 'Sin asignar';

  btn.disabled = true;
  btn.textContent = 'Generando...';

  try {
    const canvas = drawVoucherCanvas(row, scan);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    const file = new File([blob], `comprobante_${niu}.png`, { type: 'image/png' });

    let shared = false;
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: 'Comprobante de pago', text: `Comprobante de pago — NIU ${niu}` });
        shared = true;
      } catch (shareErr) {
        // user cancelled the share sheet — not an error, just don't mark as shared
        shared = false;
      }
    } else {
      // Fallback: download the image so it can be shared manually
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `comprobante_${niu}.png`;
      a.click();
      URL.revokeObjectURL(url);
      shared = true; // downloaded successfully — treat as "handled"
    }

    if (shared) {
      compartidos.push({ niu, numeroFactura: scan.numeroFactura || null, liniero: liniero || null, fecha: new Date().toISOString() });
      lsSet(LS_COMPARTIDOS, compartidos);
      await syncCompartidos();
      renderDetail();
    }
  } catch (e) {
    btn.textContent = 'Error — reintentar';
    btn.disabled = false;
    return;
  }
  btn.disabled = false;
  btn.textContent = '📤 Compartir comprobante';
}

// ---------- Corte modal ----------
const corteModalOverlay = document.getElementById('corteModalOverlay');
const corteModalNombre = document.getElementById('corteModalNombre');
const corteModalNiu = document.getElementById('corteModalNiu');
const corteModalDireccion = document.getElementById('corteModalDireccion');
const corteTipoSelect = document.getElementById('corteTipoSelect');
const corteHoraInput = document.getElementById('corteHoraInput');
const corteLinieroInput = document.getElementById('corteLinieroInput');
const corteCancelBtn = document.getElementById('corteCancelBtn');
const corteConfirmBtn = document.getElementById('corteConfirmBtn');

function nowAsHHMM() {
  const d = new Date();
  return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}

function allKnownLinieros() {
  const names = Array.from(new Set(Object.values(rutaLinieros).filter((n) => n && n.trim())));
  names.sort((a, b) => a.localeCompare(b, 'es'));
  return names;
}

function populateLinieroSelect(defaultName) {
  const names = allKnownLinieros();
  if (defaultName && !names.includes(defaultName)) names.unshift(defaultName);
  if (names.length === 0) names.push('Sin asignar');
  corteLinieroInput.innerHTML = names.map((n) => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('');
  corteLinieroInput.value = defaultName && names.includes(defaultName) ? defaultName : names[0];
}

function openCorteModal(niu) {
  const row = generalRowForNiu(niu);
  if (!row) return;
  pendingCorteRow = row;
  corteModalNombre.textContent = row.nombre || '(sin nombre)';
  corteModalNiu.textContent = row.niu;
  corteModalDireccion.textContent = row.direccion || '—';
  corteTipoSelect.value = 'poste';
  corteHoraInput.value = nowAsHHMM();
  // Defaults to the route's assigned liniero — pick another from the list only if needed.
  populateLinieroSelect(rutaLinieros[row.ruta] || getSavedLiniero());
  corteModalOverlay.style.display = 'flex';
}

function closeCorteModal() {
  corteModalOverlay.style.display = 'none';
  pendingCorteRow = null;
}

corteCancelBtn.addEventListener('click', closeCorteModal);

// ---------- Credenciales del portal Oracle ----------
// Se guardan cifradas (AES-GCM, clave derivada con PBKDF2 de una "clave de desbloqueo")
// solo en este celular. Se envían al servidor únicamente al iniciar una consulta,
// por HTTPS, y el servidor no las guarda.
const LS_CRED = 'eep_cred_vault';
let credEnMemoria = null; // { usuario, clave } mientras la app esté abierta y desbloqueada
let credResolver = null;
let credModoEditar = false;

const credModalOverlay = document.getElementById('credModalOverlay');
const credUsuario = document.getElementById('credUsuario');
const credClave = document.getElementById('credClave');
const credPin = document.getElementById('credPin');
const credPinLabel = document.getElementById('credPinLabel');
const credCamposWrap = document.getElementById('credCamposWrap');
const credEstado = document.getElementById('credEstado');
const credMsg = document.getElementById('credMsg');
const credSaveBtn = document.getElementById('credSaveBtn');
const credEditBtn = document.getElementById('credEditBtn');

const b64 = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf)));
const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

async function derivarLlave(pin, salt) {
  const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pin), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt, iterations: 250000, hash: 'SHA-256' },
    base,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
}

async function cifrarCredenciales(cred, pin) {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const llave = await derivarLlave(pin, salt);
  const datos = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, llave, new TextEncoder().encode(JSON.stringify(cred)));
  return { salt: b64(salt), iv: b64(iv), datos: b64(datos) };
}

async function descifrarCredenciales(boveda, pin) {
  const llave = await derivarLlave(pin, unb64(boveda.salt));
  const plano = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(boveda.iv) }, llave, unb64(boveda.datos));
  return JSON.parse(new TextDecoder().decode(plano));
}

function leerBoveda() {
  try { return JSON.parse(localStorage.getItem(LS_CRED)); } catch (e) { return null; }
}

function cerrarCredModal(resultado) {
  credModalOverlay.style.display = 'none';
  credPin.value = '';
  credClave.value = '';
  if (credResolver) { credResolver(resultado || null); credResolver = null; }
}

function mostrarCredModal(editar) {
  const hayBoveda = !!leerBoveda();
  credModoEditar = editar || !hayBoveda;
  credCamposWrap.style.display = credModoEditar ? 'block' : 'none';
  credEditBtn.style.display = !credModoEditar && hayBoveda ? 'block' : 'none';
  credPinLabel.textContent = credModoEditar ? 'Crea una clave de desbloqueo' : 'Clave de desbloqueo';
  credSaveBtn.textContent = credModoEditar ? '💾 Guardar' : '🔓 Desbloquear';
  credEstado.textContent = credModoEditar
    ? 'Se guardan cifradas solo en este celular. La clave de desbloqueo no se puede recuperar: si la olvidas, vuelve a escribir las credenciales.'
    : 'Escribe tu clave de desbloqueo para usar las credenciales guardadas.';
  credMsg.textContent = '';
  credUsuario.value = '';
  credClave.value = '';
  credPin.value = '';
  credModalOverlay.style.display = 'flex';
}

// Devuelve { usuario, clave } o null si el usuario cierra el modal sin desbloquear.
function obtenerCredenciales() {
  if (credEnMemoria) return Promise.resolve(credEnMemoria);
  return new Promise((resolve) => {
    credResolver = resolve;
    mostrarCredModal(false);
  });
}

document.getElementById('credBtn').addEventListener('click', () => mostrarCredModal(false));
document.getElementById('credCancelBtn').addEventListener('click', () => cerrarCredModal(null));
credEditBtn.addEventListener('click', () => mostrarCredModal(true));

credSaveBtn.addEventListener('click', async () => {
  const pin = credPin.value;
  if (!pin) { credMsg.textContent = 'Escribe la clave de desbloqueo.'; return; }
  credSaveBtn.disabled = true;
  try {
    if (credModoEditar) {
      const usuario = credUsuario.value.trim();
      const clave = credClave.value;
      if (!usuario || !clave) { credMsg.textContent = 'Escribe el usuario y la contraseña del portal.'; return; }
      if (pin.length < 4) { credMsg.textContent = 'La clave de desbloqueo debe tener al menos 4 caracteres.'; return; }
      credMsg.textContent = 'Cifrando…';
      localStorage.setItem(LS_CRED, JSON.stringify(await cifrarCredenciales({ usuario, clave }, pin)));
      credEnMemoria = { usuario, clave };
      cerrarCredModal(credEnMemoria);
    } else {
      credMsg.textContent = 'Desbloqueando…';
      credEnMemoria = await descifrarCredenciales(leerBoveda(), pin);
      cerrarCredModal(credEnMemoria);
    }
  } catch (e) {
    credMsg.textContent = 'Clave de desbloqueo incorrecta.';
  } finally {
    credSaveBtn.disabled = false;
  }
});
corteConfirmBtn.addEventListener('click', async () => {
  const liniero = corteLinieroInput.value;
  if (!liniero) return;
  if (!pendingCorteRow) return;
  corteConfirmBtn.disabled = true;
  corteConfirmBtn.textContent = 'Guardando...';
  try {
    const ahora = new Date();
    let fechaCorte = ahora;
    if (corteHoraInput.value) {
      const [h, mi] = corteHoraInput.value.split(':').map(Number);
      fechaCorte = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate(), h, mi, 0);
    }
    cortes.push({
      id: 'corte_' + ahora.getTime() + '_' + Math.random().toString(36).slice(2, 8),
      niu: pendingCorteRow.niu,
      nombre: pendingCorteRow.nombre || null,
      direccion: pendingCorteRow.direccion || null,
      ruta: pendingCorteRow.ruta || null,
      liniero,
      tipo: corteTipoSelect.value,
      fecha: fechaCorte.toISOString(),
      subidoCortes: false,
    });
    lsSet(LS_CORTES, cortes);
    saveLiniero(liniero);
    await syncCortes();
    closeCorteModal();
    renderDetail();
  } catch (e) {
    corteConfirmBtn.textContent = 'Error — reintentar';
  } finally {
    corteConfirmBtn.disabled = false;
    corteConfirmBtn.textContent = '✂️ Confirmar';
  }
});

async function marcarSubido(corteId) {
  const confirmado = window.confirm('¿Confirmas que ya realizaste el reporte de corte del servicio?');
  if (!confirmado) return;
  try {
    const c = cortes.find((x) => x.id === corteId);
    if (c) c.subidoCortes = true;
    lsSet(LS_CORTES, cortes);
    await syncCortes();
    renderDetail();
  } catch (e) {
    alert('No se pudo guardar — revisa tu conexión e intenta de nuevo.');
  }
}

async function savePhoneNumber(niu, field, value) {
  try {
    const limpio = value && !/^0+$/.test(value) ? value : null;
    value = limpio;
    // El teléfono es global por NIU — se actualiza en TODAS las filas (todos los días) de ese NIU en memoria.
    excelRows.forEach((r) => { if (r.niu === niu) r[field] = value || undefined; });
    saveRows();
    renderDetail();
  } catch (e) {
    alert('No se pudo guardar el número — revisa tu conexión e intenta de nuevo.');
  }
}

// ---------- Backend calls: scanning ----------
async function recheckSingle(niu) {
  runQueuedScan([niu], statusBar, [scanPendingBtn]);
}

async function runQueuedScan(niusList, statusEl, buttons, routaLabel) {
  const credenciales = await obtenerCredenciales();
  if (!credenciales) {
    statusEl.textContent = 'Necesitas las credenciales del portal (botón 🔑)';
    return;
  }
  statusEl.textContent = 'Iniciando consulta...';
  buttons.forEach((b) => { b.disabled = true; });
  try {
    const res = await fetch(`${BACKEND_URL}/api/scan/start`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ nius: niusList, ruta: routaLabel || selectedRoute, credenciales }),
    });
    const data = await res.json();
    statusEl.textContent = data.message;
    if (data.ok && data.jobId) pollQueuedScan(data.jobId, statusEl, buttons);
    else buttons.forEach((b) => { b.disabled = false; });
  } catch (e) {
    statusEl.textContent = 'No se pudo conectar al servidor';
    buttons.forEach((b) => { b.disabled = false; });
  }
}

async function pollQueuedScan(jobId, statusEl, buttons) {
  try {
    const res = await fetch(`${BACKEND_URL}/api/scan/status/${jobId}`);
    const data = await res.json();
    if (data.status === 'queued') {
      statusEl.textContent = data.queuePosition > 1
        ? `En cola — alguien más está consultando (posición ${data.queuePosition})`
        : 'En cola, comenzando...';
      setTimeout(() => pollQueuedScan(jobId, statusEl, buttons), 2000);
    } else if (data.status === 'running') {
      statusEl.textContent = `Consultando... ${data.progress}/${data.total}`;
      setTimeout(() => pollQueuedScan(jobId, statusEl, buttons), 2000);
    } else {
      buttons.forEach((b) => { b.disabled = false; });
      if (data.error) {
        statusEl.textContent = 'Error: ' + data.error;
      } else {
        statusEl.textContent = 'Consulta completada';
        await applyScanResults(jobId);
        lastUpdateEl.textContent = `Última actualización: ${new Date().toLocaleString('es-CO')}`;
        renderDetail();
      }
    }
  } catch (e) {
    buttons.forEach((b) => { b.disabled = false; });
    statusEl.textContent = 'No se pudo conectar al servidor';
  }
}

// Se revisa el conjunto de NIU del día/vista actualmente seleccionada (día
// concreto o "General" dentro de la ruta) — todo lo que no esté "al día" confirmado.
scanPendingBtn.addEventListener('click', () => {
  const niusList = currentRouteRowsForDay()
    .filter((r) => needsScan(r.niu))
    .map((r) => r.niu);
  if (niusList.length === 0) {
    statusBar.textContent = 'No hay pendientes para revisar';
    return;
  }
  runQueuedScan(niusList, statusBar, [scanPendingBtn]);
});

const scanAllBtn = document.getElementById('scanAllBtn');
const routesStatusBar = document.getElementById('routesStatusBar');
scanAllBtn.addEventListener('click', async () => {
  await Promise.all([syncSharedResults(), syncPagosManuales(), syncCortes()]);
  // Fuera de las rutas: siempre general (NIU único), sin importar el día.
  const niusList = generalRowsAll()
    .filter((r) => needsScan(r.niu))
    .map((r) => r.niu);
  if (niusList.length === 0) {
    routesStatusBar.textContent = 'No hay pendientes en ninguna ruta 🎉';
    return;
  }
  runQueuedScan(niusList, routesStatusBar, [scanAllBtn], 'TODAS');
});

document.querySelectorAll('.chip').forEach((chip) => {
  chip.addEventListener('click', () => {
    document.querySelectorAll('.chip').forEach((c) => c.classList.remove('active'));
    chip.classList.add('active');
    currentFilter = chip.dataset.filter;
    renderDetail();
  });
});

searchInput.addEventListener('input', (e) => {
  searchTerm = e.target.value;
  renderDetail();
});

// ---------- Stats / indicadores screen ----------
const statsBtn = document.getElementById('statsBtn');
const backFromStatsBtn = document.getElementById('backFromStatsBtn');
const monthSelect = document.getElementById('monthSelect');
const statsSummary = document.getElementById('statsSummary');
const statsList = document.getElementById('statsList');
const statsRutas = document.getElementById('statsRutas');

function monthKey(dateIso) {
  const d = new Date(dateIso);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}
function monthLabel(key) {
  const [y, m] = key.split('-');
  const nombres = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return `${nombres[parseInt(m, 10) - 1]} ${y}`;
}

statsBtn.addEventListener('click', async () => {
  showView('stats');
  await Promise.all([syncCortes(), syncCompartidos(), syncSharedResults(), syncPagosManuales()]);
  renderStats();
  renderStatsRutas();
});
backFromStatsBtn.addEventListener('click', goToRoutes);

function renderStats() {
  const currentKey = monthKey(new Date().toISOString());
  const keysInData = Array.from(new Set([...cortes.map((c) => monthKey(c.fecha)), ...compartidos.map((c) => monthKey(c.fecha))]));
  if (!keysInData.includes(currentKey)) keysInData.push(currentKey);
  keysInData.sort().reverse();

  const prevSelected = monthSelect.value;
  monthSelect.innerHTML = keysInData.map((k) => `<option value="${k}">${monthLabel(k)}</option>`).join('');
  monthSelect.value = keysInData.includes(prevSelected) ? prevSelected : currentKey;

  renderStatsForMonth(monthSelect.value);
}

monthSelect.addEventListener('change', () => renderStatsForMonth(monthSelect.value));

function renderStatsForMonth(key) {
  const cortesDelMes = cortes.filter((c) => monthKey(c.fecha) === key);
  const cortesReales = cortesDelMes.filter((c) => c.tipo !== 'notificacion');
  const notificaciones = cortesDelMes.filter((c) => c.tipo === 'notificacion');
  const compartidosDelMes = compartidos.filter((c) => monthKey(c.fecha) === key);

  const today = new Date().toISOString().slice(0, 10);
  const cortesHoy = cortesReales.filter((c) => c.fecha.slice(0, 10) === today).length;

  // start of the current week (Monday)
  const now = new Date();
  const dayOfWeek = (now.getDay() + 6) % 7; // 0 = Monday
  const monday = new Date(now);
  monday.setDate(now.getDate() - dayOfWeek);
  monday.setHours(0, 0, 0, 0);
  const cortesSemana = cortesReales.filter((c) => new Date(c.fecha) >= monday).length;

  statsSummary.innerHTML = `
    <div class="box"><div class="n">${cortesReales.length}</div><div class="l">Cortes del mes</div></div>
    <div class="box"><div class="n">${notificaciones.length}</div><div class="l">Notificaciones</div></div>
    <div class="box"><div class="n">${compartidosDelMes.length}</div><div class="l">Comprobantes compartidos</div></div>
    <div class="box"><div class="n">${cortesSemana}</div><div class="l">Cortes esta semana</div></div>
    <div class="box"><div class="n">${cortesHoy}</div><div class="l">Cortes hoy</div></div>
  `;

  // Per-liniero totals: cortes reales + notificaciones + comprobantes compartidos
  const linierosSet = new Set([
    ...cortesDelMes.map((c) => c.liniero),
    ...compartidosDelMes.map((c) => c.liniero).filter(Boolean),
  ]);

  if (linierosSet.size === 0) {
    statsList.innerHTML = `<div class="empty">Sin actividad registrada este mes</div>`;
    return;
  }

  const porLiniero = Array.from(linierosSet).map((liniero) => {
    const misCortes = cortesReales.filter((c) => c.liniero === liniero);
    const misNotificaciones = notificaciones.filter((c) => c.liniero === liniero);
    const misCompartidos = compartidosDelMes.filter((c) => c.liniero === liniero);
    const totalAcciones = misCortes.length + misNotificaciones.length + misCompartidos.length;
    const porDia = {};
    [...misCortes, ...misNotificaciones].forEach((r) => {
      const dia = r.fecha.slice(0, 10);
      porDia[dia] = (porDia[dia] || 0) + 1;
    });
    return { liniero, misCortes, misNotificaciones, misCompartidos, totalAcciones, porDia };
  }).sort((a, b) => b.totalAcciones - a.totalAcciones);

  statsList.innerHTML = porLiniero
    .map(({ liniero, misCortes, misNotificaciones, misCompartidos, porDia }) => {
      const diasOrdenados = Object.keys(porDia).sort().reverse();
      const detalleDias = diasOrdenados
        .map((d) => `<div class="stat-day-row"><span>${new Date(d).toLocaleDateString('es-CO')}</span><span>${porDia[d]} cortes/notificaciones</span></div>`)
        .join('');
      return `
      <div class="item">
        <div class="info">
          <div class="name">${escapeHtml(liniero)}</div>
          <div class="meta">✂️ ${misCortes.length} cortes · 📢 ${misNotificaciones.length} notificaciones · 📤 ${misCompartidos.length} comprobantes compartidos</div>
          <div class="stat-days">${detalleDias}</div>
        </div>
      </div>`;
    })
    .join('');
}

function renderStatsRutas() {
  const counts = {};
  generalRowsAll().forEach((r) => {
    const ruta = r.ruta || '(sin ruta)';
    if (!counts[ruta]) counts[ruta] = { total: 0, ok: 0, no: 0, unk: 0, notificado: 0, suspendido: 0, recaudado: 0, porRecaudar: 0 };
    counts[ruta].total++;
    const s = evaluarEstado(r.niu).status;
    if (s === 'ok') counts[ruta].ok++;
    else if (s === 'no') counts[ruta].no++;
    else if (s === 'unk') counts[ruta].unk++;
    else if (s === 'notificado') counts[ruta].notificado++;
    else if (s === 'suspendido') counts[ruta].suspendido++;

    const scan = scanResults[r.niu];
    if (s === 'ok') {
      if (scan && scan.valorPagado) counts[ruta].recaudado += parseMoneyValue(scan.valorPagado);
    } else {
      counts[ruta].porRecaudar += parseMoneyValue(r.saldoPendiente);
    }
  });

  const rutas = Object.keys(counts).sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));
  if (rutas.length === 0) {
    statsRutas.innerHTML = `<div class="empty">Sube un listado para ver esto</div>`;
    return;
  }

  const pct = (n, total) => (total === 0 ? 0 : Math.round((n / total) * 100));

  statsRutas.innerHTML = rutas
    .map((ruta) => {
      const c = counts[ruta];
      const pendientesPct = pct(c.no + c.unk, c.total);
      const pagadosPct = pct(c.ok, c.total);
      const notificadosPct = pct(c.notificado, c.total);
      const suspendidosPct = pct(c.suspendido, c.total);
      return `
      <div class="item">
        <div class="info" style="width:100%">
          <div class="name">Ruta ${escapeHtml(ruta)} · ${c.total} matrículas</div>
          <div class="route-pct-row"><span>🟢 Al día</span><span>${c.ok} (${pagadosPct}%)</span></div>
          <div class="route-pct-row"><span>🔴 Pendientes</span><span>${c.no + c.unk} (${pendientesPct}%)</span></div>
          <div class="route-pct-row"><span>📢 Notificados</span><span>${c.notificado} (${notificadosPct}%)</span></div>
          <div class="route-pct-row"><span>✂️ Suspendidos</span><span>${c.suspendido} (${suspendidosPct}%)</span></div>
          <div class="route-pct-row route-money-row recaudado"><span>💰 Recaudado</span><span>${formatMoneyFromNumber(c.recaudado)}</span></div>
          <div class="route-pct-row route-money-row porrecaudar"><span>⏳ Por recaudar</span><span>${formatMoneyFromNumber(c.porRecaudar)}</span></div>
        </div>
      </div>`;
    })
    .join('');
}

// ---------- Init ----------
checkMonthRollover();
loadAll();
Promise.all([syncPagosManuales(), syncCortes(), syncCompartidos(), syncRutaLinieros()]).then(() => {
  if (excelRows.length === 0) showView('upload');
  else if (selectedRoute) goToDetail();
  else goToRoutes();
});

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('service-worker.js').catch(() => {});
}