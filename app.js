// Backend en producción (Render)
const BACKEND_URL = 'https://consulta-eep-backend.onrender.com';

const LS_ROWS = 'eep_excel_rows';
const LS_ROUTE = 'eep_selected_route';
const LS_RESULTS = 'eep_scan_results';
const LS_LINIERO = 'eep_liniero_nombre';

let excelRows = [];
let scanResults = {}; // { [niu]: { pagado, ultimoPago, estadoPago, checkedAt } }
let cortes = []; // [{ niu, nombre, direccion, ruta, liniero, fecha }]
let rutaLinieros = {}; // { [ruta]: linieroName }
let selectedRoute = null;
let currentFilter = 'pendiente';
let searchTerm = '';
let pendingCorteRow = null; // row waiting for corte confirmation in the modal

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

function normalizeHeader(h) {
  return String(h)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/_/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeRows(rawRows) {
  return rawRows
    .map((raw) => {
      const row = {};
      Object.keys(raw).forEach((key) => {
        const mapped = HEADER_MAP[normalizeHeader(key)];
        if (mapped) row[mapped] = String(raw[key]).trim();
      });
      return row;
    })
    .filter((r) => r.niu); // discard rows without a NIU — unusable
}

function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        resolve(json);
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

    uploadStatus.textContent = 'Combinando con el listado del servidor...';
    try {
      const res = await fetch(`${BACKEND_URL}/api/data/rows`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows: normalized }),
      });
      const data = await res.json();
      uploadStatus.textContent = `Listo: ${data.nuevos} nuevas, ${data.actualizados} actualizadas, ${data.total} en total ✅`;
      await syncSharedRows();
    } catch (shareErr) {
      // offline fallback: merge locally only
      const byNiu = {};
      excelRows.forEach((r) => { byNiu[r.niu] = r; });
      normalized.forEach((r) => { byNiu[r.niu] = r; });
      excelRows = Object.values(byNiu);
      saveRows();
      uploadStatus.textContent = `${normalized.length} filas leídas (sin conexión — no se compartió con el servidor)`;
    }
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

// ---------- Shared data sync ----------
async function syncSharedRows() {
  try {
    const res = await fetch(`${BACKEND_URL}/api/data/rows`);
    const data = await res.json();
    if (data.rows && data.rows.length > 0) {
      excelRows = data.rows;
      saveRows();
      return true;
    }
  } catch (e) {
    // offline or backend down — keep whatever is cached locally
  }
  return false;
}

async function syncSharedResults() {
  try {
    const res = await fetch(`${BACKEND_URL}/api/results`);
    const data = await res.json();
    if (data.results) {
      scanResults = { ...scanResults, ...data.results };
      saveResults();
    }
  } catch (e) {
    // offline — keep local results
  }
}

async function syncCortes() {
  try {
    const res = await fetch(`${BACKEND_URL}/api/cortes`);
    const data = await res.json();
    cortes = data.cortes || [];
  } catch (e) {
    // offline — keep whatever we had
  }
}

async function syncRutaLinieros() {
  try {
    const res = await fetch(`${BACKEND_URL}/api/rutas-linieros`);
    const data = await res.json();
    rutaLinieros = data.rutaLinieros || {};
  } catch (e) {
    // offline — keep whatever we had
  }
}

// ---------- Routes screen ----------
async function goToRoutes() {
  showView('routes');
  await syncRutaLinieros();
  renderRoutesList();
  subtitleText.textContent = `${excelRows.length} matrículas cargadas`;
}

function renderRoutesList() {
  const routesList = document.getElementById('routesList');
  const counts = {};
  excelRows.forEach((r) => {
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
        await fetch(`${BACKEND_URL}/api/rutas-linieros`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ ruta, liniero }),
        });
        rutaLinieros[ruta] = liniero;
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
const listEl = document.getElementById('list');
const summaryEl = document.getElementById('summary');
const statusBar = document.getElementById('statusBar');
const lastUpdateEl = document.getElementById('lastUpdate');
const scanPendingBtn = document.getElementById('scanPendingBtn');
const searchInput = document.getElementById('search');

let detailRefreshTimer = null;

document.getElementById('backToRoutesBtn').addEventListener('click', () => {
  if (detailRefreshTimer) { clearInterval(detailRefreshTimer); detailRefreshTimer = null; }
  goToRoutes();
});

function currentRouteRows() {
  return excelRows.filter((r) => (r.ruta || '(sin ruta)') === selectedRoute);
}

function cortadoRecord(niu) {
  // most recent corte for this NIU, if any
  const matches = cortes.filter((c) => c.niu === niu);
  if (matches.length === 0) return null;
  return matches.sort((a, b) => new Date(b.fecha) - new Date(a.fecha))[0];
}

// Returns { status: 'ok'|'no'|'unk'|'suspendido', pagado: bool, corteInvalido: bool }
function evaluarEstado(niu) {
  const scan = scanResults[niu];
  const pagado = !!(scan && scan.pagado);
  const corte = cortadoRecord(niu);

  if (!corte || corte.tipo === 'notificacion') {
    // no physical cut — normal payment-based status
    return { status: pagado ? 'ok' : (scan ? 'no' : 'unk'), pagado, corteInvalido: false };
  }

  // Physical cut (poste/pin) — check whether the payment happened BEFORE the cut,
  // which would mean the cut was a mistake (customer had already paid) and doesn't apply.
  const corteTimestamp = new Date(corte.fecha);
  const pagoTimestamp = scan && scan.fechaRegistroPago ? parseSpanishDateTime(scan.fechaRegistroPago) : null;

  if (pagado && pagoTimestamp && pagoTimestamp < corteTimestamp) {
    return { status: 'ok', pagado: true, corteInvalido: true };
  }

  // Valid suspension — stays classified as "suspendido" even if later paid
  // (needs manual reconnection), so we surface `pagado` separately.
  return { status: 'suspendido', pagado, corteInvalido: false };
}

function statusOf(niu) {
  return evaluarEstado(niu).status;
}

function goToDetail() {
  detailRouteTitle.textContent = `Ruta ${selectedRoute}`;
  subtitleText.textContent = `Ruta ${selectedRoute}`;
  showView('detail');
  Promise.all([syncSharedResults(), syncCortes()]).then(renderDetail);
  renderDetail();

  if (detailRefreshTimer) clearInterval(detailRefreshTimer);
  detailRefreshTimer = setInterval(() => {
    Promise.all([syncSharedResults(), syncCortes()]).then(renderDetail);
  }, 20000); // pick up progress from other devices every 20s
}

function renderDetail() {
  const rows = currentRouteRows();
  const term = searchTerm.trim().toLowerCase();

  const withStatus = rows.map((r) => ({ ...r, _eval: evaluarEstado(r.niu) }));

  let filtered = withStatus.filter((r) => {
    const s = r._eval.status;
    if (currentFilter === 'pendiente' && s !== 'no' && s !== 'unk') return false;
    if (currentFilter === 'pagado' && s !== 'ok') return false;
    if (currentFilter === 'suspendido' && s !== 'suspendido') return false;
    if (term && !`${r.niu} ${r.nombre}`.toLowerCase().includes(term)) return false;
    return true;
  });

  const okCount = withStatus.filter((r) => r._eval.status === 'ok').length;
  const noCount = withStatus.filter((r) => r._eval.status === 'no').length;
  const unkCount = withStatus.filter((r) => r._eval.status === 'unk').length;
  const suspendidoCount = withStatus.filter((r) => r._eval.status === 'suspendido').length;

  summaryEl.innerHTML = `
    <div class="box"><div class="n">${withStatus.length}</div><div class="l">Total</div></div>
    <div class="box"><div class="n" style="color:#4ade80">${okCount}</div><div class="l">Al día</div></div>
    <div class="box"><div class="n" style="color:#f87171">${noCount + unkCount}</div><div class="l">Pendientes</div></div>
    <div class="box"><div class="n" style="color:#93c5fd">${suspendidoCount}</div><div class="l">Suspendidos</div></div>
  `;

  if (filtered.length === 0) {
    listEl.innerHTML = `<div class="empty">Sin resultados</div>`;
    return;
  }

  listEl.innerHTML = filtered
    .map((r) => {
      const s = r._eval.status;
      const badgeClass = s === 'ok' ? 'ok' : s === 'no' ? 'no' : s === 'suspendido' ? 'suspendido' : 'unk';
      const badgeLabel = s === 'ok' ? '✅ Al día' : s === 'no' ? '❌ Pendiente' : s === 'suspendido' ? '✂️ Suspendido' : '⏳ Sin revisar';
      const scan = scanResults[r.niu];
      const corte = cortadoRecord(r.niu);
      let infoLine;
      if (s === 'suspendido') {
        infoLine = `Suspendido por ${escapeHtml(corte.liniero)} el ${new Date(corte.fecha).toLocaleString('es-CO')}`;
        if (r._eval.pagado) infoLine += ' · 💰 Ya pagó — falta reconexión';
      } else if (r._eval.corteInvalido) {
        infoLine = `El pago fue anterior al corte — no aplica suspensión`;
      } else if (scan) {
        infoLine = `Último pago: ${escapeHtml(scan.ultimoPago || '—')}${scan.estadoPago ? ' (' + escapeHtml(scan.estadoPago) + ')' : ''}`;
      } else {
        infoLine = 'Aún no revisado';
      }
      const showCorteBtn = s === 'no' || s === 'unk';
      const showSubidoBtn = s === 'suspendido' && corte && !corte.subidoCortes;
      const subidoBadge = s === 'suspendido' && corte && corte.subidoCortes ? '<span class="uploaded-tag">📤 Reportado</span>' : '';
      return `
      <div class="item" data-niu="${escapeHtml(r.niu)}">
        <div class="info">
          <div class="name">${escapeHtml(r.nombre || '(sin nombre)')}</div>
          <div class="meta">NIU ${escapeHtml(r.niu)} · ${escapeHtml(r.direccion || '')}</div>
          <div class="meta2">Medidor ${escapeHtml(r.medidor || '—')} · Sector ${escapeHtml(r.sector || '—')} · Atraso: ${escapeHtml(r.mesesAtrasados || '0')} · Saldo: ${escapeHtml(r.saldoPendiente || '—')}</div>
          <div class="meta3">${infoLine} ${subidoBadge}</div>
        </div>
        <div class="item-actions">
          <div class="badge ${badgeClass}">${badgeLabel}</div>
          <button class="recheck-btn" data-niu="${escapeHtml(r.niu)}">🔄 Revisar</button>
          ${showCorteBtn ? `<button class="corte-btn" data-niu="${escapeHtml(r.niu)}">✂️ Suspender</button>` : ''}
          ${showSubidoBtn ? `<button class="subido-btn" data-id="${escapeHtml(corte.id)}">📤 Marcar subido a cortes</button>` : ''}
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
  listEl.querySelectorAll('.subido-btn').forEach((btn) => {
    btn.addEventListener('click', () => marcarSubido(btn.dataset.id));
  });
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
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

function openCorteModal(niu) {
  const row = excelRows.find((r) => r.niu === niu);
  if (!row) return;
  pendingCorteRow = row;
  corteModalNombre.textContent = row.nombre || '(sin nombre)';
  corteModalNiu.textContent = row.niu;
  corteModalDireccion.textContent = row.direccion || '—';
  corteTipoSelect.value = 'poste';
  corteHoraInput.value = nowAsHHMM();
  corteLinieroInput.value = getSavedLiniero();
  corteModalOverlay.style.display = 'flex';
}

function closeCorteModal() {
  corteModalOverlay.style.display = 'none';
  pendingCorteRow = null;
}

corteCancelBtn.addEventListener('click', closeCorteModal);

corteConfirmBtn.addEventListener('click', async () => {
  const liniero = corteLinieroInput.value.trim();
  if (!liniero) {
    corteLinieroInput.style.borderColor = '#dc2626';
    return;
  }
  if (!pendingCorteRow) return;
  corteConfirmBtn.disabled = true;
  corteConfirmBtn.textContent = 'Guardando...';
  try {
    await fetch(`${BACKEND_URL}/api/cortes`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        niu: pendingCorteRow.niu,
        nombre: pendingCorteRow.nombre,
        direccion: pendingCorteRow.direccion,
        ruta: pendingCorteRow.ruta,
        liniero,
        tipo: corteTipoSelect.value,
        hora: corteHoraInput.value,
      }),
    });
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
    await fetch(`${BACKEND_URL}/api/cortes/${corteId}/subido`, { method: 'POST' });
    await syncCortes();
    renderDetail();
  } catch (e) {
    alert('No se pudo guardar — revisa tu conexión e intenta de nuevo.');
  }
}

// ---------- Backend calls: scanning ----------
async function recheckSingle(niu) {
  runQueuedScan([niu], statusBar, [scanPendingBtn]);
}

async function runQueuedScan(niusList, statusEl, buttons, routaLabel) {
  statusEl.textContent = 'Iniciando consulta...';
  buttons.forEach((b) => { b.disabled = true; });
  try {
    const res = await fetch(`${BACKEND_URL}/api/scan/start`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ nius: niusList, ruta: routaLabel || selectedRoute }),
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
        await syncSharedResults();
        lastUpdateEl.textContent = `Última actualización: ${new Date().toLocaleString('es-CO')}`;
        renderDetail();
      }
    }
  } catch (e) {
    buttons.forEach((b) => { b.disabled = false; });
    statusEl.textContent = 'No se pudo conectar al servidor';
  }
}

// Only pendientes + sin revisar get (re)scanned — "al día" and "cortado" are skipped
scanPendingBtn.addEventListener('click', () => {
  const niusList = currentRouteRows()
    .filter((r) => { const s = statusOf(r.niu); return s === 'no' || s === 'unk'; })
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
  await Promise.all([syncSharedResults(), syncCortes()]);
  const niusList = excelRows
    .filter((r) => { const s = statusOf(r.niu); return s === 'no' || s === 'unk'; })
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
  await syncCortes();
  renderStats();
});
backFromStatsBtn.addEventListener('click', goToRoutes);

function renderStats() {
  const currentKey = monthKey(new Date().toISOString());
  const keysInData = Array.from(new Set(cortes.map((c) => monthKey(c.fecha))));
  if (!keysInData.includes(currentKey)) keysInData.push(currentKey);
  keysInData.sort().reverse();

  const prevSelected = monthSelect.value;
  monthSelect.innerHTML = keysInData.map((k) => `<option value="${k}">${monthLabel(k)}</option>`).join('');
  monthSelect.value = keysInData.includes(prevSelected) ? prevSelected : currentKey;

  renderStatsForMonth(monthSelect.value);
}

monthSelect.addEventListener('change', () => renderStatsForMonth(monthSelect.value));

function renderStatsForMonth(key) {
  const cortesDelMes = cortes.filter((c) => monthKey(c.fecha) === key && c.tipo !== 'notificacion');

  const today = new Date().toISOString().slice(0, 10);
  const cortesHoy = cortesDelMes.filter((c) => c.fecha.slice(0, 10) === today).length;

  // start of the current week (Monday)
  const now = new Date();
  const dayOfWeek = (now.getDay() + 6) % 7; // 0 = Monday
  const monday = new Date(now);
  monday.setDate(now.getDate() - dayOfWeek);
  monday.setHours(0, 0, 0, 0);
  const cortesSemana = cortesDelMes.filter((c) => new Date(c.fecha) >= monday).length;

  statsSummary.innerHTML = `
    <div class="box"><div class="n">${cortesDelMes.length}</div><div class="l">Total del mes</div></div>
    <div class="box"><div class="n">${cortesSemana}</div><div class="l">Esta semana</div></div>
    <div class="box"><div class="n">${cortesHoy}</div><div class="l">Hoy</div></div>
  `;

  // Per-liniero totals
  const porLiniero = {};
  cortesDelMes.forEach((c) => {
    porLiniero[c.liniero] = (porLiniero[c.liniero] || 0) + 1;
  });
  const linieros = Object.keys(porLiniero).sort((a, b) => porLiniero[b] - porLiniero[a]);

  if (linieros.length === 0) {
    statsList.innerHTML = `<div class="empty">Sin cortes registrados este mes</div>`;
    return;
  }

  statsList.innerHTML = linieros
    .map((liniero) => {
      const registros = cortesDelMes
        .filter((c) => c.liniero === liniero)
        .sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
      const porDia = {};
      registros.forEach((r) => {
        const dia = r.fecha.slice(0, 10);
        porDia[dia] = (porDia[dia] || 0) + 1;
      });
      const diasOrdenados = Object.keys(porDia).sort().reverse();
      const detalleDias = diasOrdenados
        .map((d) => `<div class="stat-day-row"><span>${new Date(d).toLocaleDateString('es-CO')}</span><span>${porDia[d]} cortes</span></div>`)
        .join('');
      return `
      <div class="item">
        <div class="info">
          <div class="name">${escapeHtml(liniero)}</div>
          <div class="meta">${porLiniero[liniero]} cortes en el mes · ${diasOrdenados.length} días activos</div>
          <div class="stat-days">${detalleDias}</div>
        </div>
      </div>`;
    })
    .join('');
}

// ---------- Init ----------
loadAll();
(async () => {
  await syncSharedRows(); // pulls the shared Excel from the server if available, overwriting local cache
  if (excelRows.length === 0) {
    showView('upload');
  } else if (selectedRoute) {
    goToDetail();
  } else {
    goToRoutes();
  }
})();

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('service-worker.js').catch(() => {});
}
