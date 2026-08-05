// Backend en producción (Render)
const BACKEND_URL = 'https://consulta-eep-backend.onrender.com';

const LS_ROWS = 'eep_excel_rows';
const LS_ROUTE = 'eep_selected_route';
const LS_RESULTS = 'eep_scan_results';
const LS_LINIERO = 'eep_liniero_nombre';

let excelRows = [];
let scanResults = {}; // { [niu]: { pagado, ultimoPago, estadoPago, fechaRegistroPago, checkedAt, ... } }
let cortes = []; // [{ niu, nombre, direccion, ruta, liniero, tipo, fecha, subidoCortes }]
let compartidos = []; // [{ niu, numeroFactura, fecha }]
let rutaLinieros = {}; // { [ruta]: linieroName }
let selectedRoute = null;
let currentFilter = 'pendiente';
let searchTerm = '';
let pendingCorteRow = null; // row waiting for corte confirmation in the modal

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
      // El Excel a veces trae "TELEFONO" y "CELULAR", o "TELEFONO" repetido dos
      // veces — en ambos casos son números celulares. Se guardan en orden de
      // aparición: el primero como "telefono", el segundo (si existe) como "celular".
      if (phones[0]) row.telefono = phones[0];
      if (phones[1]) row.celular = phones[1];
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
      if (!res.ok) {
        const errText = await res.text().catch(() => '');
        throw new Error(`El servidor respondió ${res.status}. ${errText.slice(0, 150)}`);
      }
      const data = await res.json();
      if (!data.ok) throw new Error(data.message || 'El servidor no pudo procesar el archivo');
      uploadStatus.textContent = `Listo: ${data.nuevos} nuevas, ${data.actualizados} actualizadas, ${data.total} en total ✅`;
      await syncSharedRows();
    } catch (shareErr) {
      // Error real del servidor (o sin conexión) — se lo mostramos claro al
      // usuario en vez de guardar solo localmente sin que se note.
      uploadStatus.textContent = `❌ No se pudo subir al servidor: ${shareErr.message}. Los datos NO se compartieron — intenta de nuevo.`;
      uploadStatus.className = 'upload-status error';
      return;
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
    await fetch(`${BACKEND_URL}/api/reset-mes`, { method: 'POST' });
    excelRows = [];
    scanResults = {};
    selectedRoute = null;
    saveRows();
    saveResults();
    saveRoute();
    showView('upload');
    subtitleText.textContent = 'Sube tu listado para empezar';
  } catch (e) {
    alert('No se pudo resetear — revisa tu conexión e intenta de nuevo.');
  }
});

// ---------- Shared data sync ----------
async function syncSharedRows() {
  try {
    const res = await fetch(`${BACKEND_URL}/api/data/rows`);
    const data = await res.json();
    // Always mirror the server exactly (including empty, e.g. right after a
    // monthly reset) — a conditional "only update if non-empty" left stale
    // cached data stuck forever on devices that already had something loaded.
    excelRows = Array.isArray(data.rows) ? data.rows : [];
    saveRows();
    return excelRows.length > 0;
  } catch (e) {
    // offline or backend down — keep whatever is cached locally
  }
  return false;
}

async function syncSharedResults() {
  try {
    const res = await fetch(`${BACKEND_URL}/api/results`);
    const data = await res.json();
    // Full replace, not merge — the server is authoritative, so a reset there
    // (new month) must actually clear stale local results too.
    scanResults = data.results && typeof data.results === 'object' ? data.results : {};
    saveResults();
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

async function syncCompartidos() {
  try {
    const res = await fetch(`${BACKEND_URL}/api/compartidos`);
    const data = await res.json();
    compartidos = data.compartidos || [];
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

// Returns { status: 'ok'|'no'|'unk'|'notificado'|'suspendido', pagado: bool, corteInvalido: bool }
function evaluarEstado(niu) {
  const scan = scanResults[niu];
  const pagado = !!(scan && scan.pagado);
  const corte = cortadoRecord(niu);

  if (!corte) {
    return { status: pagado ? 'ok' : (scan ? 'no' : 'unk'), pagado, corteInvalido: false };
  }

  if (corte.tipo === 'notificacion') {
    // A notification never physically cuts service — paying makes it "al día"
    // straight away; otherwise it's its own distinct category, not "pendiente".
    return { status: pagado ? 'ok' : 'notificado', pagado, corteInvalido: false };
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

function compartidoRecord(niu, numeroFactura) {
  if (!numeroFactura) return null;
  return compartidos.find((c) => c.niu === niu && c.numeroFactura === numeroFactura) || null;
}

function goToDetail() {
  detailRouteTitle.textContent = `Ruta ${selectedRoute}`;
  subtitleText.textContent = `Ruta ${selectedRoute}`;
  showView('detail');
  Promise.all([syncSharedResults(), syncCortes(), syncCompartidos()]).then(renderDetail);
  renderDetail();

  if (detailRefreshTimer) clearInterval(detailRefreshTimer);
  detailRefreshTimer = setInterval(() => {
    Promise.all([syncSharedResults(), syncCortes(), syncCompartidos()]).then(renderDetail);
  }, 20000); // pick up progress from other devices every 20s
}

function renderDetail() {
  const term = searchTerm.trim().toLowerCase();
  const searching = term.length > 0;
  // A search looks across every route, ignoring the current route + status filter.
  // Without a search, we stay scoped to the selected route as before.
  const baseRows = searching ? excelRows : currentRouteRows();

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
      const badgeClass = s === 'ok' ? 'ok' : s === 'no' ? 'no' : s === 'suspendido' ? 'suspendido' : s === 'notificado' ? 'notificado' : 'unk';
      const badgeLabel = s === 'ok' ? '✅ Al día' : s === 'no' ? '❌ Pendiente' : s === 'suspendido' ? '✂️ Suspendido' : s === 'notificado' ? '📢 Notificado' : '⏳ Sin revisar';
      const scan = scanResults[r.niu];
      const corte = cortadoRecord(r.niu);
      let infoLine;
      if (s === 'suspendido') {
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
      const showCorteBtn = s === 'no' || s === 'unk' || s === 'notificado';
      const showSubidoBtn = s === 'suspendido' && corte && !corte.subidoCortes;
      const subidoBadge = s === 'suspendido' && corte && corte.subidoCortes ? '<span class="uploaded-tag">📤 Reportado</span>' : '';

      const yaCompartido = s === 'ok' && scan && scan.numeroFactura ? compartidoRecord(r.niu, scan.numeroFactura) : null;
      const showCompartirBtn = s === 'ok' && scan && scan.numeroFactura && !yaCompartido;
      const compartidoBadge = yaCompartido ? '<span class="uploaded-tag">✅ Compartido</span>' : '';

      const atraso = parseInt(r.mesesAtrasados, 10) || 0;
      const nameClass = atraso === 0 ? 'name-atraso-0' : atraso === 1 ? 'name-atraso-1' : 'name-atraso-2';
      const corteBtnClass = showCorteBtn && atraso >= 2 ? 'corte-btn corte-btn-urgent' : 'corte-btn';

      return `
      <div class="item" data-niu="${escapeHtml(r.niu)}">
        <div class="info">
          <div class="name ${nameClass}">${escapeHtml(r.nombre || '(sin nombre)')}</div>
          <div class="meta">NIU ${escapeHtml(r.niu)} · ${escapeHtml(r.direccion || '')}${rutaTag}</div>
          <div class="meta2">Medidor ${escapeHtml(r.medidor || '—')} · Sector ${escapeHtml(r.sector || '—')} · Atraso: ${escapeHtml(r.mesesAtrasados || '0')} · Saldo: ${escapeHtml(r.saldoPendiente || '—')}</div>
          <div class="meta3">${infoLine} ${subidoBadge}${compartidoBadge}</div>
          <div class="meta4 phone-row">${renderPhoneField(r.niu, 'telefono', r.telefono)}${renderPhoneField(r.niu, 'celular', r.celular)}</div>
        </div>
        <div class="item-actions">
          <div class="badge ${badgeClass}">${badgeLabel}</div>
          <button class="recheck-btn" data-niu="${escapeHtml(r.niu)}">🔄 Revisar</button>
          ${showCorteBtn ? `<button class="${corteBtnClass}" data-niu="${escapeHtml(r.niu)}">✂️ Suspender</button>` : ''}
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
  listEl.querySelectorAll('.subido-btn').forEach((btn) => {
    btn.addEventListener('click', () => marcarSubido(btn.dataset.id));
  });
  listEl.querySelectorAll('.compartir-btn').forEach((btn) => {
    btn.addEventListener('click', () => compartirComprobante(btn.dataset.niu, btn));
  });
  listEl.querySelectorAll('.phone-input').forEach((input) => {
    input.addEventListener('blur', () => {
      const value = input.value.trim();
      if (value) savePhoneNumber(input.dataset.niu, input.dataset.field, value);
    });
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') input.blur();
    });
  });
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function renderPhoneField(niu, field, value) {
  const label = field === 'telefono' ? '📞' : '📱';
  if (value) {
    return `<a class="phone-link" href="tel:${escapeHtml(value)}">${label} ${escapeHtml(value)}</a>`;
  }
  return `<span class="phone-empty">${label} <input type="tel" class="phone-input" data-niu="${escapeHtml(niu)}" data-field="${field}" placeholder="Agregar ${field}..." /></span>`;
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
  const row = excelRows.find((r) => r.niu === niu);
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
      await fetch(`${BACKEND_URL}/api/compartidos`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ niu, numeroFactura: scan.numeroFactura, liniero }),
      });
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
  const row = excelRows.find((r) => r.niu === niu);
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

corteConfirmBtn.addEventListener('click', async () => {
  const liniero = corteLinieroInput.value;
  if (!liniero) return;
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

async function savePhoneNumber(niu, field, value) {
  const payload = { niu };
  payload[field] = value;
  try {
    const res = await fetch(`${BACKEND_URL}/api/data/telefono`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.message || 'Error desconocido');
    const row = excelRows.find((r) => r.niu === niu);
    if (row) row[field] = value;
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

// Everything except "al día" gets (re)scanned — pendientes, sin revisar, notificados y suspendidos
scanPendingBtn.addEventListener('click', () => {
  const niusList = currentRouteRows()
    .filter((r) => statusOf(r.niu) !== 'ok')
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
    .filter((r) => statusOf(r.niu) !== 'ok')
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
  await Promise.all([syncCortes(), syncCompartidos(), syncSharedResults()]);
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
  excelRows.forEach((r) => {
    const ruta = r.ruta || '(sin ruta)';
    if (!counts[ruta]) counts[ruta] = { total: 0, ok: 0, no: 0, unk: 0, notificado: 0, suspendido: 0 };
    counts[ruta].total++;
    const s = evaluarEstado(r.niu).status;
    if (s === 'ok') counts[ruta].ok++;
    else if (s === 'no') counts[ruta].no++;
    else if (s === 'unk') counts[ruta].unk++;
    else if (s === 'notificado') counts[ruta].notificado++;
    else if (s === 'suspendido') counts[ruta].suspendido++;
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
        </div>
      </div>`;
    })
    .join('');
}

async function syncSharedRowsWithRetry(maxAttempts, delayMs) {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const res = await fetch(`${BACKEND_URL}/api/data/rows`);
      if (!res.ok) throw new Error('status ' + res.status);
      const data = await res.json();
      excelRows = Array.isArray(data.rows) ? data.rows : [];
      saveRows();
      return true; // reached the backend — whatever it says (even 0 rows) is authoritative
    } catch (e) {
      if (attempt < maxAttempts) await new Promise((r) => setTimeout(r, delayMs));
    }
  }
  return false; // never reached the backend — fall back to whatever's cached locally
}

// ---------- Init ----------
loadAll();
(async () => {
  uploadStatus.textContent = 'Conectando con el servidor...';
  uploadStatus.className = 'upload-status';
  // Render's free tier can take 30-50s to wake up on the first request of the
  // day — retry a few times before concluding there's really no data yet,
  // instead of immediately forcing a fresh device to upload its own file.
  const reached = await syncSharedRowsWithRetry(6, 6000);
  uploadStatus.textContent = '';
  if (excelRows.length === 0) {
    if (!reached) {
      uploadStatus.textContent = 'No se pudo conectar con el servidor. Revisa tu conexión, o sube un archivo para empezar.';
      uploadStatus.className = 'upload-status error';
    }
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
