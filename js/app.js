/* ---------- FACOL: núcleo de la app (utilidades, sesión, menú) ---------- */
const $ = (s, el = document) => el.querySelector(s), $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
const money = n => (+n || 0).toLocaleString('es-MX', {style: 'currency', currency: 'MXN'});
const pct = n => (+n || 0).toLocaleString('es-MX', {maximumFractionDigits: 4}) + '%';
const today = () => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
const fmtDay = d => { if (!d) return '—'; const [y, m, dd] = String(d).slice(0, 10).split('-'); return `${dd}/${m}/${y}`; };
const fmtDate = d => d ? new Date(d).toLocaleString('es-MX', {dateStyle: 'short', timeStyle: 'short'}) : '—';
const fechaLarga = d => { if (!d) return ''; const [y, m, dd] = String(d).slice(0, 10).split('-'); return `${+dd} DE ${['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'][+m - 1]} DE ${y}`; };
const val = (id, el = document) => { const e = el.querySelector('#' + id); return e ? (e.type === 'checkbox' ? e.checked : e.value.trim()) : ''; };
const num = (id, el) => Logica.toNum(val(id, el));
const r2 = Logica.r2;
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), 3200); }
const errMsg = e => {
  const m = (e && (e.message || e.error_description)) || String(e || '');
  if (/invalid login credentials/i.test(m)) return 'Correo o contraseña incorrectos';
  if (/failed to fetch|network/i.test(m)) return 'Sin conexión con la base de datos';
  if (/row-level security|permission denied/i.test(m)) return 'Tu usuario no tiene permiso (debe estar en la tabla staff)';
  if (/duplicate key/i.test(m)) return 'Ya existe un registro con ese dato';
  return m;
};
async function guard(fn, okMsg) { try { const r = await fn(); if (okMsg) toast(okMsg); return r; } catch (e) { console.error(e); toast(errMsg(e)); return undefined; } }
function busy(btn, text) { if (!btn) return; if (text) { btn.dataset.t = btn.dataset.t || btn.textContent; btn.textContent = text; btn.disabled = true; } else { btn.textContent = btn.dataset.t || btn.textContent; btn.disabled = false; } }

/* ----- ventanas modales ----- */
const dlg = $('#dlg');
function modal(title, body, buttons = [], opts = {}) {
  dlg.onclose = null; dlg.style.width = opts.width || '';
  dlg.innerHTML = `<div class="hd">${esc(title)}</div><div class="bd">${body}</div><div class="ft"></div>`;
  buttons.forEach(b => {
    const el = document.createElement('button'); el.className = 'btn ' + (b.cls || ''); el.textContent = b.label;
    el.onclick = async () => { if (b.onClick) { busy(el, '…'); let r; try { r = await b.onClick(dlg); } catch (e) { console.error(e); toast(errMsg(e)); r = false; } busy(el); if (r === false) return; } dlg.close(); };
    dlg.querySelector('.ft').appendChild(el);
  });
  if (!dlg.open) dlg.showModal();
  return dlg;
}
const confirmar = (title, text, label = 'Sí, continuar', cls = 'primary') => new Promise(res => {
  modal(title, `<p>${text}</p>`, [{label: 'Cancelar', onClick: () => res(false)}, {label, cls, onClick: () => res(true)}]);
  dlg.onclose = () => res(false);
});

/* ----- archivos ----- */
const safePath = s => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w.\-]+/g, '_');
function pickFiles(accept = '', multiple = false) {
  return new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.multiple = multiple; i.onchange = () => res([...i.files]); i.click(); });
}
/** zona para arrastrar o elegir archivos; onFiles(files) */
function dropZone(el, onFiles, {accept = '', multiple = true, text = 'Arrastra aquí los archivos o haz clic para elegirlos'} = {}) {
  el.classList.add('drop'); el.innerHTML = esc(text);
  el.onclick = async () => { const f = await pickFiles(accept, multiple); if (f.length) onFiles(f); };
  el.ondragover = e => { e.preventDefault(); el.classList.add('over'); };
  el.ondragleave = () => el.classList.remove('over');
  el.ondrop = e => { e.preventDefault(); el.classList.remove('over'); const f = [...e.dataTransfer.files]; if (f.length) onFiles(multiple ? f : [f[0]]); };
}
async function readSheetRows(file, sheet) {
  const wb = XLSX.read(await file.arrayBuffer(), {type: 'array', cellDates: false});
  const name = sheet && wb.SheetNames.includes(sheet) ? sheet : wb.SheetNames[0];
  return {rows: XLSX.utils.sheet_to_json(wb.Sheets[name], {header: 1, defval: '', raw: true}), sheets: wb.SheetNames, wb};
}
/** guarda un archivo en el almacenamiento y lo registra en la tabla archivos */
async function subirArchivo(file, meta) {
  const carpeta = meta.operacion_id ? `operaciones/${meta.operacion_id}` : meta.beneficiario_id ? `beneficiarios/${meta.beneficiario_id}` : 'otros';
  const path = `${carpeta}/${meta.tipo}/${Date.now()}_${safePath(file.name)}`;
  await db.upload(path, file);
  return db.insert('archivos', {...meta, nombre: file.name, path, size: file.size});
}
async function descargarArchivo(path, nombre) {
  const blob = await db.download(path);
  if (window.desk) { const p = await desk.saveFile(new Uint8Array(await blob.arrayBuffer()), nombre); if (p) { toast('Guardado'); desk.openFile(p); } return; }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombre; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
async function verArchivo(path, nombre) {
  const blob = await db.download(path);
  if (/\.(png|jpe?g|gif|webp|pdf)$/i.test(nombre)) {
    const url = URL.createObjectURL(blob);
    const body = /\.pdf$/i.test(nombre) ? `<iframe src="${url}" style="width:100%;height:68vh;border:0"></iframe>` : `<img src="${url}" style="max-width:100%">`;
    modal(nombre, body, [{label: 'Descargar', onClick: () => { descargarArchivo(path, nombre); return false; }}, {label: 'Cerrar', cls: 'primary'}], {width: '900px'});
    dlg.onclose = () => URL.revokeObjectURL(url);
  } else descargarArchivo(path, nombre);
}
async function guardarBytes(bytes, nombre) {
  if (window.desk) { const p = await desk.saveFile(bytes, nombre); if (p) { toast('Guardado en ' + p); desk.openFile(p); } return p; }
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([bytes])); a.download = nombre; a.click(); return nombre;
}
function guardarXlsx(wb, nombre) { return guardarBytes(new Uint8Array(XLSX.write(wb, {bookType: 'xlsx', type: 'array'})), nombre); }
async function guardarPdf(html, nombre) {
  if (window.desk) { const p = await desk.savePdf(html, nombre); if (p) { toast('PDF guardado'); desk.openFile(p); } return p; }
  const w = window.open('', '_blank'); if (w) { w.document.write(html); w.document.close(); w.print(); }
}
/** hoja de Excel desde arreglo de objetos; las columnas en "texto" se guardan como texto (CLABE, cuentas) */
function hoja(rows, cols, texto = []) {
  const ws = XLSX.utils.aoa_to_sheet([cols.map(c => c.h), ...rows.map(r => cols.map(c => { const v = typeof c.v === 'function' ? c.v(r) : r[c.v]; return v ?? ''; }))]);
  const range = XLSX.utils.decode_range(ws['!ref']);
  cols.forEach((c, j) => {
    for (let i = 1; i <= range.e.r; i++) {
      const cell = ws[XLSX.utils.encode_cell({r: i, c: j})]; if (!cell) continue;
      if (texto.includes(c.h)) { cell.t = 's'; cell.v = String(cell.v); }
      else if (c.money && typeof cell.v === 'number') cell.z = '"$"#,##0.00';
    }
  });
  ws['!cols'] = cols.map(c => ({wch: c.w || Math.max(10, c.h.length + 2)}));
  return ws;
}

const FORMS_URL = 'https://gmoro27.github.io/facol-formularios';   // formularios públicos publicados en GitHub Pages

/* ----- catálogos en memoria ----- */
const cache = {};
async function cat(table, force) {
  if (!cache[table] || force) cache[table] = await db.list(table, {order: table === 'beneficiarios' ? 'nombre' : table === 'config' ? 'key' : 'nombre'});
  return cache[table];
}
const invalidate = (...t) => t.forEach(x => delete cache[x]);
async function getConfig(key, def) { const c = (await cat('config', true)).find(x => x.key === key); return c ? c.value : def; }
async function setConfig(key, value) { await db.upsert('config', [{key, value}], 'key'); invalidate('config'); }

/* ----- estados ----- */
const TAG = {
  abierta: 'info', cerrada: 'ok', cancelada: 'bad', pendiente: 'warn', dispersado: 'ok', rechazado: 'bad', solicitada: 'info', emitida: 'ok', no_aplica: '',
  enviada: 'info', respondida: 'warn', aprobada: 'ok', contrato: 'ok', aplicado: 'ok',
};
const tag = (s, label) => `<span class="tag ${TAG[s] ?? ''}">${esc(label || String(s || '').replace('_', ' '))}</span>`;

/* ===================== sesión y menú ===================== */
let me = null, view = 'inicio', viewArg = null;
const VIEWS = {};                      // cada módulo registra sus pantallas aquí
/* Cada módulo (Colorines / Retornos) tiene su propio menú. */
const MODULOS = {
  colorines: {label: 'Colorines', nav: [
    {sec: 'Operación'},
    {id: 'inicio', label: 'Tablero'},
    {id: 'colorines', label: 'Dispersiones'},
    {id: 'comisiones', label: 'Reporte de comisiones'},
    {sec: 'Personas'},
    {id: 'beneficiarios', label: 'Beneficiarios (BD)'},
    {id: 'contratos', label: 'Altas y contratos', badge: 'altas'},
    {id: 'cambios', label: 'Cambios de cuenta', badge: 'cambios'},
    {sec: 'Configuración'},
    {id: 'esquemaColorines', label: 'Datos y comisiones'},
    {id: 'comisionistas', label: 'Promotores'},
    {id: 'ajustes', label: 'Ajustes'},
  ]},
  retorno: {label: 'Retornos', nav: [
    {sec: 'Operación'},
    {id: 'inicio', label: 'Tablero'},
    {id: 'depositos', label: 'Depósitos y retornos'},
    {id: 'comisiones', label: 'Reporte de comisiones'},
    {sec: 'Configuración'},
    {id: 'clientes', label: 'Clientes y comisiones'},
    {id: 'comisionistas', label: 'Promotores'},
    {id: 'empresas', label: 'Empresas'},
    {id: 'ajustes', label: 'Ajustes'},
  ]},
};
// usuarios de solo consulta: ven contratos, la BD de beneficiarios, facturas y comprobantes
const NAV_CONSULTA = [
  {sec: 'Consulta'},
  {id: 'consultaContratos', label: 'Contratos'},
  {id: 'consultaBeneficiarios', label: 'Beneficiarios (BD)'},
  {id: 'consultaFacturas', label: 'Facturas'},
  {id: 'consultaComprobantes', label: 'Comprobantes de dispersión'},
];
const esConsulta = () => me && me.rol === 'consulta';
let modulo = (() => { try { return MODULOS[localStorage.getItem('facol-modulo')] ? localStorage.getItem('facol-modulo') : 'colorines'; } catch { return 'colorines'; } })();
function setModulo(m) { modulo = m; try { localStorage.setItem('facol-modulo', m); } catch { /* sin almacenamiento */ } go('inicio'); }
async function renderNav() {
  const badges = {};
  if (esConsulta()) {
    $('#nav').innerHTML = NAV_CONSULTA.map(n => n.sec ? `<div class="sec">${n.sec}</div>` : `<button data-v="${n.id}" class="${view === n.id ? 'on' : ''}">${esc(n.label)}</button>`).join('');
    $$('#nav button[data-v]').forEach(b => b.onclick = () => go(b.dataset.v)); return;
  }
  try {
    const [a, c] = await Promise.all([db.list('invitaciones', {eq: {estado: 'respondida'}, select: 'id'}), db.list('cambios_cuenta', {eq: {estado: 'pendiente'}, select: 'id'})]);
    badges.altas = a.length; badges.cambios = c.length;
  } catch { /* sin conexión */ }
  $('#nav').innerHTML = `<div class="mods">${Object.entries(MODULOS).map(([k, m]) => `<button data-m="${k}" class="${k === modulo ? 'on' : ''}">${m.label}</button>`).join('')}</div>` +
    MODULOS[modulo].nav.map(n => n.sec ? `<div class="sec">${n.sec}</div>` :
    `<button data-v="${n.id}" class="${view === n.id ? 'on' : ''}">${esc(n.label)}${n.badge && badges[n.badge] ? `<span class="badge">${badges[n.badge]}</span>` : ''}</button>`).join('');
  $$('#nav button[data-v]').forEach(b => b.onclick = () => go(b.dataset.v));
  $$('#nav button[data-m]').forEach(b => b.onclick = () => setModulo(b.dataset.m));
}
async function go(v, arg) {
  if (esConsulta() && !NAV_CONSULTA.some(n => n.id === v)) v = 'consultaContratos';   // consulta no entra a pantallas de captura
  view = v; viewArg = arg ?? null; $$('#nav button').forEach(b => b.classList.toggle('on', b.dataset.v === v));
  const main = $('#main'); main.innerHTML = '<div class="empty">Cargando…</div>'; main.scrollTop = 0;
  try { await VIEWS[v](main, arg); } catch (e) { console.error(e); main.innerHTML = `<div class="panel"><b>No se pudo cargar.</b><div class="err">${esc(errMsg(e))}</div></div>`; }
  renderNav();
}
const refresh = () => go(view, viewArg);

function loginBox(inner) { const L = $('#login'); L.style.display = 'flex'; L.innerHTML = `<div class="box"><h1>FACOL</h1>${inner}</div>`; }
function renderConexion(msg = '') {
  const c = dbConn.get() || {};
  loginBox(`<p class="muted" style="margin-top:0">Conecta la app con tu proyecto de Supabase (Project Settings → API).</p>
    <div class="form"><label>URL</label><input id="cUrl" placeholder="https://xxxx.supabase.co" value="${esc(c.url || '')}">
    <label>Clave pública</label><input id="cKey" placeholder="sb_publishable_… o anon key" value="${esc(c.key || '')}"></div>
    <div class="err">${esc(msg)}</div>
    <button class="btn primary" id="cGo">Conectar</button>
    <p class="small muted" style="margin-bottom:0">¿Solo quieres probarla? <button class="link" id="cDemo">Entrar en modo demostración</button> (los datos se quedan en esta PC).</p>`);
  $('#cGo').onclick = () => {
    const url = val('cUrl').replace(/\/+$/, ''), key = val('cKey');
    if (!/^https:\/\/[\w-]+\.supabase\.co$/.test(url) || key.length < 20) return renderConexion('Revisa la URL y la clave');
    dbConn.set({mode: 'supabase', url, key});
  };
  $('#cDemo').onclick = () => dbConn.set({mode: 'demo'});
}
function renderLogin(msg = '') {
  $('#side').style.display = 'none'; $('#main').innerHTML = '';
  loginBox(`<p class="muted" style="margin-top:0">${db.mode === 'demo' ? 'Modo demostración: escribe cualquier correo.' : 'Entra con tu usuario.'}</p>
    <div class="form"><label>Correo</label><input id="lEmail" type="email"><label>Contraseña</label><input id="lPwd" type="password"></div>
    <div class="err" id="lErr">${esc(msg)}</div><button class="btn primary" id="lGo">Entrar</button>
    <div style="display:flex;justify-content:space-between;margin-top:8px"><button class="link" id="lForgot">¿Olvidaste tu contraseña?</button><button class="link" id="lConn">Cambiar conexión</button></div>`);
  const goIn = async () => {
    const b = $('#lGo'); busy(b, 'Entrando…');
    try { await db.signIn(val('lEmail').toLowerCase(), $('#lPwd').value); boot(); } catch (e) { busy(b); $('#lErr').textContent = errMsg(e); }
  };
  $('#lGo').onclick = goIn; $('#lPwd').onkeydown = e => { if (e.key === 'Enter') goIn(); };
  $('#lForgot').onclick = async () => { const e = val('lEmail'); if (!e) return ($('#lErr').textContent = 'Escribe tu correo'); await guard(() => db.resetPassword(e), 'Te enviamos un correo para cambiar la contraseña'); };
  $('#lConn').onclick = () => renderConexion();
  $('#lEmail').focus();
}
async function boot() {
  if (!window.db) return renderConexion();
  const s = await db.session().catch(() => null);
  if (!s) return renderLogin();
  if (!(await db.isStaff().catch(() => false))) { await db.signOut(); return renderLogin('Tu usuario no tiene acceso. Pide que agreguen tu correo a la tabla staff.'); }
  me = s; me.rol = await db.rol().catch(() => 'operador');
  $('#login').style.display = 'none'; $('#side').style.display = 'flex';
  $('#me').innerHTML = `${db.mode === 'demo' ? '<span class="demo">DEMO</span><br>' : ''}${esc(me.email)}${esConsulta() ? '<br><span class="small">solo consulta</span>' : ''}<br><button id="out">Salir</button>`;
  $('#out').onclick = async () => { await db.signOut(); location.reload(); };
  go(esConsulta() ? 'consultaContratos' : 'inicio');
}
