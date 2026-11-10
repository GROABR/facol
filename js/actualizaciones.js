/* ---------- FACOL: actualizaciones automáticas y novedades ----------
   Al abrir, la app de escritorio busca una versión nueva, la descarga sola y muestra un aviso para reiniciar.
   Después de actualizar (o la primera vez que abres una versión nueva) aparecen las novedades. */
const NOVEDADES = [
  {version: '1.7.1', puntos: [
    'La app ahora se llama GROA, con colores acordes a la paleta.',
    'El Excel FACOL ahora trae la hoja FORMATO FACTURA (oculta) idéntica a la del FACOL de ejemplo, con un bloque por factura (hasta 4).',
    'Facturas y complementos: solo muestra los del módulo en el que estás.']},
  {version: '1.7.0', puntos: [
    'La app busca actualizaciones cada vez que la abres; cuando hay una, la descarga y te avisa para reiniciar.',
    'Después de actualizar aparece esta ventana con lo nuevo.',
    'El machote del contrato ya no viene dentro del instalador: súbelo una vez en Ajustes → Machote de contrato.']},
  {version: '1.6.0', puntos: [
    'Módulos por cliente: además de Colorines puedes crear otros clientes (Módulo → + Nuevo cliente…) con su folio, datos y comisiones.',
    'Pestaña Facturas en cada operación para subir facturas y complementos de pago.',
    'Apartado "Facturas y complementos" con buscador y filtro de facturas sin complemento.']},
  {version: '1.5.0', puntos: [
    'Comisiones: % y forma de pago de cada promotor, y sobre qué se calcula (total, Casa u otro promotor).',
    'Se pueden editar las comisiones dentro de cada operación.',
    'El Excel FACOL trae la solicitud de factura (hoja oculta); antes pregunta en cuántas facturas se divide.']},
  {version: '1.4.0', puntos: ['Borrar operaciones, buscador de contratos, subir contratos existentes y usuarios de solo consulta.']},
];
const verNum = v => String(v || '0').split('.').map(n => +n || 0).reduce((a, n) => a * 1000 + n, 0);
async function versionActual() { return window.desk ? desk.version() : NOVEDADES[0].version; }

function mostrarNovedades(desde) {
  const lista = NOVEDADES.filter(n => !desde || verNum(n.version) > verNum(desde));
  if (!lista.length) return;
  modal('Novedades de GROA', lista.map(n => `<h3 style="margin:${n === lista[0] ? 0 : '14px 0 6px'}">Versión ${esc(n.version)}</h3><ul style="margin:0;padding-left:20px">${n.puntos.map(p => `<li style="margin:4px 0">${esc(p)}</li>`).join('')}</ul>`).join(''),
    [{label: 'Entendido', cls: 'primary'}], {width: '640px'});
}

/* ----- aviso de actualización (barra abajo a la derecha) ----- */
function barraActualizacion(s) {
  let b = $('#updBar');
  const ocultar = () => b && b.remove();
  if (!s || ['sin-revisar', 'buscando', 'al-dia', 'desarrollo'].includes(s.estado) || (s.estado === 'error' && !s.manual)) return ocultar();
  if (!b) { b = document.createElement('div'); b.id = 'updBar'; b.className = 'updbar'; document.body.appendChild(b); }
  if (s.estado === 'descargando') b.innerHTML = `<b>Descargando actualización ${esc(s.version || '')}</b><div class="progress" style="margin-top:6px"><i style="width:${s.porcentaje || 3}%"></i></div>`;
  else if (s.estado === 'lista') { b.innerHTML = `<b>Actualización ${esc(s.version)} lista</b><div class="small" style="margin:4px 0 8px">Se instala al reiniciar la app (tarda unos segundos).</div>
      <div style="display:flex;gap:6px"><button class="btn sm primary" id="updGo">Reiniciar y actualizar</button><button class="btn sm" id="updLuego">Después</button></div>`;
    $('#updGo').onclick = () => desk.instalarActualizacion(); $('#updLuego').onclick = ocultar; }
  else if (s.estado === 'error') { b.innerHTML = `<b>No se pudo buscar actualizaciones</b><div class="small">${esc(s.mensaje || '')}</div><button class="btn sm" id="updX" style="margin-top:6px">Cerrar</button>`; $('#updX').onclick = ocultar; }
}

async function iniciarActualizaciones() {
  // novedades: la primera vez que se abre una versión nueva
  const v = await versionActual(); let vista = null;
  try { vista = localStorage.getItem('facol-version-vista'); localStorage.setItem('facol-version-vista', v); } catch { /* sin almacenamiento */ }
  // sin registro previo (primera vez en este equipo): se muestran solo las novedades de esta versión
  if (!vista || verNum(v) > verNum(vista)) setTimeout(() => { if (!dlg.open) mostrarNovedades(vista || NOVEDADES[1] && NOVEDADES[1].version); }, 600);
  if (!window.desk || !desk.onActualizacion) return;
  desk.onActualizacion(barraActualizacion);
  barraActualizacion(await desk.estadoActualizacion());
}

/** Panel en Ajustes. */
async function panelActualizaciones(el) {
  const v = await versionActual();
  el.innerHTML = `<p style="margin-top:0">Versión instalada: <b>${esc(v)}</b></p>
    ${window.desk ? '<p class="small muted">La app revisa sola cada vez que la abres. Si hay una versión nueva la descarga y te avisa para reiniciar.</p><div id="updEst" class="small"></div>' : '<p class="small muted">En la versión web siempre ves lo más reciente (si no, recarga con Ctrl + F5).</p>'}
    <div class="toolbar" style="margin:8px 0 0">${window.desk ? '<button class="btn" id="updBuscar">Buscar actualizaciones ahora</button>' : ''}<button class="btn" id="updNov">Ver novedades</button></div>`;
  $('#updNov').onclick = () => mostrarNovedades(null);
  if (!window.desk) return;
  const pinta = s => { const t = {buscando: 'Buscando…', 'al-dia': '✓ Tienes la versión más reciente', descargando: `Descargando ${s.version || ''} ${s.porcentaje ? s.porcentaje + '%' : ''}`,
    lista: `Versión ${s.version} lista: reinicia para instalarla`, error: 'Error: ' + (s.mensaje || ''), desarrollo: 'Estás usando la app desde la carpeta del proyecto (sin actualizaciones automáticas).', 'sin-revisar': ''}[s.estado] || '';
    if ($('#updEst')) $('#updEst').textContent = t; };
  pinta(await desk.estadoActualizacion()); desk.onActualizacion(pinta);
  $('#updBuscar').onclick = async () => pinta(await desk.buscarActualizacion());
}
