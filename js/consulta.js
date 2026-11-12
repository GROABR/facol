/* ---------- FACOL: portal de solo consulta (contratos, beneficiarios, facturas y comprobantes) ----------
   Estas pantallas no tienen botones para capturar, editar ni borrar. La base de datos además
   lo impide para el rol "consulta" (ver supabase/migracion-roles.sql). */

/** Tabla con buscador y orden, reutilizada por las cuatro pantallas. */
function listaConsulta(main, {titulo, filas, columnas, buscar, ordenes, vacio, alDibujar}) {
  main.innerHTML = `<h2>${esc(titulo)}</h2>
    <div class="toolbar"><input id="cq" placeholder="Buscar…" style="width:300px"><select id="co">${ordenes.map(([k, l]) => `<option value="${k}">${l}</option>`).join('')}</select>
      <span class="sp"></span><span class="muted small" id="cn"></span></div>
    <div class="tablewrap tall"><table><thead><tr>${columnas.map(c => `<th class="${c.n ? 'n' : ''}">${c.h}</th>`).join('')}</tr></thead><tbody id="ct"></tbody></table></div>`;
  const draw = () => {
    const q = Logica.normName(val('cq')), o = ordenes.find(x => x[0] === val('co'));
    const rows = filas.filter(f => !q || Logica.normName(buscar(f)).includes(q)).sort(o[2]);
    $('#cn').textContent = `${rows.length} de ${filas.length}`;
    $('#ct').innerHTML = rows.slice(0, 500).map(f => `<tr>${columnas.map(c => `<td class="${c.n ? 'n' : ''}">${c.v(f)}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${columnas.length}" class="empty">${filas.length ? 'Sin resultados' : vacio}</td></tr>`;
    if (alDibujar) alDibujar(rows);
  };
  $('#cq').oninput = draw; $('#co').onchange = draw; draw();
}
const porNombre = k => [(a, b) => String(a[k]).localeCompare(String(b[k]), 'es'), (a, b) => String(b[k]).localeCompare(String(a[k]), 'es')];
const porFecha = k => [(a, b) => String(b[k]).localeCompare(String(a[k])), (a, b) => String(a[k]).localeCompare(String(b[k]))];
function botonesArchivo(main, filas, nombreDe) {
  $$('[data-ver]', main).forEach(b => b.onclick = () => { const f = filas.find(x => x.id === b.dataset.ver); guard(() => verArchivo(f.path, nombreDe(f))); });
  $$('[data-dl]', main).forEach(b => b.onclick = () => { const f = filas.find(x => x.id === b.dataset.dl); guard(() => descargarArchivo(f.path, nombreDe(f))); });
}
const accionesArchivo = f => `<button class="link" data-ver="${f.id}">Ver</button> · <button class="link" data-dl="${f.id}">Descargar</button>`;

VIEWS.consultaContratos = async main => {
  const [cons, bens] = await Promise.all([db.list('contratos', {order: 'created_at', desc: true}), cat('beneficiarios', true)]);
  const filas = cons.map(c => { const b = bens.find(x => x.id === c.beneficiario_id) || {}; return {...c, nombre: b.nombre || '(sin beneficiario)', rfc: b.rfc || '', curp: b.curp || ''}; });
  const nombreDe = f => `${f.nombre}.${extDe(f.path)}`;
  const [az, za] = porNombre('nombre'), [rec, ant] = porFecha('created_at');
  listaConsulta(main, {titulo: 'Contratos', filas, vacio: 'Aún no hay contratos', buscar: f => [f.nombre, f.rfc, f.curp].join(' '),
    ordenes: [['az', 'Nombre A → Z', az], ['za', 'Nombre Z → A', za], ['rec', 'Más recientes', rec], ['ant', 'Más antiguos', ant]],
    columnas: [{h: 'Socio', v: f => `<b>${esc(f.nombre)}</b>${f.rfc ? `<div class="small muted">${esc(f.rfc)}</div>` : ''}`}, {h: 'Fecha inicio', v: f => fmtDay(f.fecha_inicio)},
      {h: 'Tipo', v: f => tag('', extDe(f.path).toUpperCase())}, {h: 'Registrado', v: f => fmtDate(f.created_at)}, {h: '', v: accionesArchivo}],
    alDibujar: () => botonesArchivo(main, filas, nombreDe)});
};

VIEWS.consultaBeneficiarios = async main => {
  const bens = (await cat('beneficiarios', true)).filter(b => b.activo !== false);
  const [az, za] = porNombre('nombre');
  listaConsulta(main, {titulo: 'Beneficiarios', filas: bens, vacio: 'La base está vacía',
    buscar: b => [b.nombre, b.nombre_alt, b.rfc, b.curp, b.clabe, b.cuenta].join(' '),
    ordenes: [['az', 'Nombre A → Z', az], ['za', 'Nombre Z → A', za]],
    columnas: [{h: 'Beneficiario', v: b => `<b>${esc(b.nombre)}</b>`}, {h: 'Banco', v: b => esc(b.banco || '')}, {h: 'Cuenta', v: b => `<span class="mono">${esc(b.cuenta || '')}</span>`},
      {h: 'CLABE', v: b => `<span class="mono">${esc(b.clabe || '')}</span>`}, {h: 'RFC', v: b => esc(b.rfc || '')}, {h: 'CURP', v: b => esc(b.curp || '')}]});
};

async function consultaArchivos(main, tipo, titulo) {
  const arch = await db.list('archivos', {eq: {tipo}, order: 'created_at', desc: true});
  const opIds = [...new Set(arch.map(a => a.operacion_id).filter(Boolean))], dIds = [...new Set(arch.map(a => a.dispersion_id).filter(Boolean))];
  const [ops, disp] = await Promise.all([opIds.length ? db.list('operaciones', {in: {id: opIds}, select: 'id,folio,cliente_nombre,fecha_deposito,modalidad'}) : [],
    dIds.length ? db.list('dispersiones', {in: {id: dIds}, select: 'id,nombre,nombre_archivo'}) : []]);
  const filas = arch.map(a => { const o = ops.find(x => x.id === a.operacion_id) || {}; const d = disp.find(x => x.id === a.dispersion_id);
    return {...a, folio: o.folio || '', cliente: o.cliente_nombre || '', fecha_op: o.fecha_deposito || '', persona: d ? (d.nombre || d.nombre_archivo) : ''}; });
  const [rec, ant] = porFecha('created_at'), [fa, fz] = porNombre('folio');
  listaConsulta(main, {titulo, filas, vacio: 'Aún no hay archivos', buscar: f => [f.folio, f.cliente, f.persona, f.nombre].join(' '),
    ordenes: [['rec', 'Más recientes', rec], ['ant', 'Más antiguos', ant], ['fz', 'Folio (mayor a menor)', fz], ['fa', 'Folio (menor a mayor)', fa]],
    columnas: [{h: 'Folio', v: f => `<b>${esc(f.folio)}</b>`}, {h: 'Fecha operación', v: f => fmtDay(f.fecha_op)}, {h: 'Cliente', v: f => esc(f.cliente)},
      ...(tipo === 'comprobante' ? [{h: 'Colaborador', v: f => esc(f.persona || 'general')}] : []),
      {h: 'Archivo', v: f => esc(f.nombre)}, {h: 'Subido', v: f => fmtDate(f.created_at)}, {h: '', v: accionesArchivo}],
    alDibujar: () => botonesArchivo(main, filas, f => f.nombre)});
}
VIEWS.consultaFacturas = main => consultaArchivos(main, 'factura', 'Facturas');

/* ===================== Enviar cuestionarios (también para usuarios de solo consulta) ===================== */
VIEWS.consultaEnvios = async main => {
  const url = await urlFormularios();
  const mias = (await db.list('invitaciones', {order: 'created_at', desc: true})).filter(i => esConsulta() ? i.created_by === me.email : true).slice(0, 100);
  const correoOk = c => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c);
  main.innerHTML = `<h2>Enviar cuestionarios</h2>
    <div class="cols">
      <div class="panel"><h3 style="margin-top:0">Cuestionario de alta</h3>
        <p class="small muted" style="margin-top:0">Le llega por correo una liga personal para llenar sus datos, firmar y subir su huella e INE. El equipo revisa y genera su contrato.</p>
        <div class="form"><label>Nombre</label><input id="eaN"><label>Correo</label><input id="eaC" type="email"><label>Celular (opcional)</label><input id="eaT"></div>
        <div class="err" id="eaE"></div><div class="toolbar" style="margin:6px 0 0"><span class="sp"></span><button class="btn primary" id="eaGo">Enviar cuestionario</button></div></div>
      <div class="panel"><h3 style="margin-top:0">Cambio de cuenta</h3>
        <p class="small muted" style="margin-top:0">Le llega por correo la liga para registrar su cuenta nueva. El cambio se aplica hasta que el equipo lo revisa.</p>
        <div class="form"><label>Nombre (opcional)</label><input id="ecN"><label>Correo</label><input id="ecC" type="email"></div>
        <div class="err" id="ecE"></div><div class="toolbar" style="margin:6px 0 0"><button class="btn" id="ecCopy">Copiar liga</button><span class="sp"></span><button class="btn primary" id="ecGo">Enviar liga de cambio</button></div></div>
    </div>
    <h3>${esConsulta() ? 'Cuestionarios que has enviado' : 'Últimos cuestionarios enviados'}</h3>
    <div class="tablewrap"><table><thead><tr><th>Enviado</th><th>Para</th><th>Correo</th><th>Estado</th><th></th></tr></thead><tbody>
    ${mias.map(i => `<tr><td>${fmtDate(i.created_at)}</td><td><b>${esc(i.nombre || '')}</b></td><td>${esc(i.correo || '')}</td>
      <td>${tag(i.estado, {enviada: 'enviado, sin contestar', respondida: 'contestado', aprobada: 'aprobado', contrato: 'con contrato', cancelada: 'cancelado'}[i.estado])}</td>
      <td>${i.estado === 'enviada' ? `<button class="link" data-re="${i.id}">Reenviar</button>` : ''}</td></tr>`).join('') || '<tr><td colspan="5" class="empty">Aún no has enviado cuestionarios</td></tr>'}
    </tbody></table></div>`;
  $('#eaGo').onclick = async () => {
    const correo = val('eaC').toLowerCase(); if (!val('eaN') || !correoOk(correo)) return ($('#eaE').textContent = 'Escribe el nombre y un correo válido');
    const b = $('#eaGo'); busy(b, 'Enviando…');
    try { const i = await db.insert('invitaciones', {nombre: val('eaN').toUpperCase(), correo, celular: val('eaT').replace(/\D/g, ''), created_by: me.email}); avisoEnvio(await enviarCorreoAlta(i, url)); refresh(); }
    catch (e) { busy(b); $('#eaE').textContent = errMsg(e); }
  };
  $('#ecGo').onclick = async () => {
    const correo = val('ecC').toLowerCase(); if (!correoOk(correo)) return ($('#ecE').textContent = 'Escribe un correo válido');
    const b = $('#ecGo'); busy(b, 'Enviando…');
    try { const r = await enviarCorreoCambio(correo, val('ecN'), url); toast(r === 'enviado' ? 'Liga de cambio de cuenta enviada' : 'Se abrió tu correo con el mensaje listo: solo presiona Enviar'); $('#ecC').value = ''; $('#ecN').value = ''; }
    catch (e) { $('#ecE').textContent = errMsg(e); }
    busy(b);
  };
  $('#ecCopy').onclick = () => { navigator.clipboard.writeText(`${url}/cambio.html`); toast('Liga copiada'); };
  $$('[data-re]', main).forEach(b => b.onclick = () => pedirCorreoYEnviar(mias.find(i => i.id === b.dataset.re), url));
};

/* ===================== Facturas y complementos de pago (todas las operaciones) ===================== */
VIEWS.facturas = async main => {
  const arch = await db.list('archivos', {in: {tipo: ['factura', 'complemento']}, order: 'created_at', desc: true});
  const opIds = [...new Set(arch.map(a => a.operacion_id).filter(Boolean))];
  const ops = opIds.length ? await db.list('operaciones', {in: {id: opIds}, select: 'id,folio,cliente_nombre,razon_social,fecha_deposito,modalidad,monto_deposito,estado'}) : [];
  const conComp = new Set(arch.filter(a => a.tipo === 'complemento').map(a => a.operacion_id));
  const filas = arch.map(a => { const o = ops.find(x => x.id === a.operacion_id) || {}; return {...a, o, folio: o.folio || '', cliente: o.razon_social || o.cliente_nombre || '', mod: o.modalidad || ''}; });
  // solo las del módulo: el equipo ve las del módulo en el que está; consulta elige el módulo (sin "todos")
  const mods = esConsulta() ? [...new Set(filas.map(f => f.mod).filter(Boolean))] : [modulo];
  if (!mods.length) mods.push(modulo);
  main.innerHTML = `<h2>Facturas y complementos de pago${esConsulta() ? '' : ' · ' + esc(nombreModulo(modulo))}</h2>
    <div class="toolbar"><input id="fq" placeholder="Buscar folio, cliente o archivo" style="width:260px">
      <select id="ft"><option value="">Facturas y complementos</option><option value="factura">Solo facturas</option><option value="complemento">Solo complementos</option></select>
      <select id="fm" ${esConsulta() ? '' : 'hidden'}>${mods.map(m => `<option value="${esc(m)}" ${m === modulo ? 'selected' : ''}>${esc(nombreModulo(m))}</option>`).join('')}</select>
      <label class="muted">Del</label><input type="date" id="fd"><label class="muted">al</label><input type="date" id="fh">
      <label><input type="checkbox" id="fs"> Sin complemento</label><span class="sp"></span><span class="muted small" id="fn"></span></div>
    <div class="tablewrap tall"><table><thead><tr><th>Folio</th><th>Fecha operación</th><th>Cliente</th><th>Tipo</th><th>Archivo</th><th>Subido</th><th></th></tr></thead><tbody id="ftb"></tbody></table></div>`;
  const draw = () => {
    const q = Logica.normName(val('fq')), t = val('ft'), m = val('fm'), d = val('fd'), h = val('fh'), sin = val('fs');
    const rows = filas.filter(f => (!t || f.tipo === t) && f.mod === m && (!d || (f.o.fecha_deposito || '') >= d) && (!h || (f.o.fecha_deposito || '') <= h)
      && (!sin || (f.tipo === 'factura' && !conComp.has(f.operacion_id))) && (!q || Logica.normName([f.folio, f.cliente, f.nombre].join(' ')).includes(q)));
    $('#fn').textContent = `${rows.length} archivo(s)`;
    $('#ftb').innerHTML = rows.slice(0, 500).map(f => `<tr><td>${esConsulta() || !f.o.id ? `<b>${esc(f.folio)}</b>` : `<button class="link" data-op="${f.o.id}"><b>${esc(f.folio)}</b></button>`}</td>
      <td>${fmtDay(f.o.fecha_deposito)}</td><td>${esc(f.cliente)}</td>
      <td>${f.tipo === 'factura' ? tag('info', 'factura') + (conComp.has(f.operacion_id) ? '' : ' ' + tag('warn', 'sin complemento')) : tag('ok', 'complemento')}</td>
      <td>${esc(f.nombre)}</td><td>${fmtDate(f.created_at)}</td><td style="white-space:nowrap">${accionesArchivo(f)}</td></tr>`).join('')
      || `<tr><td colspan="7" class="empty">${filas.length ? 'Sin resultados' : 'Aún no se han subido facturas ni complementos. Se suben en la pestaña Facturas de cada operación.'}</td></tr>`;
    botonesArchivo(main, filas, f => f.nombre);
    $$('[data-op]', main).forEach(b => b.onclick = () => go('operacion', {id: b.dataset.op, tab: 'facturas'}));
  };
  ['fq', 'ft', 'fm', 'fd', 'fh', 'fs'].forEach(id => { $('#' + id).oninput = draw; $('#' + id).onchange = draw; }); draw();
};
VIEWS.consultaComprobantes = main => consultaArchivos(main, 'comprobante', 'Comprobantes de dispersión');
