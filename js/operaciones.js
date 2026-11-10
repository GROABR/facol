/* ---------- FACOL: depósitos (operaciones FACOL), detalle, ficha de retorno y reporte de comisiones ---------- */
const MODALIDAD = new Proxy({}, {get: (_, m) => nombreModulo(m)});
const TIPOS_ARCHIVO = {ficha_deposito: 'Ficha de depósito / pago', solicitud: 'Formato de solicitud', complemento: 'Complemento de pago', comprobante: 'Comprobante de transferencia', factura: 'Factura', ficha_retorno: 'Ficha de retorno', otro: 'Otro'};

/* ===================== listado ===================== */
function listaOperaciones(modalidad) {
  return async (main, arg) => {
    const f = Object.assign({desde: '', hasta: '', estado: 'abierta', cliente: '', q: ''}, listaOperaciones.f[modalidad] || {}, arg || {});
    const clientes = await cat('clientes');
    const titulo = esDisp(modalidad) ? `${nombreModulo(modalidad)} · dispersiones` : 'Depósitos y retornos';
    main.innerHTML = `<h2>${titulo}</h2>
      <div class="toolbar">
        <input id="fq" placeholder="Folio o cliente" style="width:170px" value="${esc(f.q)}">
        <label class="muted">Del</label><input type="date" id="fd" value="${f.desde}"><label class="muted">al</label><input type="date" id="fh" value="${f.hasta}">
        <select id="fc" ${esDisp(modalidad) ? 'hidden' : ''}><option value="">Todos los clientes</option>${clientes.map(c => `<option value="${c.id}" ${c.id === f.cliente ? 'selected' : ''}>${esc(c.nombre)}</option>`).join('')}</select>
        <select id="fe">${[['', 'Todas'], ['abierta', 'Abiertas'], ['cerrada', 'Cerradas'], ['cancelada', 'Canceladas']].map(([k, l]) => `<option value="${k}" ${k === f.estado ? 'selected' : ''}>${l}</option>`).join('')}</select>
        <span class="sp"></span><button class="btn" id="fx">Exportar</button>
        <button class="btn primary" id="nuevo">+ ${esDisp(modalidad) ? 'Nueva dispersión' : 'Nuevo depósito'}</button>
      </div>
      <div class="grid" id="kpi"></div>
      <div class="tablewrap tall"><table><thead><tr><th>Folio</th><th>Fecha</th><th>Cliente</th><th class="n">Depósito</th><th class="n">Comisión</th><th class="n">${esDisp(modalidad) ? 'A dispersar' : 'Retorno'}</th>
        <th>${esDisp(modalidad) ? 'Avance' : 'Retorno'}</th><th>Factura</th><th>Estado</th></tr></thead><tbody id="tb"></tbody></table></div>`;
    const q = {eq: {modalidad}, order: 'numero', desc: true};
    let ops = await db.list('operaciones', q);
    const disp = esDisp(modalidad) && ops.length ? await db.list('dispersiones', {in: {operacion_id: ops.map(o => o.id)}, select: 'operacion_id,estado,monto'}) : [];
    const draw = () => {
      Object.assign(f, {q: val('fq'), desde: val('fd'), hasta: val('fh'), cliente: val('fc'), estado: val('fe')}); listaOperaciones.f[modalidad] = {...f};
      const qq = Logica.normName(f.q);
      const rows = ops.filter(o => (!f.estado || o.estado === f.estado) && (!f.cliente || o.cliente_id === f.cliente) && (!f.desde || o.fecha_deposito >= f.desde) && (!f.hasta || o.fecha_deposito <= f.hasta)
        && (!qq || Logica.normName(o.folio + ' ' + (o.cliente_nombre || '') + ' ' + (o.razon_social || '')).includes(qq)));
      const sum = k => rows.reduce((s, o) => s + (+o[k] || 0), 0);
      $('#kpi').innerHTML = `<div class="card"><div class="k">Operaciones</div><div class="v">${rows.length}</div></div>
        <div class="card"><div class="k">Depositado</div><div class="v">${money(sum('monto_deposito'))}</div></div>
        <div class="card"><div class="k">Comisión</div><div class="v">${money(sum('monto_comision'))}</div></div>
        <div class="card"><div class="k">${esDisp(modalidad) ? 'A dispersar' : 'Retornos'}</div><div class="v">${money(sum('monto_retorno'))}</div></div>`;
      $('#tb').innerHTML = rows.map(o => {
        let av = '';
        if (esDisp(modalidad)) { const d = disp.filter(x => x.operacion_id === o.id); const ok = d.filter(x => x.estado === 'dispersado').length;
          av = d.length ? `<div style="display:flex;align-items:center;gap:6px"><div class="progress"><i style="width:${ok / d.length * 100}%"></i></div><span class="small">${ok}/${d.length}</span></div>` : tag('warn', 'sin lista'); }
        else av = o.retorno_entregado ? tag('ok', 'entregado') : tag('pendiente');
        return `<tr class="click" data-id="${o.id}"><td><b>${esc(o.folio)}</b></td><td>${fmtDay(o.fecha_deposito)}</td><td>${esc(o.cliente_nombre || '')}</td>
          <td class="n">${money(o.monto_deposito)}</td><td class="n">${money(o.monto_comision)}</td><td class="n">${money(o.monto_retorno)}</td><td>${av}</td>
          <td>${o.requiere_factura ? tag(o.factura_estado) : tag('no_aplica', 'no aplica')}</td><td>${tag(o.estado)}</td></tr>`;
      }).join('') || `<tr><td colspan="9" class="empty">No hay operaciones con estos filtros.</td></tr>`;
      $$('#tb tr[data-id]').forEach(tr => tr.onclick = () => go('operacion', {id: tr.dataset.id}));
      draw.rows = rows;
    };
    ['fq', 'fd', 'fh', 'fc', 'fe'].forEach(id => $('#' + id).oninput = draw);
    $('#nuevo').onclick = () => esDisp(modalidad) ? go('colorinesNueva') : nuevaOperacion(modalidad);
    $('#fx').onclick = () => {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, hoja(draw.rows, [{h: 'FOLIO', v: 'folio'}, {h: 'FECHA', v: o => fmtDay(o.fecha_deposito)}, {h: 'CLIENTE', v: 'cliente_nombre', w: 30},
        {h: 'DEPÓSITO', v: 'monto_deposito', money: 1, w: 14}, {h: '% COMISIÓN', v: 'pct_comision'}, {h: 'COMISIÓN', v: 'monto_comision', money: 1, w: 14},
        {h: esDisp(modalidad) ? 'A DISPERSAR' : 'RETORNO', v: 'monto_retorno', money: 1, w: 14}, {h: 'FACTURA', v: 'factura_estado'}, {h: 'ESTADO', v: 'estado'}]), 'Operaciones');
      guardarXlsx(wb, `${esDisp(modalidad) ? nombreModulo(modalidad) : 'Depositos'} ${today()}.xlsx`);
    };
    draw();
  };
}
listaOperaciones.f = {};
VIEWS.depositos = listaOperaciones('retorno');
VIEWS.colorines = (main, arg) => listaOperaciones(modulo)(main, arg);

/* ===================== alta / edición ===================== */
async function nuevaOperacion(modalidad, op) {
  const [clientes, empresas] = await Promise.all([cat('clientes'), cat('empresas')]);
  const editando = !!op; op = op || {modalidad, fecha_deposito: today(), fecha_solicitud: today(), fecha_compromiso: today(), tipo_retorno: 'TRANSFERENCIA'};
  let fichas = [];
  modal(editando ? `Editar ${op.folio}` : esDisp(modalidad) ? `Nueva dispersión ${nombreModulo(modalidad)}` : 'Nuevo depósito', `
    ${editando ? '' : '<p class="note info" style="margin-top:0">El folio FACOL se asigna automáticamente al guardar.</p>'}
    <div class="form2">
      <label ${esDisp(op.modalidad) ? 'hidden' : ''}>Cliente</label><select id="oCli" ${esDisp(op.modalidad) ? 'hidden' : ''}><option value="">— elige —</option>${clientes.filter(c => c.activo || c.id === op.cliente_id).map(c => `<option value="${c.id}" ${c.id === op.cliente_id ? 'selected' : ''}>${esc(c.nombre)}</option>`).join('')}</select>
      <input type="hidden" id="oMod" value="${esc(op.modalidad || modalidad)}">
      <label>Nombre fiscal</label><input id="oRaz" value="${esc(op.razon_social || '')}">
      <label>Fecha de depósito</label><input id="oFec" type="date" value="${op.fecha_deposito || ''}">
      <label>Monto depositado</label><input id="oMon" type="number" step="0.01" value="${op.monto_deposito ?? ''}" placeholder="${esDisp(op.modalidad) ? 'vacío = se calcula con la lista' : ''}">
      <label>% comisión</label><input id="oPct" type="number" step="0.0001" value="${op.pct_comision ?? ''}">
      <label>Banco ordenante</label><input id="oBO" value="${esc(op.banco_ordenante || '')}">
      <label>Cuenta ordenante</label><input id="oCO" value="${esc(op.cuenta_ordenante || '')}">
      <label>Empresa que factura</label><input id="oEF" list="empL" value="${esc(op.empresa_factura || '')}">
      <label>Banco receptor</label><input id="oBR" value="${esc(op.banco_receptor || '')}">
      <label>Cuenta receptora</label><input id="oCR" value="${esc(op.cuenta_receptora || '')}">
      <label>Solicitante</label><input id="oSol" value="${esc(op.solicitante || '')}">
      <label>Autorizado por</label><input id="oAut" value="${esc(op.autorizado_por || '')}">
      <label>Compromiso de pago</label><input id="oCom" type="date" value="${op.fecha_compromiso || ''}">
      <label>Tipo de retorno</label><select id="oTR">${['TRANSFERENCIA', 'EFECTIVO', 'TRANSFERENCIA/EFECTIVO'].map(t => `<option ${t === op.tipo_retorno ? 'selected' : ''}>${t}</option>`).join('')}</select>
      <label>Lleva factura</label><label><input type="checkbox" id="oFac" ${op.requiere_factura !== false ? 'checked' : ''}> Sí</label>
    </div>
    <datalist id="empL">${empresas.map(e => `<option value="${esc(e.nombre)}">`).join('')}</datalist>
    <div class="form" style="margin-top:8px"><label>Notas</label><textarea id="oNot" rows="2">${esc(op.notas || '')}</textarea></div>
    <div class="cols" style="margin-top:10px"><div id="oCalc" class="panel"></div>${editando ? '' : '<div><div id="oDrop"></div><div class="files" id="oFiles" style="margin-top:6px"></div></div>'}</div>
    <div class="err" id="oErr"></div>`, [
    ...(editando ? [{label: 'Borrar operación', cls: 'bad', onClick: () => { setTimeout(() => borrarOperacion(op), 50); }}] : []),
    {label: 'Cancelar'},
    {label: editando ? 'Guardar' : 'Guardar y asignar folio', cls: 'primary', onClick: async () => {
      const c = esDisp(op.modalidad) ? {id: op.cliente_id || null, nombre: op.cliente_nombre} : clientes.find(x => x.id === val('oCli'));
      if (!c) { $('#oErr').textContent = 'Elige el cliente'; return false; }
      const monto = num('oMon'), p = num('oPct');
      if (!(monto > 0) && !esDisp(val('oMod'))) { $('#oErr').textContent = 'Captura el monto depositado'; return false; }
      const {comision, retorno} = Logica.calcRetorno(monto, p);
      const row = {modalidad: val('oMod'), cliente_id: c.id, cliente_nombre: c.nombre, razon_social: val('oRaz'), fecha_deposito: val('oFec') || today(),
        monto_deposito: monto, pct_comision: p, monto_comision: comision, monto_retorno: retorno, banco_ordenante: val('oBO'), cuenta_ordenante: val('oCO'),
        empresa_factura: val('oEF'), banco_receptor: val('oBR'), cuenta_receptora: val('oCR'), solicitante: val('oSol'), autorizado_por: val('oAut'),
        fecha_compromiso: val('oCom') || null, tipo_retorno: val('oTR'), requiere_factura: val('oFac'), notas: val('oNot')};
      if (!row.requiere_factura) row.factura_estado = 'no_aplica'; else if (op.factura_estado === 'no_aplica' || !editando) row.factura_estado = 'pendiente';
      if (editando && esDisp(op.modalidad)) {
        // en Colorines la comisión sale de la lista; si cambia el %, se recalcula cada persona
        delete row.monto_comision; delete row.monto_retorno;
        await db.update('operaciones', op.id, row);
        if (+op.pct_comision !== p) {
          const disp = await db.list('dispersiones', {eq: {operacion_id: op.id}});
          for (const d of disp) { const g = Logica.grossUp(d.monto, p); await db.update('dispersiones', d.id, {comision: g.comision, total: g.total}); }
          const res = Logica.resumenDispersion(disp.map(d => ({...d, ...Logica.grossUp(d.monto, p)})), p);
          await db.update('operaciones', op.id, {monto_comision: res.comision});
          await recalcularComisiones(op.id, res.total, res.comision);
        }
        toast('Operación actualizada'); return go('operacion', {id: op.id});
      }
      if (editando) {
        const cambioMonto = +op.monto_deposito !== monto || +op.pct_comision !== p;
        await db.update('operaciones', op.id, row);
        if (cambioMonto) await recalcularComisiones(op.id, monto, comision);
        toast('Operación actualizada'); return go('operacion', {id: op.id});
      }
      const {folio, numero} = await db.rpc('siguiente_folio');
      const nueva = await db.insert('operaciones', {...row, folio, numero, fecha_solicitud: today()});
      await copiarEsquema(c.id, nueva.id, monto, comision);
      for (const f of fichas) await subirArchivo(f, {operacion_id: nueva.id, tipo: 'ficha_deposito'});
      toast(`Folio asignado: ${folio}`);
      go('operacion', {id: nueva.id, tab: esDisp(row.modalidad) ? 'dispersion' : 'resumen'});
    }}], {width: '920px'});
  const calc = () => {
    const m = num('oMon'), p = num('oPct'); const {comision, retorno} = Logica.calcRetorno(m, p);
    $('#oCalc').innerHTML = `<div class="kv"><div>Depósito</div><div><b>${money(m)}</b></div><div>Comisión ${pct(p)}</div><div>${money(comision)}</div>
      <div>${esDisp(val('oMod')) ? 'Total a dispersar' : 'Devolución al cliente'}</div><div><b>${money(retorno)}</b></div><div></div><div class="small muted">${esc(Logica.numLetras(retorno))}</div></div>`;
  };
  const fill = () => {
    const c = clientes.find(x => x.id === val('oCli')); if (!c || editando) return calc();
    const ef = empresas.find(e => e.id === c.empresa_factura_id), er = empresas.find(e => e.id === c.empresa_receptora_id);
    $('#oPct').value = c.pct_comision; $('#oRaz').value = c.razon_social || c.nombre; $('#oBO').value = c.banco_ordenante || ''; $('#oCO').value = c.cuenta_ordenante || '';
    $('#oEF').value = ef ? ef.nombre : ''; $('#oBR').value = er ? er.banco || '' : ''; $('#oCR').value = er ? er.cuenta || '' : '';
    $('#oSol').value = c.solicitante || ''; $('#oAut').value = c.autorizado_por || ''; $('#oFac').checked = c.requiere_factura !== false;
    calc();
  };
  $('#oCli').onchange = fill; ['oMon', 'oPct', 'oMod'].forEach(id => $('#' + id).oninput = calc); calc();
  if (!editando) {
    const drawF = () => $('#oFiles').innerHTML = fichas.map(f => `<div class="file"><span class="nm">📎 ${esc(f.name)}</span></div>`).join('');
    dropZone($('#oDrop'), fs => { fichas.push(...fs); drawF(); }, {accept: 'image/*,.pdf', text: 'Ficha de depósito: arrastra la imagen/PDF o haz clic'});
  }
}

/** Copia el esquema de comisiones del cliente a la operación (foto del momento). */
async function copiarEsquema(clienteId, opId, base, objetivo) {
  const lines = await db.list('cliente_comisiones', {eq: {cliente_id: clienteId}, order: 'orden'});
  const coms = await cat('comisionistas');
  const map = new Map(lines.map(l => [l.id, newId()]));
  const calc = Logica.ajustarCuadre(Logica.calcComisiones(base, lines), objetivo);
  const rows = calc.map(l => {
    const c = coms.find(x => x.id === l.comisionista_id);
    return {id: map.get(l.id), operacion_id: opId, comisionista_id: l.comisionista_id, nombre: c ? c.nombre : l.concepto, concepto: l.concepto, clase: l.clase,
      pct: l.pct, base_linea_id: l.base_linea_id ? map.get(l.base_linea_id) : null, monto: l.monto, tipo_pago: c ? c.tipo_pago : null, orden: l.orden};
  });
  const ord = [...rows.filter(r => !r.base_linea_id), ...rows.filter(r => r.base_linea_id)];
  if (ord.length) await db.insert('operacion_comisiones', ord);
}
async function recalcularComisiones(opId, base, objetivo) {
  const lines = await db.list('operacion_comisiones', {eq: {operacion_id: opId}});
  for (const l of Logica.ajustarCuadre(Logica.calcComisiones(base, lines), objetivo)) if (!l.pagado) await db.update('operacion_comisiones', l.id, {monto: l.monto});
}

/* ===================== detalle ===================== */
VIEWS.operacion = async (main, {id, tab}) => {
  const [op, lines, arch] = await Promise.all([db.get('operaciones', id), db.list('operacion_comisiones', {eq: {operacion_id: id}, order: 'orden'}), db.list('archivos', {eq: {operacion_id: id}, order: 'created_at'})]);
  if (!op) throw new Error('No existe la operación');
  if (!esConsulta() && op.modalidad !== modulo && (op.modalidad === 'retorno' || MODS.some(m => m.id === op.modalidad))) { modulo = op.modalidad; try { localStorage.setItem('facol-modulo', modulo); } catch { /* */ } }
  const disp = esDisp(op.modalidad) ? await db.list('dispersiones', {eq: {operacion_id: id}, order: 'orden'}) : [];
  const pend = disp.filter(d => d.estado !== 'dispersado').length;
  const nFac = arch.filter(a => a.tipo === 'factura' || a.tipo === 'complemento').length;
  const col = esDisp(op.modalidad); tab = tab || (col ? 'dispersion' : 'resumen');
  const tabs = col ? [['dispersion', `Dispersión${disp.length ? ` (${disp.length - pend}/${disp.length})` : ''}`], ['altas', 'Altas y contratos'], ['comisiones', 'Comisiones'], ['facturas', `Facturas (${nFac})`], ['archivos', `Archivos (${arch.length})`], ['resumen', 'Datos del movimiento']]
    : [['resumen', 'Resumen'], ['comisiones', 'Comisiones'], ['facturas', `Facturas (${nFac})`], ['archivos', `Archivos (${arch.length})`]];
  const cerrada = op.estado !== 'abierta';
  main.innerHTML = `<h2><button class="link back" id="bk">← ${esDisp(op.modalidad) ? nombreModulo(op.modalidad) : 'Depósitos'}</button></h2>
    <div class="toolbar"><h2 style="margin:0">${esc(op.folio)}</h2>${tag(op.estado)}${tag('', MODALIDAD[op.modalidad])}<span class="muted">${esc(op.cliente_nombre)} · ${fmtDay(op.fecha_deposito)}</span><span class="sp"></span>
      ${col ? '<button class="btn primary" id="bXls">Excel FACOL</button><button class="btn" id="bFicha">Formato de dispersión (PDF)</button>' : '<button class="btn" id="bFicha">Ficha de retorno (PDF)</button>'}${op.requiere_factura ? '<button class="btn" id="bFact">Solicitud de factura</button>' : ''}
      <button class="btn" id="bEdit">Editar</button><button class="btn" id="bFolio">Cambiar folio</button>
      ${op.estado === 'abierta' ? '<button class="btn ok" id="bClose">Cerrar operación</button>' : '<button class="btn" id="bOpen">Reabrir</button>'}</div>
    <div class="grid">
      <div class="card"><div class="k">Depósito</div><div class="v">${money(op.monto_deposito)}</div></div>
      <div class="card"><div class="k">Comisión ${pct(op.pct_comision)}</div><div class="v">${money(op.monto_comision)}</div></div>
      <div class="card"><div class="k">${esDisp(op.modalidad) ? 'Total a dispersar' : 'Devolución al cliente'}</div><div class="v">${money(op.monto_retorno)}</div></div>
      ${esDisp(op.modalidad) ? `<div class="card click" data-tab="dispersion"><div class="k">Pendientes de dispersar</div><div class="v ${pend ? 'warn' : 'ok'}">${disp.length ? pend : '—'}</div></div>`
        : `<div class="card"><div class="k">Retorno</div><div class="v ${op.retorno_entregado ? 'ok' : 'warn'}">${op.retorno_entregado ? 'Entregado' : 'Pendiente'}</div></div>`}
      <div class="card click" data-tab="facturas"><div class="k">Factura</div><div class="v ${op.factura_estado === 'emitida' || !op.requiere_factura ? 'ok' : 'warn'}" style="font-size:18px">${op.requiere_factura ? esc(op.factura_estado) : 'No aplica'}</div></div>
    </div>
    <div class="tabs">${tabs.map(([k, l]) => `<button data-tab="${k}" class="${k === tab ? 'on' : ''}">${l}</button>`).join('')}</div>
    <div id="tabBody"></div>`;
  $('#bk').onclick = () => go(esDisp(op.modalidad) ? 'colorines' : 'depositos');
  $$('[data-tab]').forEach(b => b.onclick = () => go('operacion', {id, tab: b.dataset.tab}));
  $('#bFicha').onclick = () => fichaRetorno(op, lines, disp);
  if ($('#bXls')) $('#bXls').onclick = () => guard(() => excelFacol(op));
  $('#bFolio').onclick = () => cambiarFolio(op);
  if ($('#bFact')) $('#bFact').onclick = () => solicitudFactura(op);
  if ($('#bEdit')) $('#bEdit').onclick = () => nuevaOperacion(op.modalidad, op);
  if ($('#bClose')) $('#bClose').onclick = () => cerrarOperacion(op, lines, disp, arch);
  if ($('#bOpen')) $('#bOpen').onclick = async () => { if (await confirmar('Reabrir', `¿Reabrir ${esc(op.folio)}?`)) { await db.update('operaciones', id, {estado: 'abierta', cerrada_at: null}); refresh(); } };
  const body = $('#tabBody');
  if (tab === 'resumen') tabResumen(body, op);
  if (tab === 'dispersion') await tabDispersion(body, op, disp, arch);
  if (tab === 'altas') await tabAltas(body, op, disp);
  if (tab === 'comisiones') tabComisiones(body, op, lines);
  if (tab === 'facturas') tabFacturas(body, op, arch);
  if (tab === 'archivos') tabArchivos(body, op, arch, disp);
};

function tabResumen(el, op) {
  const kv = [['Nombre fiscal del cliente', op.razon_social], ['Fecha de solicitud', fmtDay(op.fecha_solicitud)], ['Fecha de depósito', fmtDay(op.fecha_deposito)],
    ['Banco ordenante', op.banco_ordenante], ['Cuenta ordenante', op.cuenta_ordenante], ['Empresa que factura', op.empresa_factura], ['Banco receptor', op.banco_receptor],
    ['Cuenta receptora', op.cuenta_receptora], ['Solicitante', op.solicitante], ['Autorizado por', op.autorizado_por], ['Compromiso de pago', fmtDay(op.fecha_compromiso)],
    ['Tipo de retorno', op.tipo_retorno], ['Importe con letra', Logica.numLetras(op.monto_retorno)], ['Registró', op.created_by], ['Notas', op.notas]];
  el.innerHTML = `<div class="cols"><div class="panel"><div class="kv">${kv.map(([k, v]) => `<div>${k}</div><div>${esc(v || '—')}</div>`).join('')}</div></div>
    <div class="panel"><h3 style="margin-top:0">Seguimiento</h3>
      ${op.modalidad === 'retorno' ? `<p><label><input type="checkbox" id="rEnt" ${op.retorno_entregado ? 'checked' : ''} ${op.estado !== 'abierta' ? 'disabled' : ''}> Retorno entregado al cliente</label>
        <input type="date" id="rFec" value="${op.retorno_fecha || today()}" ${op.estado !== 'abierta' ? 'disabled' : ''}></p>` : '<p class="muted">El avance de la dispersión se lleva en la pestaña Dispersión.</p>'}
      ${op.requiere_factura ? `<p>Factura: <select id="rFac" ${op.estado !== 'abierta' ? 'disabled' : ''}>${['pendiente', 'solicitada', 'emitida'].map(s => `<option ${s === op.factura_estado ? 'selected' : ''}>${s}</option>`).join('')}</select>
        <span class="small muted">Sube el PDF/XML en Archivos.</span></p>` : '<p class="muted">Esta operación no lleva factura.</p>'}
      ${op.estado === 'abierta' ? '<button class="btn bad sm" id="rCan">Cancelar operación</button>' : ''}
      ${op.cerrada_at ? `<p class="small muted">Cerrada el ${fmtDate(op.cerrada_at)}</p>` : ''}</div></div>`;
  if ($('#rEnt')) $('#rEnt').onchange = () => guard(() => db.update('operaciones', op.id, {retorno_entregado: val('rEnt'), retorno_fecha: val('rEnt') ? val('rFec') : null}), 'Actualizado').then(refresh);
  if ($('#rFac')) $('#rFac').onchange = () => guard(() => db.update('operaciones', op.id, {factura_estado: val('rFac')}), 'Actualizado').then(refresh);
  if ($('#rCan')) $('#rCan').onclick = async () => { if (await confirmar('Cancelar operación', `¿Cancelar ${esc(op.folio)}? El folio no se reutiliza.`, 'Cancelar operación', 'bad')) { await db.update('operaciones', op.id, {estado: 'cancelada'}); refresh(); } };
}

function tabComisiones(el, op, lines) {
  const top = lines.filter(l => !l.base_linea_id);
  const tot = Logica.totalesComision(lines);
  const row = (l, sub) => `<tr class="${sub ? 'sub' : ''}"><td>${sub ? '↳ ' : ''}<b>${esc(l.nombre)}</b><div class="small muted">${esc(l.concepto || '')}</div></td><td>${esc(CLASES[l.clase] || l.clase)}</td>
    <td class="n">${pct(l.pct)}${sub ? ' de ' + esc((lines.find(x => x.id === l.base_linea_id) || {}).nombre || '') : ''}</td><td class="n">${money(l.monto)}</td>
    <td class="n">${sub ? '' : money(Logica.netoLinea(l, lines))}</td><td>${esc(l.tipo_pago || '')}</td>
    <td><label><input type="checkbox" class="pg" data-id="${l.id}" ${l.pagado ? 'checked' : ''}> ${l.pagado ? 'Pagada ' + fmtDay(l.fecha_pago) : 'Pendiente'}</label></td>
    <td>${l.pagado || (esDisp(op.modalidad) && l.clase === 'casa') ? '' : `<button class="link" data-edc="${l.id}">Editar</button>`}</td></tr>`;
  el.innerHTML = `<div class="tablewrap"><table><thead><tr><th>Promotor</th><th>Clase</th><th class="n">%</th><th class="n">Comisión</th><th class="n">Neto (tras terciarias)</th><th>Pago</th><th>Estado</th><th></th></tr></thead>
    <tbody>${top.map(l => row(l) + lines.filter(x => x.base_linea_id === l.id).map(x => row(x, true)).join('')).join('') || '<tr><td colspan="7" class="empty">El cliente no tiene esquema de comisiones. Configúralo en Clientes y presiona "Volver a copiar esquema".</td></tr>'}</tbody>
    <tfoot><tr><td colspan="2">Total comisión</td><td class="n">${pct(tot.pct)}</td><td class="n">${money(tot.total)}</td><td colspan="3">${Math.abs(tot.total - op.monto_comision) > 0.02 ? tag('bad', 'Cuadre: ' + money(op.monto_comision - tot.total)) : tag('ok', 'Cuadra con la comisión cobrada')}</td></tr></tfoot></table></div>
    <div class="toolbar" style="margin-top:10px"><button class="btn" id="cRe">Recalcular</button><button class="btn" id="cCp">${esDisp(op.modalidad) ? 'Aplicar esquema actual del módulo' : 'Volver a copiar esquema del cliente'}</button></div>`;
  $$('.pg').forEach(c => c.onchange = () => guard(() => db.update('operacion_comisiones', c.dataset.id, {pagado: c.checked, fecha_pago: c.checked ? today() : null}), c.checked ? 'Marcada como pagada' : 'Marcada como pendiente').then(refresh));
  $$('[data-edc]').forEach(b => b.onclick = () => editarComisionOp(op, lines, lines.find(l => l.id === b.dataset.edc)));
  $('#cRe').onclick = () => guard(() => recalcularComisiones(op.id, esDisp(op.modalidad) ? r2(+op.monto_retorno + +op.monto_comision) : op.monto_deposito, +op.monto_comision), 'Recalculado').then(refresh);
  $('#cCp').onclick = async () => {
    if (lines.some(l => l.pagado)) return toast('Hay comisiones pagadas; desmárcalas antes de reemplazar el esquema');
    if (!(await confirmar('Copiar esquema', 'Se reemplazan las comisiones de esta operación por el esquema actual del cliente.'))) return;
    await guard(async () => { for (const l of [...lines.filter(x => x.base_linea_id), ...lines.filter(x => !x.base_linea_id)]) await db.remove('operacion_comisiones', l.id); if (esDisp(op.modalidad)) await copiarEsquemaColorines(op.id, r2(+op.monto_retorno + +op.monto_comision), +op.monto_comision, op.modalidad);
      else await copiarEsquema(op.cliente_id, op.id, op.monto_deposito, +op.monto_comision); }, 'Esquema copiado'); refresh();
  };
}

/** Facturas emitidas (PDF/XML) y complementos de pago de la operación. Se piden para cerrarla. */
function tabFacturas(el, op, arch) {
  const lista = tipo => { const fs = arch.filter(a => a.tipo === tipo);
    return fs.map(a => `<div class="file"><span class="nm" title="${esc(a.nombre)}">📎 ${esc(a.nombre)}</span><span class="small muted">${fmtDate(a.created_at)}</span>
      <button class="link" data-ver="${a.id}">Ver</button>${op.estado === 'abierta' ? `<button class="link" data-del="${a.id}" style="color:var(--bad)">Quitar</button>` : ''}</div>`).join('') || '<div class="muted small">Ninguno todavía</div>'; };
  const nF = arch.filter(a => a.tipo === 'factura').length, nC = arch.filter(a => a.tipo === 'complemento').length;
  el.innerHTML = `${op.requiere_factura ? '' : '<div class="note">Esta operación está marcada como "no lleva factura".</div>'}
    <div class="cols">
      <div class="panel"><h3 style="margin-top:0">Facturas ${nF ? tag('ok', nF + ' subida(s)') : tag('warn', 'pendiente')}</h3>
        <p class="small muted" style="margin-top:0">Sube el PDF y el XML de cada factura emitida (puedes subir varios a la vez).</p>
        <div id="fDrop"></div><div class="files" style="margin-top:8px">${lista('factura')}</div></div>
      <div class="panel"><h3 style="margin-top:0">Complementos de pago ${nC ? tag('ok', nC + ' subido(s)') : tag('warn', 'pendiente')}</h3>
        <p class="small muted" style="margin-top:0">Sube el complemento de pago (PDF/XML) de cada factura cuando se reciba.</p>
        <div id="cDrop"></div><div class="files" style="margin-top:8px">${lista('complemento')}</div></div>
    </div>
    <p class="small muted">Todas las facturas y complementos también se pueden buscar en <button class="link" id="fAll">Facturas y complementos</button>.</p>`;
  const subir = tipo => async fs => {
    await guard(async () => { for (const f of fs) await subirArchivo(f, {operacion_id: op.id, tipo}); if (tipo === 'factura' && op.requiere_factura) await db.update('operaciones', op.id, {factura_estado: 'emitida'}); }, `${fs.length} archivo(s) subido(s)`);
    go('operacion', {id: op.id, tab: 'facturas'});
  };
  if (op.estado === 'abierta' || op.estado === 'cerrada') {
    dropZone($('#fDrop'), subir('factura'), {accept: '.pdf,.xml,image/*', text: 'Arrastra aquí las facturas o haz clic'});
    dropZone($('#cDrop'), subir('complemento'), {accept: '.pdf,.xml,image/*', text: 'Arrastra aquí los complementos de pago o haz clic'});
  }
  $('#fAll').onclick = () => go('facturas');
  $$('[data-ver]', el).forEach(b => b.onclick = () => { const a = arch.find(x => x.id === b.dataset.ver); guard(() => verArchivo(a.path, a.nombre)); });
  $$('[data-del]', el).forEach(b => b.onclick = async () => { const a = arch.find(x => x.id === b.dataset.del);
    if (!(await confirmar('Quitar archivo', `¿Quitar ${esc(a.nombre)}?`, 'Quitar', 'bad'))) return;
    await guard(async () => { await db.removeFile(a.path).catch(() => {}); await db.remove('archivos', a.id);
      if (a.tipo === 'factura' && !arch.some(x => x.tipo === 'factura' && x.id !== a.id)) await db.update('operaciones', op.id, {factura_estado: 'pendiente'}); }, 'Archivo quitado');
    go('operacion', {id: op.id, tab: 'facturas'}); });
}

function tabArchivos(el, op, arch, disp) {
  el.innerHTML = `<div class="cols"><div><div class="form" style="margin-bottom:8px"><label>Tipo de archivo</label><select id="aTipo">${Object.entries(TIPOS_ARCHIVO).map(([k, l]) => `<option value="${k}">${l}</option>`).join('')}</select></div>
    <div id="aDrop"></div><p class="small muted">Fichas de depósito, comprobantes de transferencia (puedes subir varios a la vez), facturas PDF/XML.</p></div>
    <div class="files">${Object.entries(TIPOS_ARCHIVO).map(([k, l]) => { const fs = arch.filter(a => a.tipo === k); return fs.length ? `<h3 style="margin:4px 0">${l} (${fs.length})</h3>` + fs.map(a => `<div class="file"><span class="nm" title="${esc(a.nombre)}">📎 ${esc(a.nombre)}</span>
      ${a.dispersion_id ? `<span class="tag">${esc((disp.find(d => d.id === a.dispersion_id) || {}).nombre || '')}</span>` : ''}<span class="small muted">${fmtDate(a.created_at)}</span>
      <button class="link" data-ver="${a.id}">Ver</button><button class="link" data-del="${a.id}" style="color:var(--bad)">Quitar</button></div>`).join('') : ''; }).join('') || '<div class="empty">Sin archivos</div>'}</div></div>`;
  dropZone($('#aDrop'), async fs => {
    const tipo = val('aTipo');
    await guard(async () => { for (const f of fs) await subirArchivo(f, {operacion_id: op.id, tipo}); if (tipo === 'factura' && op.requiere_factura) await db.update('operaciones', op.id, {factura_estado: 'emitida'}); }, `${fs.length} archivo(s) subido(s)`);
    go('operacion', {id: op.id, tab: 'archivos'});
  });
  $$('[data-ver]').forEach(b => b.onclick = () => { const a = arch.find(x => x.id === b.dataset.ver); guard(() => verArchivo(a.path, a.nombre)); });
  $$('[data-del]').forEach(b => b.onclick = async () => { const a = arch.find(x => x.id === b.dataset.del);
    if (!(await confirmar('Quitar archivo', `¿Quitar ${esc(a.nombre)}?`, 'Quitar', 'bad'))) return;
    await guard(async () => { await db.removeFile(a.path).catch(() => {}); await db.remove('archivos', a.id); }, 'Archivo quitado'); go('operacion', {id: op.id, tab: 'archivos'}); });
}

async function cerrarOperacion(op, lines, disp, arch) {
  const faltan = [];
  if (esDisp(op.modalidad)) {
    if (!disp.length) return toast('Primero carga la lista de dispersión');
    const p = disp.filter(d => d.estado !== 'dispersado');
    if (p.length) return modal('Aún hay pendientes', `<p>No se puede cerrar: ${p.length} persona(s) no han sido dispersadas.</p><ul>${p.slice(0, 15).map(d => `<li>${esc(d.nombre || d.nombre_archivo)} — ${money(d.monto)} ${tag(d.estado)}</li>`).join('')}</ul>`, [{label: 'Entendido', cls: 'primary'}]);
  } else if (!op.retorno_entregado) faltan.push('El retorno no está marcado como entregado');
  if (op.requiere_factura && !arch.some(a => a.tipo === 'factura')) faltan.push('No se ha subido la factura');
  if (op.requiere_factura && !arch.some(a => a.tipo === 'complemento')) faltan.push('No se ha subido el complemento de pago');
  if (!arch.some(a => a.tipo === 'ficha_deposito')) faltan.push('No hay ficha de depósito');
  const np = lines.filter(l => !l.pagado && l.monto > 0).length; if (np) faltan.push(`${np} comisión(es) sin marcar como pagadas`);
  const ok = await confirmar(`Cerrar ${op.folio}`, (faltan.length ? `<div class="note">${faltan.map(esc).join('<br>')}</div>¿Cerrar de todos modos?` : 'Todo está completo. ¿Marcar como cerrada?'), 'Cerrar operación', 'ok');
  if (!ok) return;
  await guard(() => db.update('operaciones', op.id, {estado: 'cerrada', cerrada_at: new Date().toISOString(), ...(op.modalidad === 'retorno' ? {} : {retorno_entregado: true})}), `${op.folio} cerrada`);
  refresh();
}

/** Cambia el %, sobre qué se calcula y la forma de pago de una comisión de ESTA operación. En Colorines la Casa absorbe la diferencia. */
function editarComisionOp(op, lines, l) {
  const tops = lines.filter(x => !x.base_linea_id && x.id !== l.id);
  const baseTotal = esDisp(op.modalidad) ? 'Comisión total' : 'Depósito';
  modal(`Comisión de ${l.nombre} · ${op.folio}`, `<div class="form">
      <label>%</label><input id="ecP" type="number" step="0.0001" value="${l.pct}">
      <label>Sobre</label><select id="ecB"><option value="">${baseTotal}</option>${tops.map(x => `<option value="${x.id}" ${x.id === l.base_linea_id ? 'selected' : ''}>${esc(x.nombre)} (sale de su parte)</option>`).join('')}</select>
      <label>Forma de pago</label><select id="ecF">${FORMAS_PAGO.map(f => `<option ${f === l.tipo_pago ? 'selected' : ''}>${f}</option>`).join('')}</select></div>
    <p class="small muted">Solo cambia esta operación. Para que aplique a las nuevas, cámbialo en ${esDisp(op.modalidad) ? '"Datos y comisiones"' : 'el cliente'}.${esDisp(op.modalidad) ? ' La Casa se ajusta para que el total siga siendo la comisión cobrada.' : ''}</p>`, [
    {label: 'Cancelar'},
    {label: 'Guardar', cls: 'primary', onClick: async () => {
      const base_linea_id = val('ecB') || null;
      if (lines.some(x => x.base_linea_id === l.id) && base_linea_id) { toast('De esta comisión salen otras; no puede depender de otra línea'); return false; }
      await db.update('operacion_comisiones', l.id, {pct: num('ecP'), base_linea_id, tipo_pago: val('ecF')});
      const col = esDisp(op.modalidad), base = col ? r2(+op.monto_retorno + +op.monto_comision) : +op.monto_deposito;
      if (col) { // la Casa = % total − lo que va sobre el total
        const ls = await db.list('operacion_comisiones', {eq: {operacion_id: op.id}}); const casa = ls.find(x => x.clase === 'casa');
        if (casa) await db.update('operacion_comisiones', casa.id, {pct: r2((+op.pct_comision - ls.filter(x => !x.base_linea_id && x.id !== casa.id).reduce((s, x) => s + +x.pct, 0)) * 10000) / 10000});
      }
      await recalcularComisiones(op.id, base, +op.monto_comision);
      toast('Comisión actualizada'); go('operacion', {id: op.id, tab: 'comisiones'});
    }}]);
}

/* ===================== borrar operación ===================== */
async function borrarOperacion(op) {
  const [disp, arch] = await Promise.all([db.list('dispersiones', {eq: {operacion_id: op.id}, select: 'id,estado'}), db.list('archivos', {eq: {operacion_id: op.id}})]);
  const nDisp = disp.filter(d => d.estado === 'dispersado').length;
  modal(`Borrar ${op.folio}`, `<p>Se borra la operación <b>${esc(op.folio)}</b> con todo lo que tiene:</p>
    <ul><li>${disp.length} colaborador(es) de la dispersión${nDisp ? ` — <b style="color:var(--bad)">${nDisp} ya marcados como dispersados</b>` : ''}</li><li>sus comisiones</li><li>${arch.length} archivo(s) (fichas, solicitud, comprobantes, facturas)</li></ul>
    <p class="note">Esto no se puede deshacer. Los beneficiarios y sus contratos <b>no</b> se borran.</p>
    <div class="form"><label>Escribe el folio para confirmar</label><input id="bfC" placeholder="${esc(op.folio)}"></div><div class="err" id="bfE"></div>`, [
    {label: 'Cancelar'},
    {label: 'Borrar definitivamente', cls: 'bad', onClick: async () => {
      if (val('bfC').toUpperCase() !== String(op.folio).toUpperCase()) { $('#bfE').textContent = 'El folio no coincide'; return false; }
      for (const a of arch) await db.removeFile(a.path).catch(() => {});
      await db.remove('operaciones', op.id);   // borra en cascada dispersiones, comisiones y registros de archivos
      if (db.mode === 'demo') for (const t of ['dispersiones', 'operacion_comisiones', 'archivos']) for (const r of await db.list(t, {eq: {operacion_id: op.id}})) await db.remove(t, r.id);
      toast(`${op.folio} borrada`); go(esDisp(op.modalidad) ? 'colorines' : 'depositos');
    }}]);
}

/* ===================== cambiar folio / consecutivo ===================== */
/** Cambia el folio de una operación y, si se pide, recorre el consecutivo de las posteriores de la misma modalidad. */
async function cambiarFolio(op) {
  const todas = await db.list('operaciones', {order: 'numero'});
  const posteriores = todas.filter(o => o.modalidad === op.modalidad && o.numero > op.numero && o.id !== op.id);
  modal(`Cambiar folio ${op.folio}`, `<div class="form"><label>Nuevo folio</label><input id="cfF" value="${esc(op.folio)}">
      <label></label><label><input type="checkbox" id="cfR" ${posteriores.length ? 'checked' : 'disabled'}> Recorrer el consecutivo de las ${posteriores.length} operación(es) posteriores</label></div>
    <p class="small muted">Puedes cambiar el número (FACOL2360) o el nombre completo (FACOL2354N). El siguiente folio que asigne la app se ajusta solo.</p>
    <div id="cfP"></div><div class="err" id="cfE"></div>`, [
    {label: 'Cancelar'},
    {label: 'Aplicar', cls: 'primary', onClick: async () => {
      const plan = calcular(); if (plan.error) { $('#cfE').textContent = plan.error; return false; }
      // dos pasos para no chocar con el folio único: primero temporales, luego los definitivos
      for (const p of plan.cambios) await db.update('operaciones', p.id, {folio: 'TMP-' + p.id});
      for (const p of plan.cambios) await db.update('operaciones', p.id, {folio: p.folio, numero: p.numero});
      const mismos = (await db.list('operaciones', {eq: {modalidad: op.modalidad}, select: 'numero'})).map(o => +o.numero || 0);
      const max = Math.max(0, ...mismos);
      if (op.modalidad === 'colorines' || op.modalidad === 'retorno') await setConfig('folio_siguiente', Math.max(max, ...(await db.list('operaciones', {select: 'numero,modalidad'})).filter(o => ['colorines', 'retorno'].includes(o.modalidad)).map(o => +o.numero || 0)) + 1);
      else { const lista = await getConfig('modulos', []); const m = lista.find(x => x.id === op.modalidad); if (m) { m.siguiente = max + 1; await setConfig('modulos', lista); } }
      toast(`Folio actualizado${plan.cambios.length > 1 ? ` (${plan.cambios.length} operaciones)` : ''}`); refresh();
    }}], {width: '640px'});
  const calcular = () => {
    const nuevo = val('cfF').toUpperCase().replace(/\s+/g, ''); const m = nuevo.match(/\d+/);
    if (!nuevo || !m) return {error: 'El folio debe llevar un número, por ejemplo FACOL2360'};
    const N = +m[0]; const recorrer = val('cfR');
    const cambios = [{id: op.id, de: op.folio, folio: nuevo, numero: N}];
    if (recorrer) posteriores.forEach((o, i) => cambios.push({id: o.id, de: o.folio, folio: o.folio.replace(/\d+/, String(N + i + 1)), numero: N + i + 1}));
    const ids = new Set(cambios.map(c => c.id)); const vistos = new Set();
    for (const c of cambios) {
      if (vistos.has(c.folio)) return {error: `El folio ${c.folio} se repetiría`}; vistos.add(c.folio);
      const choca = todas.find(o => !ids.has(o.id) && (o.folio === c.folio || (o.modalidad === op.modalidad && +o.numero === c.numero)));
      if (choca) return {error: `Ya existe ${choca.folio} (${MODALIDAD[choca.modalidad]}). Elige otro número.`, cambios};
    }
    return {cambios: cambios.filter(c => c.de !== c.folio || c.id === op.id), recorrer};
  };
  const prev = () => { const p = calcular(); $('#cfE').textContent = p.error || '';
    $('#cfP').innerHTML = p.cambios ? `<div class="tablewrap" style="max-height:240px"><table><tr><th>Antes</th><th>Después</th></tr>${p.cambios.map(c => `<tr><td>${esc(c.de)}</td><td><b>${esc(c.folio)}</b></td></tr>`).join('')}</table></div>` : ''; };
  $('#cfF').oninput = prev; $('#cfR').onchange = prev; prev();
}

/* ===================== documentos ===================== */
const DOC_CSS = `body{font-family:Arial,sans-serif;font-size:11px;color:#111;margin:0}h1{font-size:15px;margin:0 0 8px;text-align:center;background:#0f2f5c;color:#fff;padding:6px}
  h2{font-size:12px;margin:12px 0 4px;background:#e3edfb;padding:4px 6px}table{width:100%;border-collapse:collapse}td,th{border:1px solid #999;padding:3px 6px;text-align:left}th{background:#f0f0f0}
  td.n,th.n{text-align:right}.k{width:38%;background:#f7f7f7;font-weight:bold}.blk{page-break-inside:avoid;margin-bottom:6px}.firma{margin-top:40px;display:flex;justify-content:space-around;text-align:center}
  .firma div{border-top:1px solid #000;width:35%;padding-top:4px}.pb{page-break-before:always}.top{display:flex;justify-content:space-between;align-items:center}`;
async function fichaRetorno(op, lines, disp) {
  const coms = await cat('comisionistas'); const cta = l => { const c = coms.find(x => x.id === l.comisionista_id) || {}; return [c.cuenta, c.banco, c.clabe].filter(Boolean).join(' / '); };
  const kv =rows => `<table>${rows.map(([k, v]) => `<tr><td class="k">${k}</td><td>${esc(v ?? '')}</td></tr>`).join('')}</table>`;
  const top = lines.filter(l => !l.base_linea_id);
  const bloques = [...top, ...lines.filter(l => l.base_linea_id)].filter(l => l.monto > 0).map(l => `<div class="blk"><h2>${l.clase === 'casa' ? 'COMISIÓN CASA' : l.base_linea_id ? 'COMISIÓN DESGLOSADA DE ' + esc((lines.find(x => x.id === l.base_linea_id) || {}).nombre || '') : 'DATOS DEL COMISIONISTA O PROMOTOR'}</h2>${kv([
    ['FECHA DE ENTREGA', fmtDay(op.fecha_compromiso || op.fecha_deposito)], ['COMISIONISTA O PROMOTOR', l.nombre], ['TIPO DE PAGO', l.tipo_pago || ''],
    ['MONTO DE LA COMISIÓN A COBRAR', money(l.monto) + (l.base_linea_id ? ` (${pct(l.pct)})` : '')], ['CUENTA / BANCO / CLABE', cta(l)]])}</div>`).join('');
  const tabla = disp.length ? `<div class="pb"></div><h1>FORMATO DE DISPERSIÓN · ${esc(op.folio)}</h1><table><tr><th>#</th><th>ESTRATEGIA</th><th>BENEFICIARIO</th><th class="n">MONTO A DISPERSAR</th><th>BANCO</th><th>CUENTA</th><th>CLABE INTERBANCARIA</th></tr>
    ${disp.map((d, i) => `<tr><td>${i + 1}</td><td>${esc(d.estrategia)}</td><td>${esc(d.nombre || d.nombre_archivo)}</td><td class="n">${money(d.monto)}</td><td>${esc(d.banco || '')}</td><td>${esc(d.cuenta || '')}</td><td>${esc(d.clabe || '')}</td></tr>`).join('')}
    <tr><th colspan="3">TOTAL</th><th class="n">${money(disp.reduce((s, d) => s + +d.monto, 0))}</th><th colspan="3"></th></tr></table>` : '';
  const html = `<!doctype html><html><head><meta charset="utf-8"><style>${DOC_CSS}</style></head><body>
    <div class="top"><b>${esc(op.folio)}</b><span>${fmtDay(today())}</span></div><h1>FICHA DE RETORNO · DATOS DEL MOVIMIENTO</h1>
    ${kv([['FECHA DE SOLICITUD', fmtDay(op.fecha_solicitud)], ['NOMBRE FISCAL DEL CLIENTE', op.razon_social || op.cliente_nombre], ['BANCO ORDENANTE', op.banco_ordenante], ['CUENTA ORDENANTE', op.cuenta_ordenante],
      ['FECHA DE DEPÓSITO', fmtDay(op.fecha_deposito)], ['EMPRESA QUE FACTURA', op.empresa_factura], ['BANCO RECEPTOR', op.banco_receptor], ['CUENTA RECEPTORA', op.cuenta_receptora],
      ['SOLICITANTE', op.solicitante], ['FECHA COMPROMISO DE PAGO', fmtDay(op.fecha_compromiso)], ['AUTORIZADO POR', op.autorizado_por], ['MONTO DEPOSITADO', money(op.monto_deposito)],
      ['COMISIÓN', pct(op.pct_comision)], ['MONTO DE COMISIÓN', money(op.monto_comision)], [esDisp(op.modalidad) ? 'TOTAL A DISPERSAR' : 'DEVOLUCIÓN AL CLIENTE', money(op.monto_retorno)],
      ['IMPORTE CON LETRA', Logica.numLetras(op.monto_retorno)], ['TIPO DE RETORNO', op.tipo_retorno]])}
    ${bloques}
    <div class="firma"><div>ENTREGA</div><div>RECIBE</div></div>${tabla}</body></html>`;
  guardarPdf(html, `${op.folio} Ficha de retorno.pdf`);
}
function solicitudFactura(op) {
  modal('Solicitud de factura', `<div class="form"><label>Número de facturas</label><input id="sfN" type="number" min="1" value="1"><label>Concepto</label><input id="sfC" value="COMISIONES MERCANTILES">
    <label>Uso CFDI</label><input id="sfU" value="G03 Gastos en general"><label>Forma de pago</label><input id="sfF" value="03 TRANSFERENCIA ELECTRONICA DE FONDOS"><label>Método de pago</label><input id="sfM" value="PUE PAGO EN UNA SOLA EXHIBICION">
    <label>Observaciones</label><input id="sfO"></div>`, [{label: 'Cancelar'}, {label: 'Generar PDF', cls: 'primary', onClick: async () => {
    const n = Math.max(1, +val('sfN') || 1); const sub = r2(op.monto_deposito / 1.16 / n), iva = r2(sub * 0.16);
    const one = i => `<div class="${i ? 'pb' : ''}"><div class="top"><b>${esc(op.folio)}</b><span>Factura ${i + 1} de ${n}</span></div><h1>SOLICITUD DE FACTURA</h1>
      <table><tr><td class="k">EMPRESA QUE FACTURA</td><td>${esc(op.empresa_factura || '')}</td></tr><tr><td class="k">NOMBRE O RAZÓN SOCIAL CLIENTE</td><td>${esc(op.razon_social || op.cliente_nombre)}</td></tr>
      <tr><td class="k">FECHA</td><td>${fmtDay(op.fecha_deposito)}</td></tr><tr><td class="k">FORMA DE PAGO</td><td>${esc(val('sfF'))}</td></tr><tr><td class="k">MÉTODO DE PAGO</td><td>${esc(val('sfM'))}</td></tr>
      <tr><td class="k">USO CFDI</td><td>${esc(val('sfU'))}</td></tr><tr><td class="k">MONEDA</td><td>MXN-PESO MEXICANO</td></tr></table>
      <h2>CONCEPTOS</h2><table><tr><th>CANTIDAD</th><th>DESCRIPCIÓN</th><th class="n">VALOR UNITARIO</th><th class="n">IMPORTE</th></tr><tr><td>1</td><td>${esc(val('sfC'))}</td><td class="n">${money(sub)}</td><td class="n">${money(sub)}</td></tr>
      <tr><td colspan="3" class="n">SUBTOTAL</td><td class="n">${money(sub)}</td></tr><tr><td colspan="3" class="n">IVA 16%</td><td class="n">${money(iva)}</td></tr><tr><th colspan="3" class="n">TOTAL</th><th class="n">${money(sub + iva)}</th></tr></table>
      <p><b>OBSERVACIONES:</b> ${esc(val('sfO'))}</p></div>`;
    await guardarPdf(`<!doctype html><html><head><meta charset="utf-8"><style>${DOC_CSS}</style></head><body>${Array.from({length: n}, (_, i) => one(i)).join('')}</body></html>`, `${op.folio} Solicitud de factura.pdf`);
    if (op.factura_estado === 'pendiente') { await db.update('operaciones', op.id, {factura_estado: 'solicitada'}); refresh(); }
  }}]);
}

/* ===================== reporte de comisiones por rango de fechas ===================== */
VIEWS.comisiones = async (main, arg) => {
  const d = new Date(); const ini = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
  const f = Object.assign({desde: ini, hasta: today(), cliente: '', promotor: '', estado: ''}, VIEWS.comisiones.f || {}, arg || {});
  const [clientes, coms] = await Promise.all([cat('clientes'), cat('comisionistas')]);
  main.innerHTML = `<h2>Reporte de comisiones</h2>
    <div class="toolbar"><label class="muted">Del</label><input type="date" id="rd" value="${f.desde}"><label class="muted">al</label><input type="date" id="rh" value="${f.hasta}">
      <select id="rc" ${esDisp(modulo) ? 'hidden' : ''}><option value="">Todos los clientes</option>${clientes.map(c => `<option value="${c.id}" ${c.id === f.cliente ? 'selected' : ''}>${esc(c.nombre)}</option>`).join('')}</select>
      <select id="rp"></select>
      <select id="re">${[['', 'Pagadas y pendientes'], ['pend', 'Solo pendientes'], ['pag', 'Solo pagadas']].map(([k, l]) => `<option value="${k}" ${k === f.estado ? 'selected' : ''}>${l}</option>`).join('')}</select>
      <span class="sp"></span><button class="btn" id="rPay">Marcar seleccionadas pagadas</button><button class="btn primary" id="rX">Exportar Excel</button></div>
    <div class="grid" id="rk"></div><h3>Por promotor</h3><div class="tablewrap" id="rRes"></div><h3>Detalle por operación</h3><div class="tablewrap tall" id="rDet"></div>`;
  let ops = [], lines = [];
  const load = async () => {
    Object.assign(f, {desde: val('rd'), hasta: val('rh'), cliente: val('rc'), promotor: val('rp'), estado: val('re')}); VIEWS.comisiones.f = {...f};
    const q = {neq: {estado: 'cancelada'}, order: 'fecha_deposito'}; q.eq = {modalidad: modulo}; q.gte = f.desde ? {fecha_deposito: f.desde} : {}; q.lte = f.hasta ? {fecha_deposito: f.hasta} : {};
    if (f.cliente) q.eq.cliente_id = f.cliente;
    ops = await db.list('operaciones', q);
    lines = ops.length ? await db.list('operacion_comisiones', {in: {operacion_id: ops.map(o => o.id)}, order: 'orden'}) : [];
    opcionesPromotor(); draw();
  };
  // promotores del filtro: los del esquema Colorines, los dados de alta y los que aparezcan en las operaciones
  const esqNombres = esDisp(modulo) ? Logica.lineasColorines(await esquemaColorines(modulo)).map(l => l.nombre) : [];
  const opcionesPromotor = () => {
    const nombres = [...new Set([...esqNombres, ...coms.map(c => c.nombre), ...lines.map(l => l.nombre)].filter(Boolean))];
    if (f.promotor && !nombres.includes(f.promotor)) f.promotor = '';
    $('#rp').innerHTML = `<option value="">Todos los promotores</option>` + nombres.map(n => `<option ${n === f.promotor ? 'selected' : ''}>${esc(n)}</option>`).join('');
  };
  opcionesPromotor();
  const filt = () => lines.filter(l => (!f.promotor || l.nombre === f.promotor) && (f.estado !== 'pend' || !l.pagado) && (f.estado !== 'pag' || l.pagado) && l.monto > 0);
  const draw = () => {
    const ls = filt(); const opById = new Map(ops.map(o => [o.id, o]));
    const per = new Map();
    for (const l of ls) { const k = l.nombre; const p = per.get(k) || {nombre: l.nombre, n: 0, total: 0, pagado: 0, terciarias: 0}; p.n++; p.total += +l.monto; if (l.pagado) p.pagado += +l.monto;
      if (!l.base_linea_id) p.terciarias += lines.filter(x => x.base_linea_id === l.id).reduce((s, x) => s + +x.monto, 0); per.set(k, p); }
    const P = [...per.values()].sort((a, b) => b.total - a.total);
    const tot = ls.reduce((s, l) => s + +l.monto, 0), pag = ls.filter(l => l.pagado).reduce((s, l) => s + +l.monto, 0);
    $('#rk').innerHTML = `<div class="card"><div class="k">Operaciones</div><div class="v">${ops.length}</div></div><div class="card"><div class="k">Depositado</div><div class="v">${money(ops.reduce((s, o) => s + +o.monto_deposito, 0))}</div></div>
      <div class="card"><div class="k">Comisiones del periodo</div><div class="v">${money(tot)}</div></div><div class="card"><div class="k">Pagado</div><div class="v ok">${money(pag)}</div></div><div class="card"><div class="k">Por pagar</div><div class="v warn">${money(tot - pag)}</div></div>`;
    $('#rRes').innerHTML = `<table><thead><tr><th>Promotor</th><th class="n">Operaciones</th><th class="n">Comisión</th><th class="n">De ahí paga a terciarias</th><th class="n">Pagado</th><th class="n">Pendiente</th></tr></thead><tbody>
      ${P.map(p => `<tr><td><b>${esc(p.nombre)}</b></td><td class="n">${p.n}</td><td class="n">${money(p.total)}</td><td class="n">${p.terciarias ? money(p.terciarias) : '—'}</td><td class="n">${money(p.pagado)}</td><td class="n">${money(p.total - p.pagado)}</td></tr>`).join('') || '<tr><td colspan="6" class="empty">Sin comisiones en el periodo</td></tr>'}</tbody></table>`;
    $('#rDet').innerHTML = `<table><thead><tr><th><input type="checkbox" id="rAll"></th><th>Folio</th><th>Fecha</th><th>Cliente</th><th>Promotor</th><th>Clase</th><th class="n">%</th><th class="n">Base</th><th class="n">Comisión</th><th>Estado</th></tr></thead><tbody>
      ${ls.map(l => { const o = opById.get(l.operacion_id); const par = l.base_linea_id && lines.find(x => x.id === l.base_linea_id);
        return `<tr class="${par ? 'sub' : ''}"><td><input type="checkbox" class="rs" data-id="${l.id}" ${l.pagado ? 'disabled' : ''}></td><td><button class="link" data-op="${o.id}">${esc(o.folio)}</button></td><td>${fmtDay(o.fecha_deposito)}</td><td>${esc(o.cliente_nombre)}</td>
          <td>${par ? '↳ ' : ''}${esc(l.nombre)}</td><td>${esc(CLASES[l.clase] || '')}${par ? ' de ' + esc(par.nombre) : ''}</td><td class="n">${pct(l.pct)}</td><td class="n">${money(par ? par.monto : o.monto_deposito)}</td><td class="n">${money(l.monto)}</td>
          <td>${l.pagado ? tag('ok', 'pagada ' + fmtDay(l.fecha_pago)) : tag('pendiente')}</td></tr>`; }).join('') || '<tr><td colspan="10" class="empty">—</td></tr>'}</tbody></table>`;
    $('#rAll').onchange = e => $$('.rs:not(:disabled)').forEach(c => c.checked = e.target.checked);
    $$('[data-op]').forEach(b => b.onclick = () => go('operacion', {id: b.dataset.op, tab: 'comisiones'}));
    draw.ls = ls; draw.P = P;
  };
  ['rd', 'rh', 'rc'].forEach(id => $('#' + id).onchange = load); ['rp', 're'].forEach(id => $('#' + id).onchange = () => { f.promotor = val('rp'); f.estado = val('re'); VIEWS.comisiones.f = {...f}; draw(); });
  $('#rPay').onclick = async () => { const ids = $$('.rs:checked').map(c => c.dataset.id); if (!ids.length) return toast('Selecciona comisiones');
    if (!(await confirmar('Marcar pagadas', `¿Marcar ${ids.length} comisión(es) como pagadas hoy?`))) return;
    await guard(() => db.updateWhere('operacion_comisiones', {id: ids}, {pagado: true, fecha_pago: today()}), 'Comisiones marcadas como pagadas'); load(); };
  $('#rX').onclick = () => {
    const opById = new Map(ops.map(o => [o.id, o])); const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, hoja(draw.P, [{h: 'PROMOTOR', v: 'nombre', w: 26}, {h: 'OPERACIONES', v: 'n'}, {h: 'COMISIÓN', v: p => r2(p.total), money: 1, w: 14},
      {h: 'PAGA A TERCIARIAS', v: p => r2(p.terciarias), money: 1, w: 16}, {h: 'PAGADO', v: p => r2(p.pagado), money: 1, w: 14}, {h: 'PENDIENTE', v: p => r2(p.total - p.pagado), money: 1, w: 14}]), 'Por promotor');
    XLSX.utils.book_append_sheet(wb, hoja(draw.ls, [{h: 'FOLIO', v: l => opById.get(l.operacion_id).folio}, {h: 'FECHA', v: l => fmtDay(opById.get(l.operacion_id).fecha_deposito)},
      {h: 'CLIENTE', v: l => opById.get(l.operacion_id).cliente_nombre, w: 24}, {h: 'DEPÓSITO', v: l => +opById.get(l.operacion_id).monto_deposito, money: 1, w: 14},
      {h: 'PROMOTOR', v: 'nombre', w: 22}, {h: 'CONCEPTO', v: 'concepto', w: 30}, {h: 'CLASE', v: l => CLASES[l.clase] || l.clase},
      {h: 'SOBRE', v: l => l.base_linea_id ? (lines.find(x => x.id === l.base_linea_id) || {}).nombre : 'DEPÓSITO'}, {h: '%', v: l => +l.pct},
      {h: 'COMISIÓN', v: l => +l.monto, money: 1, w: 14}, {h: 'ESTADO', v: l => l.pagado ? 'PAGADA' : 'PENDIENTE'}, {h: 'FECHA PAGO', v: l => l.fecha_pago ? fmtDay(l.fecha_pago) : ''}]), 'Detalle');
    guardarXlsx(wb, `Comisiones ${f.desde} a ${f.hasta}.xlsx`);
  };
  load();
};
