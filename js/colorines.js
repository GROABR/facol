/* ---------- FACOL: Colorines — lista de personas, búsqueda en la BD, cálculo y control de la dispersión ---------- */
const datosBanco = b => b ? {beneficiario_id: b.id, nombre: b.nombre, banco: b.banco || Logica.bancoDeClabe(b.clabe), cuenta: b.cuenta || '', clabe: b.clabe || '', swift: b.swift || '', aba: b.aba || ''}
  : {beneficiario_id: null, nombre: null, banco: '', cuenta: '', clabe: '', swift: '', aba: ''};
function avisoCuenta(d) {
  if (!d.beneficiario_id) return tag('bad', 'no está en la BD');
  if (d.estrategia === 'EFECTIVO') return '';
  if (d.estrategia === 'USA') return d.swift || d.aba ? '' : tag('warn', 'sin SWIFT/ABA');
  if (!d.clabe) return d.cuenta ? tag('warn', 'sin CLABE') : tag('bad', 'sin cuenta');
  const v = Logica.validarClabe(d.clabe, d.banco); return !v.ok ? tag('bad', v.motivo) : v.aviso ? tag('warn', v.aviso) : '';
}

async function tabDispersion(el, op, disp) {
  if (!disp.length) return cargarListaEnOperacion(el, op);
  const pctC = +op.pct_comision; const res = Logica.resumenDispersion(disp, pctC, +op.monto_deposito);
  const abierta = op.estado === 'abierta';
  const comps = await db.list('archivos', {eq: {operacion_id: op.id, tipo: 'comprobante'}});
  el.innerHTML = `<div class="grid">
      <div class="card"><div class="k">Monto deseado (${res.n} personas)</div><div class="v">${money(res.monto)}</div></div>
      <div class="card"><div class="k">Comisión ${pct(pctC)}</div><div class="v">${money(res.comision)}</div></div>
      <div class="card"><div class="k">Total requerido</div><div class="v">${money(res.total)}</div></div>
      <div class="card"><div class="k">Diferencia vs depósito</div><div class="v ${Math.abs(res.diferencia) < 1 ? 'ok' : 'bad'}">${money(res.diferencia)}</div></div>
      <div class="card"><div class="k">Dispersado</div><div class="v ok">${money(res.dispersado)}</div></div>
      <div class="card"><div class="k">Pendiente (${res.nPendientes})</div><div class="v ${res.nPendientes ? 'warn' : 'ok'}">${money(res.pendiente)}</div></div></div>
    ${Math.abs(res.diferencia) >= 1 ? `<div class="note">El depósito (${money(op.monto_deposito)}) no coincide con lo que requiere la lista (${money(res.total)}). Revisa montos o el depósito.</div>` : ''}
    <div class="toolbar">
      <input id="dq" placeholder="Buscar persona" style="width:200px">
      <select id="df"><option value="">Todos</option><option value="pendiente">Pendientes</option><option value="dispersado">Dispersados</option><option value="rechazado">Rechazados</option><option value="alerta">Con alerta de cuenta</option></select>
      <span class="sp"></span>
      ${abierta ? `<button class="btn ok" id="dOk">Marcar seleccionados dispersados</button><button class="btn" id="dPend">Regresar a pendiente</button>
      <button class="btn" id="dAdd">+ Persona</button><button class="btn" id="dSync">Actualizar cuentas desde BD</button>` : ''}
      <button class="btn primary" id="dX">Excel FACOL (comisiones y dispersión)</button>
      ${abierta && !disp.some(d => d.estado === 'dispersado') ? '<button class="btn bad" id="dRep">Reemplazar lista</button>' : ''}</div>
    <div class="tablewrap tall"><table><thead><tr><th><input type="checkbox" id="dAll"></th><th>#</th><th>Beneficiario</th><th>Estrategia</th><th class="n">Monto</th><th class="n">Comisión</th><th class="n">Total</th>
      <th>Banco</th><th>Cuenta</th><th>CLABE</th><th>Estado</th><th></th></tr></thead><tbody id="dTb"></tbody>
      <tfoot><tr><td colspan="4">Total</td><td class="n">${money(res.monto)}</td><td class="n">${money(res.comision)}</td><td class="n">${money(res.total)}</td><td colspan="5">${Object.entries(res.porEstrategia).map(([k, v]) => `${k}: ${money(v)}`).join(' · ')}</td></tr></tfoot></table></div>`;
  const draw = () => {
    const q = Logica.normName(val('dq')), f = val('df');
    $('#dTb').innerHTML = disp.filter(d => (!q || Logica.normName((d.nombre || '') + ' ' + d.nombre_archivo).includes(q)) && (!f || (f === 'alerta' ? !!avisoCuenta(d) : d.estado === f))).map(d => {
      const nComp = comps.filter(c => c.dispersion_id === d.id).length;
      return `<tr data-id="${d.id}"><td>${abierta ? `<input type="checkbox" class="ds" value="${d.id}">` : ''}</td><td>${disp.indexOf(d) + 1}</td>
        <td><b>${esc(d.nombre || d.nombre_archivo)}</b>${d.nombre && Logica.normName(d.nombre) !== Logica.normName(d.nombre_archivo) ? `<div class="small muted">archivo: ${esc(d.nombre_archivo)}</div>` : ''}${d.puesto ? `<div class="small muted">${esc(d.puesto)}</div>` : ''} ${avisoCuenta(d)}</td>
        <td>${esc(d.estrategia)}</td><td class="n">${money(d.monto)}</td><td class="n">${money(d.comision)}</td><td class="n">${money(d.total)}</td>
        <td>${esc(d.banco || '')}</td><td class="mono">${esc(d.cuenta || '')}</td><td class="mono">${esc(d.clabe || '')}</td>
        <td>${tag(d.estado)}${d.fecha_dispersion ? `<div class="small muted">${fmtDay(d.fecha_dispersion)}</div>` : ''}${d.referencia ? `<div class="small muted">${esc(d.referencia)}</div>` : ''}</td>
        <td style="white-space:nowrap">${nComp ? `<button class="link" data-comp="${d.id}">📎${nComp}</button> ` : ''}${abierta ? `<button class="link" data-ed="${d.id}">Editar</button>` : ''}</td></tr>`;
    }).join('');
    $$('[data-ed]').forEach(b => b.onclick = () => editarLinea(op, disp.find(d => d.id === b.dataset.ed), disp));
    $$('[data-comp]').forEach(b => b.onclick = () => { const c = comps.filter(x => x.dispersion_id === b.dataset.comp); if (c.length === 1) return guard(() => verArchivo(c[0].path, c[0].nombre)); go('operacion', {id: op.id, tab: 'archivos'}); });
  };
  $('#dq').oninput = draw; $('#df').onchange = draw; draw();
  $('#dAll') && ($('#dAll').onchange = e => $$('.ds').forEach(c => c.checked = e.target.checked));
  const sel = () => $$('.ds:checked').map(c => c.value);
  if ($('#dOk')) $('#dOk').onclick = () => { const ids = sel(); if (!ids.length) return toast('Selecciona personas'); marcarDispersados(op, disp.filter(d => ids.includes(d.id))); };
  if ($('#dPend')) $('#dPend').onclick = async () => { const ids = sel(); if (!ids.length) return toast('Selecciona personas');
    await guard(() => db.updateWhere('dispersiones', {id: ids}, {estado: 'pendiente', fecha_dispersion: null}), 'Regresadas a pendiente'); refresh(); };
  if ($('#dAdd')) $('#dAdd').onclick = () => editarLinea(op, null, disp);
  if ($('#dSync')) $('#dSync').onclick = async () => {
    const bens = await cat('beneficiarios', true); let n = 0;
    for (const d of disp.filter(x => x.beneficiario_id && x.estado !== 'dispersado')) {
      const b = bens.find(x => x.id === d.beneficiario_id); if (!b) continue; const nd = datosBanco(b);
      if (nd.clabe !== d.clabe || nd.cuenta !== d.cuenta || nd.banco !== d.banco) { await db.update('dispersiones', d.id, nd); n++; }
    }
    toast(n ? `${n} cuenta(s) actualizadas` : 'Todas las cuentas ya coinciden con la BD'); refresh();
  };
  if ($('#dRep')) $('#dRep').onclick = async () => { if (!(await confirmar('Reemplazar lista', 'Se borra la lista actual para cargar otra.', 'Borrar lista', 'bad'))) return;
    await guard(async () => { for (const d of disp) await db.remove('dispersiones', d.id); }); refresh(); };
  $('#dX').onclick = () => guard(() => excelFacol(op));
}

function marcarDispersados(op, lineas) {
  let files = [];
  modal(`Marcar ${lineas.length} como dispersado(s)`, `<div class="form"><label>Fecha de dispersión</label><input type="date" id="mF" value="${today()}"><label>Referencia / rastreo</label><input id="mR" placeholder="opcional">
    <label>Comprobante(s)</label><div><div id="mD"></div><div class="files" id="mL" style="margin-top:6px"></div></div></div>
    <p class="small muted">${lineas.length === 1 ? 'El comprobante queda ligado a esta persona.' : 'Si subes un comprobante general (layout del banco), queda en Archivos de la operación.'}</p>`, [
    {label: 'Cancelar'}, {label: 'Confirmar', cls: 'ok', onClick: async () => {
      await db.updateWhere('dispersiones', {id: lineas.map(l => l.id)}, {estado: 'dispersado', fecha_dispersion: val('mF') || today(), ...(val('mR') ? {referencia: val('mR')} : {})});
      for (const f of files) await subirArchivo(f, {operacion_id: op.id, tipo: 'comprobante', dispersion_id: lineas.length === 1 ? lineas[0].id : null});
      const rest = (await db.list('dispersiones', {eq: {operacion_id: op.id}, select: 'estado'})).filter(d => d.estado !== 'dispersado').length;
      toast(rest ? `Listo. Faltan ${rest} por dispersar` : 'Todos dispersados: ya puedes cerrar la operación'); refresh();
    }}]);
  dropZone($('#mD'), fs => { files.push(...fs); $('#mL').innerHTML = files.map(f => `<div class="file"><span class="nm">📎 ${esc(f.name)}</span></div>`).join(''); }, {accept: 'image/*,.pdf', text: 'Arrastra o elige el comprobante'});
}

async function editarLinea(op, d, disp) {
  const bens = await cat('beneficiarios');
  const nuevo = !d; d = d || {estrategia: 'SCU', monto: 0, estado: 'pendiente', nombre_archivo: ''};
  let benId = d.beneficiario_id;
  modal(nuevo ? 'Agregar persona' : (d.nombre || d.nombre_archivo), `<div class="form">
    <label>Nombre (archivo)</label><input id="eNA" value="${esc(d.nombre_archivo || '')}">
    <label>Beneficiario en BD</label><div><input id="eB" placeholder="Buscar en la BD…" style="width:100%" value="${esc(d.nombre || '')}"><div id="eBL" class="small" style="max-height:150px;overflow:auto"></div></div>
    <label>Estrategia</label><select id="eE">${Logica.ESTRATEGIAS.map(e => `<option ${e === d.estrategia ? 'selected' : ''}>${e}</option>`).join('')}</select>
    <label>Monto a dispersar</label><input id="eM" type="number" step="0.01" value="${d.monto}">
    <label>Estado</label><select id="eS">${['pendiente', 'dispersado', 'rechazado'].map(s => `<option ${s === d.estado ? 'selected' : ''}>${s}</option>`).join('')}</select>
    <label>Notas</label><input id="eN" value="${esc(d.notas || '')}"></div><div id="eInfo" class="note info"></div>`, [
    ...(!nuevo ? [{label: 'Quitar de la lista', cls: 'bad', onClick: async () => { if (d.estado === 'dispersado') { toast('Ya fue dispersado'); return false; }
      if (!(await confirmar('Quitar', `¿Quitar a ${esc(d.nombre || d.nombre_archivo)} de la lista?`, 'Quitar', 'bad'))) return false; await db.remove('dispersiones', d.id); refresh(); }}] : []),
    {label: 'Cancelar'},
    {label: 'Guardar', cls: 'primary', onClick: async () => {
      const b = bens.find(x => x.id === benId); const g = Logica.grossUp(num('eM'), +op.pct_comision);
      const row = {...datosBanco(b), nombre_archivo: val('eNA') || (b && b.nombre) || '', estrategia: val('eE'), monto: g.monto, comision: g.comision, total: g.total, estado: val('eS'),
        fecha_dispersion: val('eS') === 'dispersado' ? (d.fecha_dispersion || today()) : null, notas: val('eN')};
      if (!row.nombre_archivo) { toast('Falta el nombre'); return false; }
      nuevo ? await db.insert('dispersiones', {...row, operacion_id: op.id, orden: disp.length}) : await db.update('dispersiones', d.id, row);
      toast('Guardado'); refresh();
    }}], {width: '680px'});
  const info = () => { const b = bens.find(x => x.id === benId); $('#eInfo').innerHTML = b ? `<b>${esc(b.nombre)}</b> · ${esc(b.banco || '')} · cuenta <span class="mono">${esc(b.cuenta || '—')}</span> · CLABE <span class="mono">${esc(b.clabe || '—')}</span> ${avisoCuenta({...datosBanco(b), estrategia: val('eE')})}` : 'Sin beneficiario: no hay datos bancarios. Búscalo o agrégalo en Beneficiarios.'; };
  const search = () => { const q = val('eB'); if (q.length < 3) { $('#eBL').innerHTML = ''; return; }
    const m = bens.map(b => ({b, s: Math.max(Logica.nameScore(q, b.nombre), Logica.normName(b.nombre).includes(Logica.normName(q)) ? 0.95 : 0)})).filter(x => x.s > 0.5).sort((a, b) => b.s - a.s).slice(0, 8);
    $('#eBL').innerHTML = m.map(x => `<div><button class="link" data-b="${x.b.id}">${esc(x.b.nombre)}</button> <span class="muted">${esc(x.b.banco || '')} ${esc(x.b.clabe || '')}</span></div>`).join('') || '<span class="muted">Sin coincidencias</span>';
    $$('#eBL [data-b]').forEach(a => a.onclick = () => { benId = a.dataset.b; $('#eB').value = bens.find(x => x.id === benId).nombre; $('#eBL').innerHTML = ''; info(); }); };
  $('#eB').oninput = search; $('#eE').onchange = info; info();
}

function buscarBeneficiario(nombre, bens) {
  return new Promise(res => {
    let elegido = null;
    modal('Buscar beneficiario', `<input id="bq" style="width:100%" value="${esc(nombre)}"><div id="bl" style="margin-top:8px;max-height:50vh;overflow:auto"></div>`, [{label: 'Cancelar', onClick: () => res(null)}]);
    dlg.onclose = () => res(elegido);
    const s = () => { const q = val('bq');
      const m = bens.map(b => ({b, s: Math.max(Logica.nameScore(q, b.nombre), Logica.normName(b.nombre).includes(Logica.normName(q)) ? 0.95 : 0)})).filter(x => x.s > 0.4).sort((a, b) => b.s - a.s).slice(0, 20);
      $('#bl').innerHTML = m.map(x => `<div class="file"><span class="nm"><b>${esc(x.b.nombre)}</b> <span class="muted small">${esc(x.b.banco || '')} ${esc(x.b.clabe || '')}</span></span><button class="btn sm" data-b="${x.b.id}">Elegir</button></div>`).join('') || '<div class="muted">Sin coincidencias</div>';
      $$('#bl [data-b]').forEach(b => b.onclick = () => { elegido = bens.find(x => x.id === b.dataset.b); res(elegido); dlg.close(); }); };
    $('#bq').oninput = s; s();
  });
}

