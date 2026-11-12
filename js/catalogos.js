/* ---------- FACOL: clientes (con su esquema de comisiones), promotores, empresas y ajustes ---------- */
const CLASES = {promotor: 'Promotor', terciaria: 'Terciaria', casa: 'Casa'};
const TIPOS_PAGO = ['EFECTIVO', 'TRANSFERENCIA', 'TRANSFERENCIA/EFECTIVO', 'CASA'];
const newId = () => crypto.randomUUID();

/* ===================== Clientes ===================== */
VIEWS.clientes = async main => {
  const [cli, lines, emp] = await Promise.all([cat('clientes', true), db.list('cliente_comisiones'), cat('empresas')]);
  const empN = id => (emp.find(e => e.id === id) || {}).nombre || '—';
  const col = esDisp(modulo);
  main.innerHTML = `<h2>${col ? 'Clientes' : 'Clientes y esquema de comisiones'}</h2>
    <div class="toolbar"><input id="q" placeholder="Buscar cliente" style="width:260px"><span class="sp"></span><button class="btn primary" id="nuevo">+ Nuevo cliente</button></div>
    <style>${col ? '.solo-ret{display:none}' : ''}</style>
    <div class="tablewrap"><table><thead><tr><th>Cliente</th><th>Razón social</th><th class="n solo-ret">% comisión</th><th class="solo-ret">Reparto</th><th>Factura</th><th>Empresa que factura</th></tr></thead><tbody id="tb"></tbody></table></div>`;
  const draw = () => {
    const q = Logica.normName(val('q'));
    const rows = cli.filter(c => !q || Logica.normName(c.nombre + ' ' + (c.razon_social || '')).includes(q));
    $('#tb').innerHTML = rows.map(c => {
      const ls = lines.filter(l => l.cliente_id === c.id); const tot = Logica.totalesComision(ls.map(l => ({...l, monto: +l.pct})));
      const ok = Math.abs(tot.pct - c.pct_comision) < 0.0001;
      return `<tr class="click" data-id="${c.id}"><td><b>${esc(c.nombre)}</b>${c.activo ? '' : ' ' + tag('cancelada', 'inactivo')}</td><td>${esc(c.razon_social || '')}</td><td class="n solo-ret">${pct(c.pct_comision)}</td>
        <td class="solo-ret">${ls.length} líneas ${ls.length ? (ok ? tag('ok', 'cuadra') : tag('bad', `suma ${pct(tot.pct)}`)) : tag('warn', 'sin reparto')}</td><td>${c.requiere_factura ? 'Sí' : 'No'}</td><td>${esc(empN(c.empresa_factura_id))}</td></tr>`;
    }).join('') || `<tr><td colspan="6" class="empty">Sin clientes. Crea el primero con "+ Nuevo cliente".</td></tr>`;
    $$('#tb tr[data-id]').forEach(tr => tr.onclick = () => editarCliente(cli.find(c => c.id === tr.dataset.id)));
  };
  $('#q').oninput = draw; $('#nuevo').onclick = () => editarCliente(null); draw();
};

async function editarCliente(c) {
  const [emp, coms, lineasDb] = await Promise.all([cat('empresas'), cat('comisionistas'), c ? db.list('cliente_comisiones', {eq: {cliente_id: c.id}, order: 'orden'}) : []]);
  c = c || {pct_comision: 11.5, requiere_factura: true, activo: true, autorizado_por: 'DIRECCIÓN'};
  let lineas = lineasDb.map(l => ({...l}));
  const empOpts = sel => `<option value="">—</option>` + emp.map(e => `<option value="${e.id}" ${e.id === sel ? 'selected' : ''}>${esc(e.nombre)}</option>`).join('');
  modal(c.id ? 'Cliente: ' + c.nombre : 'Nuevo cliente', `
    <div class="form2">
      <label>Nombre corto</label><input id="cNom" value="${esc(c.nombre || '')}">
      <label>Razón social</label><input id="cRaz" value="${esc(c.razon_social || '')}">
      <label>RFC</label><input id="cRfc" value="${esc(c.rfc || '')}">
      <label ${esDisp(modulo) ? 'hidden' : ''}>% comisión total</label><input id="cPct" type="number" step="0.0001" value="${c.pct_comision}" ${esDisp(modulo) ? 'hidden' : ''}>
      <label>Empresa que factura</label><select id="cEF">${empOpts(c.empresa_factura_id)}</select>
      <label>Empresa receptora</label><select id="cER">${empOpts(c.empresa_receptora_id)}</select>
      <label>Banco ordenante</label><input id="cBO" value="${esc(c.banco_ordenante || '')}">
      <label>Cuenta ordenante</label><input id="cCO" value="${esc(c.cuenta_ordenante || '')}">
      <label>Solicitante</label><input id="cSol" value="${esc(c.solicitante || '')}">
      <label>Autorizado por</label><input id="cAut" value="${esc(c.autorizado_por || '')}">
      <label>Lleva factura</label><label><input type="checkbox" id="cFac" ${c.requiere_factura ? 'checked' : ''}> Sí, se factura</label>
      <label>Activo</label><label><input type="checkbox" id="cAct" ${c.activo !== false ? 'checked' : ''}> Activo</label>
    </div>
    ${esDisp(modulo) ? '<p class="note info">En los módulos de dispersión las comisiones se configuran en <b>Datos y comisiones</b>.</p>' : ''}
    <div ${esDisp(modulo) ? 'hidden' : ''}>
    <h3>Reparto de la comisión (Retornos)</h3>
    <p class="small muted" style="margin-top:0">Las líneas "sobre: Depósito" deben sumar el % total. Una terciaria desglosada se calcula sobre otra línea (ej. ABEL 35% de CASA) y sale de esa línea.</p>
    <div class="tablewrap"><table><thead><tr><th>Concepto</th><th>Promotor</th><th>Clase</th><th class="n">%</th><th>Sobre</th><th class="n">Ej. $100,000</th><th></th></tr></thead><tbody id="lTb"></tbody>
    <tfoot><tr><td colspan="3">Suma sobre depósito</td><td class="n" id="lSum"></td><td colspan="3"></td></tr></tfoot></table></div>
    <div class="toolbar" style="margin-top:8px"><button class="btn sm" id="lAdd">+ Línea</button><span class="sp"></span><span class="small muted">¿Falta un promotor? Agrégalo en "Promotores".</span></div></div>
    <div class="err" id="cErr"></div>`, [
    ...(c.id ? [{label: 'Eliminar', cls: 'bad', onClick: async () => {
      const ops = await db.list('operaciones', {eq: {cliente_id: c.id}, select: 'id'});
      if (ops.length) { toast('Tiene operaciones; mejor márcalo como inactivo'); return false; }
      if (!(await confirmar('Eliminar cliente', `¿Eliminar ${esc(c.nombre)}?`, 'Eliminar', 'bad'))) return false;
      await db.remove('clientes', c.id); invalidate('clientes'); refresh();
    }}] : []),
    {label: 'Cancelar'},
    {label: 'Guardar', cls: 'primary', onClick: async () => {
      leer();
      const row = {nombre: val('cNom').toUpperCase(), razon_social: val('cRaz'), rfc: val('cRfc').toUpperCase(), pct_comision: num('cPct'),
        empresa_factura_id: val('cEF') || null, empresa_receptora_id: val('cER') || null, banco_ordenante: val('cBO'), cuenta_ordenante: val('cCO'),
        solicitante: val('cSol'), autorizado_por: val('cAut'), requiere_factura: val('cFac'), activo: val('cAct')};
      if (!row.nombre) { $('#cErr').textContent = 'Falta el nombre'; return false; }
      const tot = Logica.totalesComision(lineas.map(l => ({...l, monto: +l.pct})));
      if (!esDisp(modulo) && lineas.length && Math.abs(tot.pct - row.pct_comision) > 0.0001 && !(await confirmar('El reparto no cuadra', `Las líneas suman ${pct(tot.pct)} y la comisión es ${pct(row.pct_comision)}. ¿Guardar así?`))) return false;
      const saved = c.id ? await db.update('clientes', c.id, row) : await db.insert('clientes', row);
      for (const l of lineasDb) await db.remove('cliente_comisiones', l.id);
      const ordenadas = [...lineas.filter(l => !l.base_linea_id), ...lineas.filter(l => l.base_linea_id)];
      if (ordenadas.length) await db.insert('cliente_comisiones', ordenadas.map(l => ({id: l.id, cliente_id: saved.id, comisionista_id: l.comisionista_id || null,
        concepto: l.concepto, clase: l.clase, pct: +l.pct || 0, base_linea_id: l.base_linea_id || null, orden: lineas.indexOf(l)})));
      invalidate('clientes'); toast('Cliente guardado'); refresh();
    }},
  ], {width: '980px'});
  const comOpts = sel => `<option value="">—</option>` + coms.map(x => `<option value="${x.id}" ${x.id === sel ? 'selected' : ''}>${esc(x.nombre)}</option>`).join('');
  function leer() {
    $$('#lTb tr').forEach(tr => { const l = lineas.find(x => x.id === tr.dataset.id); if (!l) return;
      l.concepto = $('.lc', tr).value.trim(); l.comisionista_id = $('.lp', tr).value || null; l.clase = $('.lk', tr).value;
      l.pct = Logica.toNum($('.lpct', tr).value); l.base_linea_id = $('.lb', tr).value || null; });
  }
  function draw() {
    const calc = Logica.calcComisiones(100000, lineas);
    $('#lTb').innerHTML = lineas.map((l, i) => `<tr data-id="${l.id}">
      <td><input class="lc" value="${esc(l.concepto || '')}" style="width:150px"></td>
      <td><select class="lp" style="width:150px">${comOpts(l.comisionista_id)}</select></td>
      <td><select class="lk">${Object.entries(CLASES).map(([k, v]) => `<option value="${k}" ${k === l.clase ? 'selected' : ''}>${v}</option>`).join('')}</select></td>
      <td class="n"><input class="lpct" type="number" step="0.0001" value="${l.pct ?? ''}" style="width:80px;text-align:right"></td>
      <td><select class="lb"><option value="">Depósito</option>${lineas.filter(o => o.id !== l.id && !o.base_linea_id).map(o => `<option value="${o.id}" ${o.id === l.base_linea_id ? 'selected' : ''}>${esc(o.concepto || 'línea')}</option>`).join('')}</select></td>
      <td class="n">${money(calc[i].monto)}</td><td><button class="link lx" title="Quitar">✕</button></td></tr>`).join('') || '<tr><td colspan="7" class="muted">Sin líneas</td></tr>';
    $('#lSum').textContent = pct(Logica.totalesComision(calc).pct);
    $$('#lTb input, #lTb select').forEach(e => e.onchange = () => { leer(); draw(); });
    $$('#lTb .lp').forEach(s => s.addEventListener('change', () => { const tr = s.closest('tr'); const ci = $('.lc', tr); if (!ci.value && s.value) { leer(); lineas.find(x => x.id === tr.dataset.id).concepto = coms.find(x => x.id === s.value).nombre; draw(); } }));
    $$('#lTb .lx').forEach(b => b.onclick = () => { leer(); const id = b.closest('tr').dataset.id; lineas = lineas.filter(x => x.id !== id).map(x => x.base_linea_id === id ? {...x, base_linea_id: null} : x); draw(); });
  }
  $('#lAdd').onclick = () => { leer(); lineas.push({id: newId(), concepto: '', clase: 'promotor', pct: 0}); draw(); };
  draw();
}

/* ===================== Esquema fijo de comisiones Colorines ===================== */
async function esquemaColorines(mod = modulo) { return Logica.normalizarEsquema(await getConfig(cfgModulo(mod, 'esquema'), null)); }
/** Copia el esquema Colorines a una operación (foto del momento). */
async function copiarEsquemaColorines(opId, base, objetivo, mod = modulo) {
  const e = await esquemaColorines(mod); const coms = await cat('comisionistas', true);
  // quien esté en el esquema y no exista en Promotores se da de alta solo
  for (const l of Logica.lineasColorines(e)) if (l.nombre && !coms.some(c => c.nombre === l.nombre))
    coms.push(await db.insert('comisionistas', {nombre: l.nombre, tipo_pago: l.tipo_pago || (l.clase === 'casa' ? 'CASA' : 'EFECTIVO')}));
  invalidate('comisionistas');
  const lines = Logica.lineasColorines(e).map(l => ({...l, comisionista_id: (coms.find(c => c.nombre === l.nombre) || {}).id || l.comisionista_id || null}));
  const map = new Map(lines.map(l => [l.id, newId()]));
  const calc = Logica.ajustarCuadre(Logica.calcComisiones(base, lines), objetivo);
  const rows = calc.map(l => { const c = coms.find(x => x.id === l.comisionista_id);
    return {id: map.get(l.id), operacion_id: opId, comisionista_id: l.comisionista_id, nombre: l.nombre, concepto: l.concepto, clase: l.clase, pct: l.pct,
      base_linea_id: l.base_linea_id ? map.get(l.base_linea_id) : null, monto: l.monto, tipo_pago: l.tipo_pago || (c ? c.tipo_pago : null), orden: l.orden}; });
  await db.insert('operacion_comisiones', [...rows.filter(r => !r.base_linea_id), ...rows.filter(r => r.base_linea_id)]);
}

/* Colorines trabaja con un solo cliente y una sola empresa: sus datos fijos van aquí. */
const DATOS_COLORINES = {cliente: 'STAR SEEDS', razon_social: 'STAR SEEDS PROYECTOS DE SURESTE S.A.P.I. DE C.V.', rfc: '', empresa_factura: 'BUILD FORCE',
  banco_receptor: 'BANORTE', cuenta_receptora: '****67295', banco_ordenante: 'BBVA', cuenta_ordenante: '', solicitante: 'ABEL ORTIZ', autorizado_por: 'DIRECCIÓN', requiere_factura: true,
  concepto_factura: 'COMISIONES MERCANTILES', uso_cfdi: 'G03 Gastos en general', forma_pago_sat: '03 Transferencia electrónica de fondos', metodo_pago: 'PUE Pago en una sola exhibición'};
const DATOS_VACIOS = {cliente: '', razon_social: '', rfc: '', empresa_factura: '', banco_receptor: '', cuenta_receptora: '', banco_ordenante: '', cuenta_ordenante: '',
  solicitante: '', autorizado_por: 'DIRECCIÓN', requiere_factura: true, concepto_factura: 'COMISIONES MERCANTILES', uso_cfdi: 'G03 Gastos en general',
  forma_pago_sat: '03 Transferencia electrónica de fondos', metodo_pago: 'PUE Pago en una sola exhibición'};
async function datosColorines(mod = modulo) { const d = await getConfig(cfgModulo(mod, 'datos'), null); return {...(mod === 'colorines' ? DATOS_COLORINES : DATOS_VACIOS), ...(d || {})}; }
const CAMPOS_DATOS = [['cliente', 'Cliente (nombre corto)'], ['razon_social', 'Nombre fiscal del cliente'], ['rfc', 'RFC del cliente'], ['empresa_factura', 'Empresa que factura'],
  ['banco_receptor', 'Banco receptor'], ['cuenta_receptora', 'Cuenta receptora'], ['banco_ordenante', 'Banco ordenante'], ['cuenta_ordenante', 'Cuenta ordenante'],
  ['solicitante', 'Solicitante'], ['autorizado_por', 'Autorizado por']];   // forma de pago, uso CFDI, etc. salen tal cual del formato de factura estandarizado
const FORMAS_PAGO = ['EFECTIVO', 'TRANSFERENCIA', 'TRANSFERENCIA/EFECTIVO', 'CASA'];

VIEWS.esquemaColorines = async main => {
  const [e0, dc, coms0] = await Promise.all([esquemaColorines(modulo), datosColorines(modulo), cat('comisionistas', true)]);
  const e = JSON.parse(JSON.stringify(e0));
  // datos de pago de cada uno (vienen de Promotores)
  const pago = n => { const c = coms0.find(x => x.nombre === n) || {}; return {banco: c.banco || '', cuenta: c.cuenta || '', clabe: c.clabe || '', tipo_pago: c.tipo_pago || ''}; };
  e.casa = {...pago(e.casa.nombre || 'CASA'), ...e.casa, ...{tipo_pago: e.casa.tipo_pago || pago(e.casa.nombre || 'CASA').tipo_pago || 'CASA'}};
  e.lineas = e.lineas.map(l => ({...pago(l.nombre), ...l, tipo_pago: l.tipo_pago || pago(l.nombre).tipo_pago || 'EFECTIVO'}));
  const lmods = modulo === 'colorines' ? [] : await getConfig('modulos', []); const mm = lmods.find(x => x.id === modulo);
  main.innerHTML = `<h2>Datos y comisiones · ${esc(nombreModulo(modulo))}</h2>
    ${mm ? `<div class="panel" style="margin-bottom:14px"><h3 style="margin-top:0">Módulo</h3><div class="form2">
      <label>Nombre</label><input id="mN" value="${esc(mm.nombre)}"><label>Prefijo del folio</label><input id="mP" value="${esc(mm.prefijo)}">
      <label>Siguiente número</label><input id="mS" type="number" value="${esc(mm.siguiente)}"><span></span><span class="small muted">Próximo folio: <b id="mPrev"></b></span></div>
      <div class="toolbar" style="margin:10px 0 0"><button class="btn bad" id="mDel">Eliminar cliente</button><span class="sp"></span><button class="btn primary" id="mSave">Guardar módulo</button></div></div>` : ''}
    <div class="panel" style="margin-bottom:14px"><h3 style="margin-top:0">Cliente, empresa y factura</h3>
      <p class="small muted" style="margin-top:0">Se usan en todas las dispersiones: aparecen en el Excel FACOL, el formato de dispersión y la solicitud de factura.</p>
      <div class="form2">${CAMPOS_DATOS.map(([k, l]) => `<label>${l}</label><input id="d_${k}" value="${esc(dc[k] || '')}">`).join('')}
        <label>Lleva factura</label><label><input type="checkbox" id="d_requiere_factura" ${dc.requiere_factura !== false ? 'checked' : ''}> Sí</label></div>
      <div class="toolbar" style="margin:10px 0 0"><span class="sp"></span><button class="btn primary" id="dSave">Guardar datos</button></div></div>
    <h3>Comisiones y forma de pago</h3>
    <p class="muted" style="margin-top:0">Cada promotor tiene su % y <b>sobre qué</b> se calcula: sobre la comisión total, sobre la Casa o sobre otro promotor (sale de la parte de ese). La Casa se queda con lo que resta. Cada operación guarda una copia; los cambios aplican a las nuevas, y en cada operación puedes ajustar sus comisiones.</p>
    <div class="panel"><div class="form" style="max-width:420px"><label>% total que se cobra</label><input id="eP" type="number" step="0.0001" value="${e.pct}"></div>
      <div class="tablewrap" style="margin-top:10px"><table><thead><tr><th>Promotor</th><th class="n">%</th><th>Sobre</th><th>Forma de pago</th><th>Banco</th><th>Cuenta</th><th>CLABE</th><th></th></tr></thead><tbody id="eT"></tbody></table></div>
      <button class="btn sm" id="eAdd" style="margin-top:6px">+ Promotor</button></div>
    <div class="panel" style="margin-top:14px"><h3 style="margin-top:0">Ejemplo con una dispersión de <input id="eEj" type="number" value="100000" style="width:120px"></h3><div id="eV"></div></div>
    <div class="toolbar" style="margin-top:12px"><span class="sp"></span><button class="btn primary" id="eSave">Guardar esquema</button></div>`;
  if (mm) {
    const pv = () => $('#mPrev').textContent = val('mP').toUpperCase() + val('mS'); $('#mP').oninput = $('#mS').oninput = pv; pv();
    $('#mSave').onclick = () => guard(async () => {
      const p = val('mP').toUpperCase().replace(/[^A-Z0-9]/g, ''); if (!val('mN') || !p) throw new Error('Falta el nombre o el prefijo');
      if (lmods.some(x => x.id !== mm.id && x.prefijo === p)) throw new Error('Ese prefijo ya lo usa otro cliente');
      Object.assign(mm, {nombre: val('mN').toUpperCase(), prefijo: p, siguiente: Math.max(1, Math.floor(num('mS')) || 1)}); await setConfig('modulos', lmods); await cargarModulos(); renderNav();
    }, 'Módulo guardado');
    $('#mDel').onclick = async () => {
      const ops = await db.list('operaciones', {eq: {modalidad: mm.id}, select: 'id'});
      if (ops.length) return toast(`No se puede: tiene ${ops.length} operación(es). Bórralas primero.`);
      if (!(await confirmar('Eliminar cliente', `¿Eliminar el módulo ${esc(mm.nombre)}? Sus datos y comisiones se borran.`, 'Eliminar', 'bad'))) return;
      await guard(async () => { await setConfig('modulos', lmods.filter(x => x.id !== mm.id)); await cargarModulos(); }, 'Cliente eliminado'); setModulo('colorines');
    };
  }
  $('#dSave').onclick = () => guard(() => setConfig(cfgModulo(modulo, 'datos'), {...Object.fromEntries(CAMPOS_DATOS.map(([k]) => [k, ['cliente', 'razon_social', 'rfc'].includes(k) ? val('d_' + k).toUpperCase() : val('d_' + k)])), requiere_factura: val('d_requiere_factura')}), 'Datos guardados');
  const fila = (l, i) => {
    const casa = i < 0; const top = e.lineas.filter((x, j) => j !== i && (!x.base || x.base === 'total') && x.nombre);
    const sobre = casa ? '<span class="muted small">lo que resta</span>' : `<select class="tb">${[['total', 'Comisión total'], ['casa', 'Casa'], ...top.map(x => [x.nombre, x.nombre])].map(([k, t]) => `<option value="${esc(k)}" ${k === (l.base || 'total') ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select>`;
    return `<tr data-i="${i}"><td>${casa ? `<b>${esc(l.nombre || 'CASA')}</b>` : `<input class="tn" value="${esc(l.nombre)}" style="width:150px">`}</td>
      <td class="n">${casa ? `<b id="eCasa"></b>` : `<input class="tp" type="number" step="0.0001" value="${l.pct}" style="width:78px;text-align:right">`}</td><td>${sobre}</td>
      <td><select class="tf">${FORMAS_PAGO.map(f => `<option ${f === l.tipo_pago ? 'selected' : ''}>${f}</option>`).join('')}</select></td>
      <td><input class="tbk" value="${esc(l.banco || '')}" style="width:100px"></td><td><input class="tct" value="${esc(l.cuenta || '')}" style="width:110px"></td><td><input class="tcl" value="${esc(l.clabe || '')}" style="width:165px"></td>
      <td>${casa ? '' : '<button class="link tx" style="color:var(--bad)">✕</button>'}</td></tr>`;
  };
  const leer = () => {
    e.pct = num('eP');
    $$('#eT tr').forEach(tr => { const i = +tr.dataset.i; const o = i < 0 ? e.casa : e.lineas[i];
      if (i >= 0) { o.nombre = $('.tn', tr).value.trim().toUpperCase(); o.pct = Logica.toNum($('.tp', tr).value); o.base = $('.tb', tr).value; }
      o.tipo_pago = $('.tf', tr).value; o.banco = $('.tbk', tr).value.trim().toUpperCase(); o.cuenta = $('.tct', tr).value.replace(/\s/g, ''); o.clabe = $('.tcl', tr).value.replace(/\s/g, ''); });
  };
  const ver = () => {
    leer(); const ls = Logica.calcComisiones(num('eEj'), Logica.lineasColorines(e)); const casa = ls.find(l => l.clase === 'casa');
    if ($('#eCasa')) $('#eCasa').textContent = pct(casa.pct);
    const hijos = p => ls.filter(x => x.base_linea_id === p.id).map(x => `<tr class="sub"><td>↳ ${esc(x.nombre)} <span class="small muted">(${pct(x.pct)} de ${esc(p.nombre)}, sale de su parte)</span></td><td>${esc(x.tipo_pago || '')}</td><td class="n">${money(x.monto)}</td></tr>`).join('');
    $('#eV').innerHTML = `<table><thead><tr><th>Concepto</th><th>Forma de pago</th><th class="n">Monto</th></tr></thead><tbody>
      ${ls.filter(l => !l.base_linea_id).map(l => `<tr><td><b>${esc(l.nombre)}</b> <span class="small muted">${pct(l.pct)} ${l.clase === 'casa' ? '' : 'del total'}</span></td><td>${esc(l.tipo_pago || '')}</td><td class="n">${money(l.monto)}</td></tr>${hijos(l)}
        ${ls.some(x => x.base_linea_id === l.id) ? `<tr class="sub"><td>↳ ${esc(l.nombre)} neto</td><td></td><td class="n">${money(Logica.netoLinea(l, ls))}</td></tr>` : ''}`).join('')}
      </tbody><tfoot><tr><td>Total comisión (${pct(Logica.totalesComision(ls).pct)})</td><td></td><td class="n">${money(Logica.totalesComision(ls).total)}</td></tr></tfoot></table>
      ${casa.pct < 0 ? '<div class="err">Los promotores "sobre el total" suman más que el % total</div>' : ''}
      ${e.lineas.some(l => l.clabe && !Logica.validarClabe(l.clabe).ok) || (e.casa.clabe && !Logica.validarClabe(e.casa.clabe).ok) ? '<div class="err">Hay una CLABE con error</div>' : ''}`;
  };
  const drawT = () => { $('#eT').innerHTML = fila(e.casa, -1) + e.lineas.map((l, i) => fila(l, i)).join('');
    $$('#eT input, #eT select').forEach(x => { x.oninput = ver; x.onchange = () => { leer(); drawT(); }; });
    $$('#eT .tx').forEach(b => b.onclick = () => { leer(); const i = +b.closest('tr').dataset.i; const n = e.lineas[i].nombre;
      e.lineas.splice(i, 1); e.lineas.forEach(x => { if (x.base === n) x.base = 'total'; }); drawT(); }); ver(); };
  $('#eAdd').onclick = () => { leer(); e.lineas.push({nombre: '', pct: 0, base: 'total', tipo_pago: 'EFECTIVO'}); drawT(); };
  ['eP', 'eEj'].forEach(id => $('#' + id).oninput = ver); drawT();
  $('#eSave').onclick = () => guard(async () => {
    leer(); e.lineas = e.lineas.filter(l => l.nombre);
    if (Logica.lineasColorines(e)[0].pct < 0) throw new Error('Los promotores "sobre el total" suman más que el % total');
    const nombres = e.lineas.map(l => l.nombre); if (new Set(nombres).size !== nombres.length) throw new Error('Hay promotores repetidos');
    for (const x of [e.casa, ...e.lineas]) if (x.clabe) { const v = Logica.validarClabe(x.clabe); if (!v.ok) throw new Error(`${x.nombre || 'CASA'}: ${v.motivo}`); x.clabe = Logica.normalizarClabe(x.clabe); }
    // los datos de pago se guardan también en Promotores (los usan el Excel y el formato de dispersión)
    const coms = await cat('comisionistas', true);
    for (const x of [{...e.casa, nombre: e.casa.nombre || 'CASA'}, ...e.lineas]) {
      const datos = {tipo_pago: x.tipo_pago, banco: x.banco || null, cuenta: x.cuenta || null, clabe: x.clabe || null};
      const c = coms.find(y => y.nombre === x.nombre);
      x.id = c ? (await db.update('comisionistas', c.id, datos), c.id) : (await db.insert('comisionistas', {nombre: x.nombre, ...datos})).id;
    }
    const limpio = {pct: e.pct, casa: {nombre: e.casa.nombre || 'CASA', id: e.casa.id, tipo_pago: e.casa.tipo_pago},
      lineas: e.lineas.map(l => ({nombre: l.nombre, pct: l.pct, base: l.base || 'total', tipo_pago: l.tipo_pago, id: l.id}))};
    await setConfig(cfgModulo(modulo, 'esquema'), limpio); invalidate('comisionistas');
  }, 'Esquema guardado');
};

/* ===================== Promotores / comisionistas ===================== */
VIEWS.comisionistas = async main => {
  const rows = await cat('comisionistas', true);
  main.innerHTML = `<h2>Promotores y comisionistas</h2>
    <div class="toolbar"><span class="muted">Personas que cobran comisión por operación (promotores, terciarias, casa).</span><span class="sp"></span><button class="btn primary" id="nuevo">+ Nuevo</button></div>
    <div class="tablewrap"><table><thead><tr><th>Nombre</th><th>Tipo de pago</th><th>Banco</th><th>Cuenta</th><th>CLABE</th></tr></thead><tbody>
    ${rows.map(r => `<tr class="click" data-id="${r.id}"><td><b>${esc(r.nombre)}</b>${r.activo ? '' : ' ' + tag('cancelada', 'inactivo')}</td><td>${esc(r.tipo_pago || '')}</td><td>${esc(r.banco || '')}</td><td class="mono">${esc(r.cuenta || '')}</td><td class="mono">${esc(r.clabe || '')}</td></tr>`).join('') || '<tr><td colspan="5" class="empty">Sin promotores</td></tr>'}
    </tbody></table></div>`;
  $('#nuevo').onclick = () => editarSimple('comisionistas', null);
  $$('tr[data-id]').forEach(tr => tr.onclick = () => editarSimple('comisionistas', rows.find(r => r.id === tr.dataset.id)));
};
VIEWS.empresas = async main => {
  const rows = await cat('empresas', true);
  main.innerHTML = `<h2>Empresas</h2>
    <div class="toolbar"><span class="muted">Empresas que facturan y cuentas receptoras de los depósitos.</span><span class="sp"></span><button class="btn primary" id="nuevo">+ Nueva</button></div>
    <div class="tablewrap"><table><thead><tr><th>Empresa</th><th>RFC</th><th>Banco</th><th>Cuenta</th><th>Factura</th></tr></thead><tbody>
    ${rows.map(r => `<tr class="click" data-id="${r.id}"><td><b>${esc(r.nombre)}</b></td><td>${esc(r.rfc || '')}</td><td>${esc(r.banco || '')}</td><td class="mono">${esc(r.cuenta || '')}</td><td>${r.requiere_factura ? 'Sí' : 'No'}</td></tr>`).join('') || '<tr><td colspan="5" class="empty">Sin empresas</td></tr>'}
    </tbody></table></div>`;
  $('#nuevo').onclick = () => editarSimple('empresas', null);
  $$('tr[data-id]').forEach(tr => tr.onclick = () => editarSimple('empresas', rows.find(r => r.id === tr.dataset.id)));
};
const CAMPOS = {
  comisionistas: [['nombre', 'Nombre'], ['tipo_pago', 'Tipo de pago', TIPOS_PAGO], ['banco', 'Banco'], ['cuenta', 'Cuenta'], ['clabe', 'CLABE'], ['notas', 'Notas'], ['activo', 'Activo', 'bool']],
  empresas: [['nombre', 'Nombre'], ['rfc', 'RFC'], ['banco', 'Banco'], ['cuenta', 'Cuenta'], ['clabe', 'CLABE'], ['requiere_factura', 'Lleva factura', 'bool'], ['notas', 'Notas'], ['activo', 'Activo', 'bool']],
};
function editarSimple(table, r) {
  const campos = CAMPOS[table]; r = r || {activo: true, requiere_factura: true, tipo_pago: 'EFECTIVO'};
  modal(r.id ? r.nombre : 'Nuevo', `<div class="form">${campos.map(([k, l, t]) => `<label>${l}</label>` + (t === 'bool' ? `<label><input type="checkbox" id="f_${k}" ${r[k] ? 'checked' : ''}> Sí</label>` :
    Array.isArray(t) ? `<select id="f_${k}">${t.map(o => `<option ${o === r[k] ? 'selected' : ''}>${o}</option>`).join('')}</select>` : `<input id="f_${k}" value="${esc(r[k] ?? '')}">`)).join('')}</div><div class="err" id="fErr"></div>`, [
    ...(r.id ? [{label: 'Eliminar', cls: 'bad', onClick: async () => { if (!(await confirmar('Eliminar', `¿Eliminar ${esc(r.nombre)}?`, 'Eliminar', 'bad'))) return false; await db.remove(table, r.id); invalidate(table); refresh(); }}] : []),
    {label: 'Cancelar'},
    {label: 'Guardar', cls: 'primary', onClick: async () => {
      const row = Object.fromEntries(campos.map(([k]) => [k, val('f_' + k)])); row.nombre = String(row.nombre).toUpperCase();
      if (!row.nombre) { $('#fErr').textContent = 'Falta el nombre'; return false; }
      if (row.clabe) { const v = Logica.validarClabe(row.clabe); if (!v.ok) { $('#fErr').textContent = v.motivo; return false; } row.clabe = Logica.normalizarClabe(row.clabe); }
      r.id ? await db.update(table, r.id, row) : await db.insert(table, row); invalidate(table); toast('Guardado'); refresh();
    }}]);
}

/* ===================== Ajustes ===================== */
VIEWS.ajustes = async main => {
  const [sig, pre, formUrl] = await Promise.all([getConfig('folio_siguiente', 2355), getConfig('folio_prefijo', 'FACOL'), urlFormularios()]);
  const hayMachote = await db.exists('plantillas/contrato.docx').catch(() => false);
  const c = dbConn.get() || {};
  main.innerHTML = `<h2>Ajustes</h2>
    <div class="cols">
      <div class="panel"><h3 style="margin-top:0">Folios FACOL</h3>
        <div class="form"><label>Prefijo</label><input id="aPre" value="${esc(pre)}"><label>Siguiente número</label><input id="aSig" type="number" value="${esc(sig)}"></div>
        <p class="small muted">El próximo depósito recibirá el folio <b id="aPrev"></b>. Si el número ya existe, se salta al siguiente libre.</p>
        <button class="btn primary" id="aFol">Guardar folios</button></div>
      <div class="panel"><h3 style="margin-top:0">Formularios públicos</h3>
        <div class="form"><label>Dirección (URL)</label><input id="aUrl" placeholder="https://usuario.github.io/facol-formularios" value="${esc(formUrl)}"></div>
        <p class="small muted">Donde publicaste la carpeta <b>formularios</b> (ver LEEME). Con esto se arman las ligas del cuestionario de alta y de cambio de cuenta.</p>
        <button class="btn primary" id="aUrlG">Guardar</button></div>
      <div class="panel"><h3 style="margin-top:0">Machote de contrato</h3>
        <p class="small">${hayMachote ? tag('ok', 'Machote propio cargado') : tag('info', 'Usando el machote incluido (CONTRATO BASE)')}</p>
        <p class="small muted">Word (.docx) con campos entre corchetes: [SOCIO], [RFC], [FECHA INICIO], [HUELLA]…</p>
        <div class="toolbar"><button class="btn" id="aTpl">Subir machote .docx</button><button class="btn" id="aTplD">Descargar el actual</button><button class="btn" id="aTplC">Ver campos</button></div></div>
      <div class="panel"><h3 style="margin-top:0">Conexión</h3>
        <p class="small">${db.mode === 'demo' ? tag('warn', 'Modo demostración') + ' los datos están solo en esta PC.' : 'Supabase: <span class="mono">' + esc(c.url) + '</span>'}</p>
        <div class="toolbar"><button class="btn" id="aConn">Cambiar conexión</button>${db.mode === 'demo' ? '<button class="btn" id="aSeed">Cargar datos de ejemplo</button><button class="btn bad" id="aReset">Borrar demo</button>' : ''}</div>
        <p class="small muted">Versión ${esc(window.desk ? await desk.version() : 'web')}</p>
        ${db.mode === 'demo' ? '<button class="btn sm" id="aComo">Probar como usuario de solo consulta</button>' : ''}</div>
    </div>
    <div class="panel" style="margin-top:14px"><h3 style="margin-top:0">Actualizaciones</h3><div id="aUpd"></div></div>
    <div class="panel" style="margin-top:14px"><h3 style="margin-top:0">Usuarios</h3><div id="aUsr" class="muted">Cargando…</div></div>`;
  panelUsuarios($('#aUsr')); panelActualizaciones($('#aUpd'));
  if ($('#aComo')) $('#aComo').onclick = () => { sessionStorage.setItem('facol-demo-rol', 'consulta'); location.reload(); };
  const prev = () => $('#aPrev').textContent = val('aPre') + val('aSig'); $('#aPre').oninput = $('#aSig').oninput = prev; prev();
  $('#aFol').onclick = () => guard(async () => { await setConfig('folio_prefijo', val('aPre').toUpperCase()); await setConfig('folio_siguiente', +val('aSig')); }, 'Folios guardados');
  $('#aUrlG').onclick = () => guard(() => setConfig('formularios_url', normalizarUrlForm(val('aUrl'))).then(() => { $('#aUrl').value = normalizarUrlForm(val('aUrl')); }), 'Guardado');
  $('#aTpl').onclick = async () => { const [f] = await pickFiles('.docx'); if (!f) return;
    await guard(async () => { const campos = await Docx.campos(await f.arrayBuffer()); if (!campos.length) throw new Error('El documento no tiene campos [ENTRE CORCHETES]'); await db.upload('plantillas/contrato.docx', f); }, 'Machote actualizado'); refresh(); };
  $('#aTplD').onclick = async () => guardarBytes(new Uint8Array(await machote()), 'CONTRATO BASE.docx');
  $('#aTplC').onclick = async () => { const campos = await Docx.campos(await machote()); modal('Campos del machote', `<p>${campos.map(c => `<span class="tag info" style="margin:2px">[${esc(c)}]</span>`).join(' ')}</p>`, [{label: 'Cerrar', cls: 'primary'}]); };
  $('#aConn').onclick = () => { $('#side').style.display = 'none'; renderConexion(); };
  if ($('#aSeed')) $('#aSeed').onclick = () => guard(cargarEjemplo, 'Datos de ejemplo cargados').then(() => go('inicio'));
  if ($('#aReset')) $('#aReset').onclick = async () => { if (await confirmar('Borrar demostración', 'Se borran todos los datos de prueba de esta PC.', 'Borrar', 'bad')) { db.reset(); location.reload(); } };
};
const ROLES = {admin: 'Administrador (todo + usuarios)', operador: 'Operador (todo menos usuarios)', consulta: 'Solo consulta (contratos, BD, facturas, comprobantes)'};
async function panelUsuarios(el) {
  if (me.rol !== 'admin') { el.innerHTML = 'Solo un administrador puede dar de alta usuarios.'; return; }
  let users; try { users = await db.list('staff', {order: 'email'}); } catch (e) { el.innerHTML = esc(errMsg(e)); return; }
  el.classList.remove('muted');
  el.innerHTML = `<div class="tablewrap"><table><thead><tr><th>Correo</th><th>Nombre</th><th>Permiso</th><th></th></tr></thead><tbody>
    ${users.map(u => `<tr><td>${esc(u.email)}</td><td>${esc(u.nombre || '')}</td><td><select data-rol="${esc(u.email)}" ${u.email === me.email ? 'disabled' : ''}>${Object.entries(ROLES).map(([k, l]) => `<option value="${k}" ${k === u.rol ? 'selected' : ''}>${l}</option>`).join('')}</select></td>
      <td>${u.email === me.email ? '<span class="small muted">tú</span>' : `<button class="link" data-quita="${esc(u.email)}" style="color:var(--bad)">Quitar acceso</button>`}</td></tr>`).join('')}
    </tbody></table></div>
    <div class="form2" style="margin-top:10px"><label>Correo</label><input id="uE" type="email"><label>Nombre</label><input id="uN">
      <label>Permiso</label><select id="uR">${Object.entries(ROLES).map(([k, l]) => `<option value="${k}" ${k === 'consulta' ? 'selected' : ''}>${l}</option>`).join('')}</select><span></span><button class="btn primary" id="uAdd">Dar acceso</button></div>
    <p class="note info">Después de dar acceso, crea su usuario con contraseña en Supabase → <b>Authentication → Users → Add user</b> (marca "Auto Confirm User"). Con ese correo y contraseña entra a la app.</p>`;
  $$('[data-rol]', el).forEach(s => s.onchange = () => guard(() => db.updateWhere('staff', {email: s.dataset.rol}, {rol: s.value}), 'Permiso actualizado'));
  $$('[data-quita]', el).forEach(b => b.onclick = async () => { if (!(await confirmar('Quitar acceso', `¿Quitar el acceso de ${esc(b.dataset.quita)}?`, 'Quitar', 'bad'))) return;
    const u = users.find(x => x.email === b.dataset.quita); await guard(() => db.removeWhere('staff', {email: u.email}), 'Acceso quitado'); panelUsuarios(el); });
  $('#uAdd', el).onclick = () => guard(async () => { const email = val('uE').toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Correo no válido');
    await db.upsert('staff', [{email, nombre: val('uN'), rol: val('uR')}], 'email'); panelUsuarios(el); }, 'Acceso dado');
}

async function machote() {
  try { if (await db.exists('plantillas/contrato.docx')) return await (await db.download('plantillas/contrato.docx')).arrayBuffer(); } catch { /* usa el incluido */ }
  if (window.desk) { const u8 = await desk.machoteIncluido(); if (!u8) throw new Error('No hay machote de contrato: súbelo una vez en Ajustes → Machote de contrato'); return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength); }
  const r = await fetch('plantillas/CONTRATO BASE.docx');
  if (!r.ok) throw new Error('No hay machote de contrato: súbelo en Ajustes → Machote de contrato');
  return r.arrayBuffer();
}

/** Carga el esquema del FACOL2354 como ejemplo (cliente STAR SEEDS, sus promotores y empresas). */
async function cargarEjemplo() {
  const existe = (await cat('clientes', true)).find(c => c.nombre === 'STAR SEEDS'); if (existe) return;
  const emp = await db.insert('empresas', [{nombre: 'SAGVE SA DE CV', rfc: 'SAG140423N85', requiere_factura: true}, {nombre: 'BUILD FORCE', banco: 'BANORTE', cuenta: '****67295', requiere_factura: true}]);
  const nombres = ['ARTURO', 'CRISTIAN', 'OSCAR HERNANDEZ', 'MANUEL', 'CASA', 'LORENA', 'ABEL ORTIZ'];
  const pagos = ['EFECTIVO', 'TRANSFERENCIA/EFECTIVO', 'EFECTIVO', 'TRANSFERENCIA/EFECTIVO', 'CASA', 'TRANSFERENCIA/EFECTIVO', 'TRANSFERENCIA'];
  const com = await db.insert('comisionistas', nombres.map((n, i) => ({nombre: n, tipo_pago: pagos[i]})));
  const cli = await db.insert('clientes', {nombre: 'STAR SEEDS', razon_social: 'STAR SEEDS PROYECTOS DE SURESTE S.A.P.I. DE C.V.', pct_comision: 11.5,
    empresa_factura_id: emp[1].id, empresa_receptora_id: emp[1].id, banco_ordenante: 'BBVA', solicitante: 'ABEL ORTIZ', autorizado_por: 'DIRECCIÓN', requiere_factura: true});
  const ids = nombres.map(() => newId());
  const def = [['COMISIÓN PROMOTOR INTERNO', 'promotor', 1.4], ['COMISIÓN FLUJO', 'terciaria', 0.2], ['COMISIÓN PROMOTOR', 'promotor', 1.6], ['COMISIÓN PROMOTOR', 'promotor', 1],
    ['COMISIÓN ALSOCORP', 'casa', 7.3], ['COMISIÓN DE EFECTIVO ESA FLUJOS', 'terciaria', 0], ['COMISIÓN SOBRE CASA', 'terciaria', 35]];
  await db.insert('cliente_comisiones', def.map(([concepto, clase, p], i) => ({id: ids[i], cliente_id: cli.id, comisionista_id: com[i].id, concepto: concepto + ' ' + nombres[i], clase, pct: p,
    base_linea_id: i === 6 ? ids[4] : null, orden: i})));
  invalidate('clientes', 'comisionistas', 'empresas');
}
