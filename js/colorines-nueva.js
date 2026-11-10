/* ---------- FACOL: Colorines — pantalla de nueva dispersión, revisión de la solicitud, Excel FACOL y altas/contratos ---------- */

/** Lee el formato de solicitud (Excel) en cualquier hoja. */
async function leerSolicitud(file) {
  const {rows, sheets, wb} = await readSheetRows(file);
  let items, hojaUsada = sheets[0], err;
  for (const s of sheets) {
    try { items = Logica.parsePrenomina(s === sheets[0] ? rows : XLSX.utils.sheet_to_json(wb.Sheets[s], {header: 1, defval: '', raw: true})); hojaUsada = s; break; } catch (e) { err = err || e; }
  }
  if (!items) throw err;
  if (!items.length) throw new Error('El archivo no tiene personas con monto');
  return {items, hoja: hojaUsada};
}

/** Tabla de revisión: busca a cada persona en la BD y deja corregir. ctx = {pct, deposito(), onChange(lines, res)} */
function revisionLista(el, items, bens, ctx) {
  const idx = Logica.indexBens(bens);
  items.forEach(it => { if (it.match) return; const m = Logica.matchNombre(it.nombre_archivo, bens, idx); it.match = m; it.benId = m.tipo === 'ninguno' ? null : m.ben.id; });
  const lines = () => items.map(it => ({...it, ...Logica.grossUp(it.monto, ctx.pct()), ...datosBanco(bens.find(b => b.id === it.benId))}));
  const draw = () => {
    const ls = lines(); const nNo = ls.filter(l => !l.beneficiario_id).length; const nSug = items.filter(i => i.match.tipo === 'sugerido' && !i.confirmado).length;
    el.innerHTML = `<div class="toolbar"><span>${tag('ok', `${items.length - nNo} encontradas`)} ${nSug ? tag('warn', `${nSug} por confirmar`) : ''} ${nNo ? tag('bad', `${nNo} no están en la BD`) : ''}</span><span class="sp"></span>
      ${nSug ? '<button class="btn sm" id="rConf">Confirmar las parecidas</button>' : ''}</div>
      <div class="tablewrap" style="max-height:420px"><table><thead><tr><th>#</th><th>En la solicitud</th><th>Beneficiario en la BD</th><th>Modo</th><th class="n">Monto</th><th class="n">Comisión</th><th class="n">Total</th><th>Banco</th><th>CLABE</th></tr></thead><tbody>
      ${ls.map((l, i) => { const it = items[i]; const m = it.match;
        const opts = [...new Map([...(m.candidatos || []), ...(it.benId ? [bens.find(b => b.id === it.benId)] : [])].filter(Boolean).map(b => [b.id, b])).values()];
        const est = !l.beneficiario_id ? tag('bad', 'no está en la BD') : m.tipo === 'exacto' || it.confirmado ? '' : tag('warn', m.duplicado ? 'nombre repetido' : `parecido ${Math.round(m.score * 100)}%`);
        return `<tr><td>${i + 1}</td><td>${esc(it.nombre_archivo)}${it.puesto ? `<div class="small muted">${esc(it.puesto)}</div>` : ''}</td>
          <td><select data-i="${i}" class="rb" style="max-width:270px"><option value="">— sin beneficiario —</option>${opts.map(b => `<option value="${b.id}" ${b.id === it.benId ? 'selected' : ''}>${esc(b.nombre)}</option>`).join('')}<option value="__buscar">Buscar otro…</option></select>
            ${est} ${l.beneficiario_id ? avisoCuenta(l) : ''}</td>
          <td><select data-i="${i}" class="re">${Logica.ESTRATEGIAS.map(e => `<option ${e === it.estrategia ? 'selected' : ''}>${e}</option>`).join('')}</select></td>
          <td class="n">${money(l.monto)}</td><td class="n">${money(l.comision)}</td><td class="n">${money(l.total)}</td><td>${esc(l.banco || '')}</td><td class="mono">${esc(l.clabe || '')}</td></tr>`; }).join('')}
      </tbody></table></div>
      ${nNo ? '<p class="small muted">Las personas que no están en la BD se guardan pendientes y sin cuenta. Desde la pestaña "Altas y contratos" de la operación les mandas la encuesta; al aprobarla, su cuenta se liga sola.</p>' : ''}`;
    $$('.rb', el).forEach(s => s.onchange = async () => {
      const it = items[+s.dataset.i];
      if (s.value === '__buscar') { const b = await buscarBeneficiario(it.nombre_archivo, bens); if (b) { it.benId = b.id; it.match.candidatos = [b, ...(it.match.candidatos || [])]; it.confirmado = true; } return draw(); }
      it.benId = s.value || null; it.confirmado = true; draw();
    });
    $$('.re', el).forEach(s => s.onchange = () => { items[+s.dataset.i].estrategia = s.value; draw(); });
    if ($('#rConf', el)) $('#rConf', el).onclick = () => { items.forEach(i => { if (i.match.tipo === 'sugerido') i.confirmado = true; }); draw(); };
    ctx.onChange(ls, Logica.resumenDispersion(ls, ctx.pct(), ctx.deposito()), {nNo, nSug});
  };
  draw();
  return {lines, redraw: draw};
}
const rowDispersion = (l, opId, i) => ({operacion_id: opId, nombre_archivo: l.nombre_archivo, puesto: l.puesto || null, estrategia: l.estrategia,
  monto: l.monto, comision: l.comision, total: l.total, ...datosBanco(l.beneficiario_id ? {id: l.beneficiario_id, nombre: l.nombre, banco: l.banco, cuenta: l.cuenta, clabe: l.clabe, swift: l.swift, aba: l.aba} : null), orden: i});

/* ===================== Nueva dispersión ===================== */
VIEWS.colorinesNueva = async main => {
  const [bens, esq, dc] = await Promise.all([cat('beneficiarios', true), esquemaColorines(modulo), datosColorines(modulo)]);
  const st = {fichas: [], solicitud: null, items: null, rev: null, lines: [], res: null, info: {}};
  main.innerHTML = `<h2><button class="link back" id="bk">← ${esc(nombreModulo(modulo))}</button></h2>
    <h2>Nueva dispersión ${esc(nombreModulo(modulo))}</h2>
    <div class="cols">
      <div class="panel"><h3 style="margin-top:0">1. Ficha de pago</h3>
        <p class="small muted" style="margin-top:0">Cliente: <b>${esc(dc.cliente)}</b> · factura ${esc(dc.empresa_factura || '—')} <button class="link" id="nDat">cambiar datos</button></p>
        <div class="form"><label>Fecha de depósito</label><input type="date" id="nFec" value="${today()}">
          <label>Monto depositado</label><input type="number" step="0.01" id="nMon" placeholder="lo que dice la ficha">
          <label>% comisión</label><input type="number" step="0.0001" id="nPct" value="${esq.pct}">
          <label>Banco ordenante</label><input id="nBO" value="${esc(dc.banco_ordenante || '')}"><label>Solicitante</label><input id="nSol" value="${esc(dc.solicitante || '')}"></div>
        <div id="nFicha" style="margin-top:10px"></div><div class="files" id="nFichas" style="margin-top:6px"></div></div>
      <div class="panel"><h3 style="margin-top:0">2. Formato de solicitud</h3>
        <p class="small muted" style="margin-top:0">El Excel con los colaboradores y cuánto debe recibir cada uno (como "SAGVE DISP"). Columnas: NOMBRE, TOTAL/MONTO y MODO (SCU, PAYPAL, EFECTIVO, USA).</p>
        <div id="nSolDrop"></div><div id="nSolInfo" class="small" style="margin-top:6px"></div>
        <h3>3. Resultado</h3><div id="nRes" class="muted">Sube la solicitud para calcular.</div></div>
    </div>
    <h3>Colaboradores</h3><div id="nRev" class="empty">Aquí aparece la lista al subir la solicitud.</div>
    <h3>Comisiones de la operación <span class="small muted" style="font-weight:400">(esquema del módulo: Casa y promotores)</span></h3><div id="nCom" class="muted">Sube la solicitud para calcular las comisiones.</div>
    <div class="toolbar" style="margin-top:14px"><span class="sp"></span><span class="err" id="nErr"></span><button class="btn primary" id="nGo">Guardar, asignar folio y generar Excel FACOL</button></div>`;
  $('#bk').onclick = () => go('colorines');
  const pctNow = () => num('nPct') || 0; const depNow = () => num('nMon');
  const esquemaAhora = () => Logica.lineasColorines({...esq, pct: pctNow()});
  const calc = () => {
    const res = st.ls ? Logica.resumenDispersion(st.ls, pctNow(), depNow() || undefined) : null; st.res = res;
    if (res) {
      const dep = depNow();
      $('#nRes').innerHTML = `<div class="kv"><div>Colaboradores</div><div><b>${res.n}</b> ${Object.entries(res.porEstrategia).map(([k, v]) => tag('info', `${k} ${money(v)}`)).join(' ')}</div>
        <div>Monto a dispersar</div><div><b>${money(res.monto)}</b></div><div>Comisión ${pct(pctNow())}</div><div>${money(res.comision)}</div>
        <div>Depósito requerido</div><div><b>${money(res.total)}</b></div><div>Depósito de la ficha</div><div>${dep ? money(dep) : '<span class="muted">sin capturar (se usa el requerido)</span>'}</div>
        <div>Diferencia</div><div>${dep ? `<b style="color:${Math.abs(dep - res.total) < 1 ? 'var(--ok)' : 'var(--bad)'}">${money(dep - res.total)}</b>` : '—'}</div></div>`;
    }
    const base = res ? res.total : depNow();
    if (res && base) {
      const ls = Logica.ajustarCuadre(Logica.calcComisiones(base, esquemaAhora()), res.comision);
      const hijos = p => ls.filter(x => x.base_linea_id === p.id);
      $('#nCom').innerHTML = `<div class="tablewrap"><table><thead><tr><th>Concepto</th><th class="n">%</th><th>Forma de pago</th><th class="n">Comisión</th></tr></thead><tbody>
        ${ls.filter(l => !l.base_linea_id).map(l => `<tr><td><b>${esc(l.nombre)}</b></td><td class="n">${pct(l.pct)}</td><td>${esc(l.tipo_pago || '')}</td><td class="n">${money(l.monto)}</td></tr>` +
          hijos(l).map(x => `<tr class="sub"><td>↳ <b>${esc(x.nombre)}</b> (sale de ${esc(l.nombre)})</td><td class="n">${pct(x.pct)} de ${esc(l.nombre)}</td><td>${esc(x.tipo_pago || '')}</td><td class="n">${money(x.monto)}</td></tr>`).join('') +
          (hijos(l).length ? `<tr class="sub"><td>↳ ${esc(l.nombre)} neto</td><td></td><td></td><td class="n">${money(Logica.netoLinea(l, ls))}</td></tr>` : '')).join('')}
        </tbody><tfoot><tr><td>Total comisión</td><td></td><td></td><td class="n">${money(Logica.totalesComision(ls).total)}</td></tr></tfoot></table></div>`;
    }
  };
  $('#nDat').onclick = () => go('esquemaColorines');
  $('#nMon').oninput = calc; $('#nPct').oninput = () => { if (st.rev) st.rev.redraw(); else calc(); };
  const drawFichas = () => $('#nFichas').innerHTML = st.fichas.map((f, i) => `<div class="file"><span class="nm">📎 ${esc(f.name)}</span><button class="link" data-x="${i}" style="color:var(--bad)">Quitar</button></div>`).join('');
  $('#nFichas').onclick = e => { const i = e.target.dataset.x; if (i !== undefined) { st.fichas.splice(+i, 1); drawFichas(); } };
  dropZone($('#nFicha'), fs => { st.fichas.push(...fs); drawFichas(); }, {accept: 'image/*,.pdf', text: 'Ficha de pago (depósito): arrastra la imagen o PDF, o haz clic'});
  dropZone($('#nSolDrop'), async ([f]) => {
    await guard(async () => {
      const {items, hoja} = await leerSolicitud(f);
      st.solicitud = f; st.items = items;
      $('#nSolInfo').innerHTML = `📎 <b>${esc(f.name)}</b> · hoja ${esc(hoja)} · ${items.length} colaboradores`;
      $('#nRev').classList.remove('empty');
      st.rev = revisionLista($('#nRev'), items, bens, {pct: pctNow, deposito: () => depNow() || undefined, onChange: (ls, res, info) => { st.ls = ls; st.info = info; calc(); }});
    });
  }, {accept: '.xlsx,.xls,.xlsm,.csv', multiple: false, text: 'Formato de solicitud: arrastra el Excel o haz clic'});

  $('#nGo').onclick = async () => {
    $('#nErr').textContent = '';
    if (!st.rev) return ($('#nErr').textContent = 'Sube el formato de solicitud');
    const ls = st.rev.lines(); const res = Logica.resumenDispersion(ls, pctNow());
    const dep = depNow() || res.total; const avisos = [];
    if (!st.fichas.length) avisos.push('No subiste la ficha de pago');
    if (Math.abs(dep - res.total) >= 1) avisos.push(`El depósito (${money(dep)}) no coincide con lo requerido (${money(res.total)})`);
    if (st.info.nNo) avisos.push(`${st.info.nNo} colaborador(es) no están en la BD (quedan pendientes de alta)`);
    if (st.info.nSug) avisos.push(`${st.info.nSug} coincidencia(s) por parecido sin confirmar`);
    if (avisos.length && !(await confirmar('Antes de guardar', `<div class="note">${avisos.map(esc).join('<br>')}</div>¿Guardar de todos modos?`, 'Guardar'))) return;
    const b = $('#nGo'); busy(b, 'Guardando…');
    try {
      const {folio, numero} = await siguienteFolio(modulo);
      const fac = dc.requiere_factura !== false;
      const op = await db.insert('operaciones', {folio, numero, modalidad: modulo, cliente_id: null, cliente_nombre: dc.cliente, razon_social: dc.razon_social || dc.cliente,
        fecha_deposito: val('nFec') || today(), fecha_solicitud: today(), fecha_compromiso: val('nFec') || today(), monto_deposito: dep, pct_comision: pctNow(),
        monto_comision: res.comision, monto_retorno: res.monto, banco_ordenante: val('nBO'), cuenta_ordenante: dc.cuenta_ordenante || '', empresa_factura: dc.empresa_factura || '',
        banco_receptor: dc.banco_receptor || '', cuenta_receptora: dc.cuenta_receptora || '', solicitante: val('nSol'), autorizado_por: dc.autorizado_por || '',
        tipo_retorno: 'TRANSFERENCIA', requiere_factura: fac, factura_estado: fac ? 'pendiente' : 'no_aplica'});
      await copiarEsquemaColorines(op.id, res.total, res.comision, modulo);
      await db.insert('dispersiones', ls.map((l, i) => rowDispersion(l, op.id, i)));
      for (const f of st.fichas) await subirArchivo(f, {operacion_id: op.id, tipo: 'ficha_deposito'});
      await subirArchivo(st.solicitud, {operacion_id: op.id, tipo: 'solicitud'});
      toast(`Folio asignado: ${folio}`);
      await excelFacol(op);
      go('operacion', {id: op.id, tab: 'dispersion'});
    } catch (e) { console.error(e); busy(b); $('#nErr').textContent = errMsg(e); }
  };
};

/** Para una operación que se quedó sin lista: carga el formato de solicitud ahí mismo. */
async function cargarListaEnOperacion(el, op) {
  if (op.estado !== 'abierta') { el.innerHTML = '<div class="empty">Operación cerrada sin lista de dispersión.</div>'; return; }
  el.innerHTML = `<div class="panel"><h3 style="margin-top:0">Formato de solicitud</h3><div id="lDrop"></div></div><div id="lRev" style="margin-top:12px"></div>`;
  dropZone($('#lDrop'), async ([f]) => guard(async () => {
    const {items} = await leerSolicitud(f); const bens = await cat('beneficiarios', true);
    const box = $('#lRev'); box.innerHTML = '<div id="lTab"></div><div class="toolbar" style="margin-top:10px"><span class="sp"></span><span id="lSum" class="muted"></span><button class="btn primary" id="lSave">Guardar lista</button></div>';
    const rev = revisionLista($('#lTab'), items, bens, {pct: () => +op.pct_comision, deposito: () => +op.monto_deposito || undefined,
      onChange: (ls, res) => { $('#lSum') && ($('#lSum').textContent = `${res.n} personas · ${money(res.monto)} + comisión ${money(res.comision)} = ${money(res.total)}`); }});
    $('#lSave').onclick = () => guard(async () => {
      const ls = rev.lines(); const res = Logica.resumenDispersion(ls, +op.pct_comision);
      await db.insert('dispersiones', ls.map((l, i) => rowDispersion(l, op.id, i)));
      await subirArchivo(f, {operacion_id: op.id, tipo: 'solicitud'});
      await db.update('operaciones', op.id, {monto_comision: res.comision, monto_retorno: res.monto, ...(+op.monto_deposito ? {} : {monto_deposito: res.total})});
      await recalcularComisiones(op.id, res.total, res.comision);
      toast('Lista guardada'); go('operacion', {id: op.id, tab: 'dispersion'});
    });
  }), {accept: '.xlsx,.xls,.xlsm,.csv', multiple: false, text: 'Arrastra aquí el Excel de la solicitud o haz clic'});
}

/** Genera el Excel FACOL de la operación (cálculo de comisiones, SSP, prenómina y formato de dispersión). */
/** Pregunta en cuántas facturas se divide el depósito (recuerda la última respuesta de cada operación). */
function preguntarFacturas(op) {
  const key = 'facol-nfact-' + op.id; let prev = 1; try { prev = +localStorage.getItem(key) || 1; } catch { /* sin almacenamiento */ }
  return new Promise(res => {
    let r = null;
    modal(`Solicitud de factura · ${op.folio}`, `<p style="margin-top:0">El Excel lleva una hoja oculta con la solicitud de factura del depósito (${money(op.monto_deposito)}).</p>
      <div class="form"><label>¿En cuántas facturas se divide?</label><input id="nfN" type="number" min="1" max="${MAX_FACTURAS}" step="1" value="${Math.min(prev, MAX_FACTURAS)}"></div><div id="nfV" class="small muted" style="margin-top:8px"></div>`, [
      {label: 'Cancelar', onClick: () => res(null)},
      {label: 'Generar Excel', cls: 'primary', onClick: () => { const n = Math.min(MAX_FACTURAS, Math.max(1, Math.floor(num('nfN')) || 1)); try { localStorage.setItem(key, n); } catch { /* */ } r = n; res(n); }}]);
    const ver = () => { const n = Math.min(MAX_FACTURAS, Math.max(1, Math.floor(num('nfN')) || 1)); const t = r2(op.monto_deposito / n);
      $('#nfV').textContent = `${n} factura(s) de aprox. ${money(t)} (subtotal ${money(r2(t / 1.16))} + IVA ${money(r2(t - t / 1.16))})`; };
    $('#nfN').oninput = ver; ver(); $('#nfN').focus();
    dlg.onclose = () => res(r);
  });
}
/** Formato de factura (copia del FACOL de ejemplo) que viaja con la app. */
async function plantillaFactura() {
  try {
    if (window.desk && desk.leerFormato) { const u8 = await desk.leerFormato('formatos/formato-factura.xlsx'); return u8 ? u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) : null; }
    const r = await fetch('formatos/formato-factura.xlsx'); return r.ok ? r.arrayBuffer() : null;
  } catch { return null; }
}
const MAX_FACTURAS = 4;   // bloques que trae la hoja FORMATO FACTURA
async function excelFacol(op) {
  let factura = null;
  if (op.requiere_factura !== false) { const n = await preguntarFacturas(op); if (!n) return null; factura = {n, datos: await datosColorines(op.modalidad), plantilla: await plantillaFactura()}; }
  const [lineas, comisiones, coms] = await Promise.all([db.list('dispersiones', {eq: {operacion_id: op.id}, order: 'orden'}), db.list('operacion_comisiones', {eq: {operacion_id: op.id}, order: 'orden'}), cat('comisionistas')]);
  const bytes = await FacolExcel.generar(ExcelJS, {op, lineas, comisiones, coms, factura});
  return guardarBytes(bytes, `${op.folio}.xlsx`);
}

/* ===================== Altas y contratos de los colaboradores de la operación ===================== */
async function tabAltas(el, op, disp) {
  const [inv, bens] = await Promise.all([db.list('invitaciones', {neq: {estado: 'cancelada'}, order: 'created_at', desc: true}), cat('beneficiarios', true)]);
  const benIds = [...new Set(disp.map(d => d.beneficiario_id).filter(Boolean))];
  const cons = benIds.length ? await db.list('contratos', {in: {beneficiario_id: benIds}, order: 'created_at', desc: true}) : [];
  const url = await urlFormularios();
  const key = s => Logica.tokens(s).sort().join(' ');
  const invDe = d => inv.find(i => (d.beneficiario_id && i.beneficiario_id === d.beneficiario_id) || key(i.nombre) === key(d.nombre_archivo) ||
    (i.respuesta && Logica.nameScore(Logica.valoresContrato(i.respuesta).SOCIO, d.nombre_archivo) >= 0.86));
  const filas = disp.map(d => {
    const b = bens.find(x => x.id === d.beneficiario_id); const c = b && cons.find(x => x.beneficiario_id === b.id); const i = invDe(d);
    let estado, accion;
    if (c) { estado = tag('ok', 'contrato generado ' + fmtDay(c.fecha_inicio)); accion = `<button class="link" data-dl="${c.id}">Descargar</button> · <button class="link" data-gen="${d.id}">Generar otro</button>`; }
    else if (i && i.estado === 'respondida') { estado = tag('warn', 'encuesta contestada'); accion = `<button class="btn sm primary" data-rev="${i.id}">Revisar y generar contrato</button>`; }
    else if (b && (i && ['aprobada', 'contrato'].includes(i.estado) || Object.keys(b.datos || {}).length)) { estado = tag('info', 'datos completos'); accion = `<button class="btn sm primary" data-gen="${d.id}">Generar contrato</button>`; }
    else if (i && i.estado === 'enviada') { estado = tag('info', 'encuesta enviada ' + fmtDay(i.created_at)); accion = `<button class="link" data-mail="${i.id}">Reenviar correo</button> · <button class="link" data-link="${i.id}">Ver liga</button>`; }
    else { estado = b ? tag('', 'sin encuesta') : tag('bad', 'no está en la BD'); accion = `<button class="btn sm" data-env="${d.id}">Enviar encuesta</button>`; }
    return {d, b, i, estado, accion, falta: !c && !(i && ['enviada', 'respondida'].includes(i.estado)) && !(b && Object.keys(b.datos || {}).length)};
  });
  const faltan = filas.filter(f => f.falta);
  el.innerHTML = `<div class="toolbar"><span class="muted">Encuesta de alta y contrato de cada colaborador de esta dispersión.</span><span class="sp"></span>
      ${faltan.length ? `<button class="btn primary" id="aAll">Enviar encuesta a los que faltan (${faltan.length})</button>` : ''}<button class="btn" id="aRef">Actualizar</button></div>
    ${url ? '' : '<div class="note">Configura en Ajustes la dirección de los formularios para que las ligas funcionen.</div>'}
    <div class="tablewrap"><table><thead><tr><th>#</th><th>Colaborador</th><th>En la BD</th><th>Estado</th><th></th></tr></thead><tbody>
    ${filas.map((f, n) => `<tr><td>${n + 1}</td><td><b>${esc(f.d.nombre_archivo)}</b></td><td>${f.b ? esc(f.b.nombre) : '—'}</td><td>${f.estado}</td><td style="white-space:nowrap">${f.accion}</td></tr>`).join('')}
    </tbody></table></div>`;
  $('#aRef').onclick = refresh;
  const correoDe = d => { const b = bens.find(x => x.id === d.beneficiario_id); return (b && b.correo) || ''; };
  const crear = (d, correo) => db.insert('invitaciones', {nombre: String(d.nombre || d.nombre_archivo).toUpperCase(), correo: correo || correoDe(d) || null, beneficiario_id: d.beneficiario_id || null, asesor: op.folio});
  $$('[data-env]', el).forEach(b => b.onclick = () => guard(async () => { const i = await crear(disp.find(d => d.id === b.dataset.env)); pedirCorreoYEnviar(i, url); }));
  $$('[data-mail]', el).forEach(b => b.onclick = () => pedirCorreoYEnviar(inv.find(i => i.id === b.dataset.mail), url));
  $$('[data-link]', el).forEach(b => b.onclick = () => mostrarLiga(inv.find(i => i.id === b.dataset.link), url));
  $$('[data-rev]', el).forEach(b => b.onclick = () => revisarAlta(inv.find(i => i.id === b.dataset.rev)));
  $$('[data-gen]', el).forEach(b => b.onclick = () => { const f = filas.find(x => x.d.id === b.dataset.gen); if (!f.b) return toast('Primero debe estar en la BD'); generarContrato(f.b, f.i && f.i.beneficiario_id === f.b.id ? f.i : null); });
  $$('[data-dl]', el).forEach(b => b.onclick = () => { const c = cons.find(x => x.id === b.dataset.dl); const bn = bens.find(x => x.id === c.beneficiario_id); guard(() => descargarArchivo(c.path, `${bn ? bn.nombre : 'CONTRATO'}.docx`)); });
  if ($('#aAll')) $('#aAll').onclick = () => {
    modal(`Enviar encuesta por correo (${faltan.length})`, `<p class="small muted" style="margin-top:0">Revisa el correo de cada colaborador. A quien no tenga correo no se le envía.</p>
      <div class="tablewrap" style="max-height:50vh"><table><thead><tr><th>Colaborador</th><th>Correo</th></tr></thead><tbody>${faltan.map((f, n) => `<tr><td>${esc(f.d.nombre || f.d.nombre_archivo)}</td><td><input type="email" data-n="${n}" value="${esc(correoDe(f.d))}" style="width:260px"></td></tr>`).join('')}</tbody></table></div>`, [
      {label: 'Cancelar'},
      {label: 'Enviar correos', cls: 'primary', onClick: async () => {
        let env = 0, red = 0, sin = 0;
        for (const inp of $$('#dlg [data-n]')) {
          const c = inp.value.trim().toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c)) { sin++; continue; }
          const i = await crear(faltan[+inp.dataset.n].d, c); (await enviarCorreoAlta(i, url)) === 'enviado' ? env++ : red++;
        }
        toast(`Enviados: ${env}${red ? ` · por enviar desde tu correo: ${red}` : ''}${sin ? ` · sin correo: ${sin}` : ''}`); refresh();
      }}], {width: '640px'});
  };
}
