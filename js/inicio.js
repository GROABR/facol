/* ---------- FACOL: tablero (según el módulo elegido) ---------- */
VIEWS.inicio = async main => {
  const esCol = esDisp(modulo);
  const [todas, inv, cam] = await Promise.all([db.list('operaciones', {eq: {estado: 'abierta', modalidad: modulo}, order: 'numero', desc: true}),
    db.list('invitaciones', {eq: {estado: 'respondida'}, select: 'id'}), db.list('cambios_cuenta', {eq: {estado: 'pendiente'}, select: 'id'})]);
  const ops = todas;
  const disp = esCol && ops.length ? await db.list('dispersiones', {in: {operacion_id: ops.map(o => o.id)}, select: 'operacion_id,estado,monto'}) : [];
  const pendDisp = disp.filter(d => d.estado !== 'dispersado');
  const listas = ops.filter(o => { const d = disp.filter(x => x.operacion_id === o.id); return d.length && d.every(x => x.estado === 'dispersado'); });
  const sinRetorno = ops.filter(o => !o.retorno_entregado);
  const sinFactura = ops.filter(o => o.requiere_factura && o.factura_estado !== 'emitida');
  const vacio = esCol ? !(await cat('beneficiarios', true)).length : !(await cat('clientes', true)).length;
  main.innerHTML = `<h2>${nombreModulo(modulo)} · Tablero</h2>
    ${vacio ? `<div class="note info"><b>Para empezar:</b> ${esCol ? 'revisa <b>Datos y comisiones</b> (presiona Guardar en ambos) e importa tu <b>BD ALTAS</b> en Beneficiarios' : 'da de alta tus <b>Clientes</b> y <b>Empresas</b> con su reparto de comisiones'}.
      ${db.mode === 'demo' ? '<br>En modo demostración puedes <button class="link" id="seed">cargar el ejemplo del FACOL2354</button>.' : ''}</div>` : ''}
    ${esCol ? '<div class="toolbar"><button class="btn primary" id="nueva">+ Nueva dispersión</button></div>' : ''}
    <div class="grid">
      <div class="card click" data-go="${esCol ? 'colorines' : 'depositos'}"><div class="k">Operaciones abiertas</div><div class="v">${ops.length}</div></div>
      ${esCol ? `<div class="card click" data-go="colorines"><div class="k">Colaboradores por dispersar</div><div class="v ${pendDisp.length ? 'warn' : 'ok'}">${pendDisp.length}</div><div class="small muted">${money(pendDisp.reduce((s, d) => s + +d.monto, 0))}</div></div>
      <div class="card click" data-go="colorines"><div class="k">Listas para cerrar</div><div class="v ${listas.length ? 'ok' : ''}">${listas.length}</div></div>`
      : `<div class="card click" data-go="depositos"><div class="k">Retornos por entregar</div><div class="v ${sinRetorno.length ? 'warn' : 'ok'}">${sinRetorno.length}</div></div>`}
      <div class="card"><div class="k">Facturas pendientes</div><div class="v ${sinFactura.length ? 'warn' : 'ok'}">${sinFactura.length}</div></div>
      ${esCol ? `<div class="card click" data-go="contratos"><div class="k">Encuestas por revisar</div><div class="v ${inv.length ? 'warn' : ''}">${inv.length}</div></div>
      <div class="card click" data-go="cambios"><div class="k">Cambios de cuenta</div><div class="v ${cam.length ? 'warn' : ''}">${cam.length}</div></div>` : ''}
    </div>
    <h3>Operaciones abiertas</h3>
    <div class="tablewrap"><table><thead><tr><th>Folio</th><th>Fecha</th><th>Cliente</th><th class="n">Depósito</th><th>Pendiente</th></tr></thead><tbody>
    ${ops.slice(0, 30).map(o => { const d = disp.filter(x => x.operacion_id === o.id); const p = [];
      if (esCol) p.push(d.length ? (d.some(x => x.estado !== 'dispersado') ? `${d.filter(x => x.estado !== 'dispersado').length} por dispersar` : 'lista para cerrar') : 'falta lista');
      else if (!o.retorno_entregado) p.push('retorno');
      if (o.requiere_factura && o.factura_estado !== 'emitida') p.push('factura');
      return `<tr class="click" data-id="${o.id}"><td><b>${esc(o.folio)}</b></td><td>${fmtDay(o.fecha_deposito)}</td><td>${esc(o.cliente_nombre)}</td><td class="n">${money(o.monto_deposito)}</td>
        <td>${p.map(x => tag(x === 'lista para cerrar' ? 'ok' : 'warn', x)).join(' ') || tag('ok', 'todo listo')}</td></tr>`; }).join('') || '<tr><td colspan="5" class="empty">No hay operaciones abiertas</td></tr>'}
    </tbody></table></div>`;
  $$('[data-go]').forEach(c => c.onclick = () => go(c.dataset.go));
  $$('tr[data-id]').forEach(tr => tr.onclick = () => go('operacion', {id: tr.dataset.id}));
  if ($('#nueva')) $('#nueva').onclick = () => go('colorinesNueva');
  if ($('#seed')) $('#seed').onclick = () => guard(cargarEjemplo, 'Ejemplo cargado').then(() => go('inicio'));
};

boot();
