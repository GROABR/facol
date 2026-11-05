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
VIEWS.consultaComprobantes = main => consultaArchivos(main, 'comprobante', 'Comprobantes de dispersión');
