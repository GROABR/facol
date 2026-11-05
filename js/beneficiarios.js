/* ---------- FACOL: base de beneficiarios (BD ALTAS) y solicitudes de cambio de cuenta ---------- */
const BEN_CAMPOS = [['apellido_paterno', 'Apellido paterno'], ['apellido_materno', 'Apellido materno'], ['nombres', 'Nombre(s)'], ['banco', 'Banco'], ['cuenta', 'Cuenta'], ['clabe', 'CLABE'],
  ['swift', 'SWIFT'], ['aba', 'ABA'], ['rfc', 'RFC'], ['curp', 'CURP'], ['correo', 'Correo'], ['celular', 'Celular'], ['observaciones', 'Observaciones']];
const nombreCompleto = d => [d.apellido_paterno, d.apellido_materno, d.nombres].map(x => String(x || '').trim()).filter(Boolean).join(' ').toUpperCase();
function alertaBen(b) {
  if (!b.clabe) return b.cuenta ? 'sin CLABE' : 'sin cuenta';
  const v = Logica.validarClabe(b.clabe, b.banco); return !v.ok ? v.motivo : v.aviso || '';
}

VIEWS.beneficiarios = async (main, arg) => {
  const bens = await cat('beneficiarios', true);
  const f = Object.assign({q: '', alerta: false}, arg || {});
  main.innerHTML = `<h2>Beneficiarios <span class="muted" style="font-size:14px;font-weight:400">${bens.length} en la base</span></h2>
    <div class="toolbar"><input id="bq" placeholder="Buscar por nombre, RFC, CURP, CLABE" style="width:300px" value="${esc(f.q)}">
      <label><input type="checkbox" id="ba" ${f.alerta ? 'checked' : ''}> Solo con alerta de cuenta</label><span class="sp"></span>
      <button class="btn" id="bImp">Importar BD ALTAS (.xlsx)</button><button class="btn" id="bExp">Exportar</button><button class="btn primary" id="bNew">+ Nuevo</button></div>
    <div class="tablewrap tall"><table><thead><tr><th>#</th><th>Beneficiario</th><th>Banco</th><th>Cuenta</th><th>CLABE</th><th>RFC</th><th>Alerta</th></tr></thead><tbody id="bTb"></tbody></table></div>
    <p class="small muted" id="bCnt"></p>`;
  const draw = () => {
    const q = Logica.normName(val('bq')), qd = val('bq').replace(/\D/g, ''), soloA = val('ba');
    const rows = bens.filter(b => (!q || Logica.normName([b.nombre, b.nombre_alt, b.rfc, b.curp].join(' ')).includes(q) || (qd.length >= 4 && (String(b.clabe || '') + ' ' + String(b.cuenta || '')).includes(qd))) && (!soloA || alertaBen(b)));
    $('#bTb').innerHTML = rows.slice(0, 400).map(b => { const a = alertaBen(b);
      return `<tr class="click" data-id="${b.id}"><td>${b.indice ?? ''}</td><td><b>${esc(b.nombre)}</b>${b.activo === false ? ' ' + tag('cancelada', 'baja') : ''}</td><td>${esc(b.banco || '')}</td>
        <td class="mono">${esc(b.cuenta || '')}</td><td class="mono">${esc(b.clabe || '')}</td><td>${esc(b.rfc || '')}</td><td>${a ? tag(/inválid|incorrect|dígitos|sin cuenta/i.test(a) ? 'bad' : 'warn', a) : ''}</td></tr>`; }).join('') || '<tr><td colspan="7" class="empty">Sin resultados. Importa tu BD ALTAS para empezar.</td></tr>';
    $('#bCnt').textContent = rows.length > 400 ? `Mostrando 400 de ${rows.length}; afina la búsqueda.` : `${rows.length} resultado(s)`;
    $$('#bTb tr[data-id]').forEach(tr => tr.onclick = () => editarBeneficiario(bens.find(b => b.id === tr.dataset.id)));
  };
  $('#bq').oninput = draw; $('#ba').onchange = draw; draw();
  $('#bNew').onclick = () => editarBeneficiario(null);
  $('#bImp').onclick = importarBd;
  $('#bExp').onclick = () => { const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, hoja(bens, [{h: 'INDICE', v: 'indice'}, {h: 'PRENOMINA', v: 'nombre_alt', w: 30}, {h: 'BENEFICIARIO', v: 'nombre', w: 40}, {h: 'BANCO BENEFICIARIO', v: 'banco', w: 16},
      {h: 'CUENTA', v: 'cuenta', w: 14}, {h: 'CLABE INTERBANCARIA', v: 'clabe', w: 21}, {h: 'SWFIT CODE', v: 'swift'}, {h: 'ABA NUMBER', v: 'aba'}, {h: 'RFC', v: 'rfc', w: 15}, {h: 'CURP', v: 'curp', w: 20},
      {h: 'CORREO', v: 'correo', w: 26}, {h: 'CELULAR', v: 'celular', w: 13}, {h: 'OBSERVACIONES', v: 'observaciones', w: 30}, {h: 'ALERTA', v: b => alertaBen(b), w: 30}], ['CUENTA', 'CLABE INTERBANCARIA', 'CELULAR', 'ABA NUMBER']), 'BDCOLBCOL');
    guardarXlsx(wb, `BD ALTAS ${today()}.xlsx`); };
};

async function editarBeneficiario(b) {
  const nuevo = !b; b = b || {activo: true, datos: {}};
  const [hist, contratos] = nuevo ? [[], []] : await Promise.all([db.list('beneficiario_cuentas_hist', {eq: {beneficiario_id: b.id}, order: 'cambiado_at', desc: true}), db.list('contratos', {eq: {beneficiario_id: b.id}, order: 'created_at', desc: true})]);
  const datos = b.datos || {};
  modal(nuevo ? 'Nuevo beneficiario' : b.nombre, `<div class="form2">
      <label>Nombre completo (BD)</label><input id="xNom" value="${esc(b.nombre || '')}" placeholder="APELLIDOS NOMBRES">
      <label>Como aparece en prenómina</label><input id="xAlt" value="${esc(b.nombre_alt || '')}">
      ${BEN_CAMPOS.map(([k, l]) => `<label>${l}</label><input id="x_${k}" value="${esc(b[k] ?? '')}" class="${k === 'clabe' && b.clabe && !Logica.validarClabe(b.clabe).ok ? 'bad' : ''}">`).join('')}
      <label>Activo</label><label><input type="checkbox" id="xAct" ${b.activo !== false ? 'checked' : ''}> Activo</label></div>
    <div id="xCl" class="small" style="margin-top:6px"></div>
    ${Object.keys(datos).length ? `<h3>Datos del cuestionario de alta</h3><div class="kv small">${Object.entries(datos).filter(([, v]) => typeof v !== 'object').map(([k, v]) => `<div>${esc(k)}</div><div>${esc(v)}</div>`).join('')}</div>` : ''}
    ${hist.length ? `<h3>Historial de cuentas</h3><table><tr><th>Fecha</th><th>Banco</th><th>Cuenta</th><th>CLABE</th><th>Origen</th></tr>${hist.map(h => `<tr><td>${fmtDate(h.cambiado_at)}</td><td>${esc(h.banco || '')}</td><td class="mono">${esc(h.cuenta || '')}</td><td class="mono">${esc(h.clabe || '')}</td><td>${esc(h.origen || '')} ${esc(h.cambiado_por || '')}</td></tr>`).join('')}</table>` : ''}
    ${contratos.length ? `<h3>Contratos</h3>${contratos.map(c => `<div class="file"><span class="nm">📄 Contrato ${fmtDay(c.fecha_inicio)}</span><button class="link" data-ct="${esc(c.path)}">Descargar</button></div>`).join('')}` : ''}`, [
    {label: 'Cancelar'},
    ...(!nuevo ? [{label: 'Generar contrato', onClick: () => { generarContrato(b); return false; }}] : []),
    {label: 'Guardar', cls: 'primary', onClick: async () => {
      const row = Object.fromEntries(BEN_CAMPOS.map(([k]) => [k, val('x_' + k) || null]));
      row.nombre = (val('xNom') || nombreCompleto(row)).toUpperCase().replace(/\s+/g, ' '); row.nombre_alt = val('xAlt') || null; row.activo = val('xAct');
      row.banco = row.banco ? Logica.normBanco(row.banco) : (Logica.bancoDeClabe(row.clabe) || null);
      if (!row.nombre) { toast('Falta el nombre'); return false; }
      if (row.clabe) { row.clabe = Logica.normalizarClabe(row.clabe); const v = Logica.validarClabe(row.clabe); if (!v.ok && !(await confirmar('CLABE con error', esc(v.motivo) + '. ¿Guardar de todos modos?'))) return false; }
      row.updated_at = new Date().toISOString();
      if (nuevo) { const n = await db.insert('beneficiarios', row); await db.insert('beneficiario_cuentas_hist', {beneficiario_id: n.id, banco: row.banco, cuenta: row.cuenta, clabe: row.clabe, origen: 'alta'}); }
      else {
        if (row.clabe !== b.clabe || row.cuenta !== b.cuenta || row.banco !== b.banco) await db.insert('beneficiario_cuentas_hist', {beneficiario_id: b.id, banco: b.banco, cuenta: b.cuenta, clabe: b.clabe, origen: 'edicion (cuenta anterior)'});
        await db.update('beneficiarios', b.id, row);
      }
      invalidate('beneficiarios'); toast('Beneficiario guardado'); if (view === 'beneficiarios') refresh();
    }}], {width: '900px'});
  const chk = () => { const v = val('x_clabe'); if (!v) return ($('#xCl').innerHTML = ''); const r = Logica.validarClabe(v, val('x_banco'));
    $('#xCl').innerHTML = r.ok ? tag('ok', 'CLABE válida · ' + (r.banco || 'banco no identificado')) + (r.aviso ? ' ' + tag('warn', r.aviso) : '') : tag('bad', r.motivo); $('#x_clabe').classList.toggle('bad', !r.ok); };
  $('#x_clabe').oninput = chk; $('#x_banco').oninput = chk; chk();
  $$('[data-ct]').forEach(a => a.onclick = () => guard(() => descargarArchivo(a.dataset.ct, `CONTRATO ${b.nombre}.${extDe(a.dataset.ct)}`)));
}

async function importarBd() {
  const [f] = await pickFiles('.xlsx,.xls,.xlsm'); if (!f) return;
  await guard(async () => {
    const {rows, sheets, wb} = await readSheetRows(f, 'BDCOLBCOL');
    let items; try { items = Logica.parseBdAltas(rows); } catch (e) {
      for (const s of sheets) { try { items = Logica.parseBdAltas(XLSX.utils.sheet_to_json(wb.Sheets[s], {header: 1, defval: '', raw: true})); break; } catch { /* otra hoja */ } }
      if (!items) throw e;
    }
    const bens = await cat('beneficiarios', true); const idx = Logica.indexBens(bens);
    const nuevos = [], cambios = [], dup = new Set();
    for (const it of items) {
      const key = Logica.tokens(it.nombre).sort().join(' '); if (dup.has(key)) continue; dup.add(key);
      const ex = (idx.get(key) || [])[0];
      if (!ex) nuevos.push(it);
      else if ((it.clabe && it.clabe !== ex.clabe) || (it.cuenta && it.cuenta !== ex.cuenta) || (it.banco && it.banco !== ex.banco)) cambios.push({ex, it});
    }
    const malos = items.filter(i => i.clabe && !Logica.validarClabe(i.clabe).ok).length;
    modal('Importar BD ALTAS', `<p>${items.length} registros leídos de <b>${esc(f.name)}</b>.</p><ul><li><b>${nuevos.length}</b> nuevos</li><li><b>${cambios.length}</b> con cuenta distinta a la que ya está en la app (se actualiza y se guarda la anterior en el historial)</li>
      <li>${items.length - nuevos.length - cambios.length} sin cambios o repetidos en el archivo</li>${malos ? `<li>${tag('warn', malos + ' con CLABE inválida')} (se importan; aparecerán con alerta)</li>` : ''}</ul>
      ${cambios.length ? `<details><summary>Ver cambios de cuenta</summary><table class="small">${cambios.slice(0, 100).map(c => `<tr><td>${esc(c.ex.nombre)}</td><td class="mono">${esc(c.ex.clabe || '—')}</td><td>→</td><td class="mono">${esc(c.it.clabe || '—')}</td></tr>`).join('')}</table></details>` : ''}`, [
      {label: 'Cancelar'}, {label: 'Importar', cls: 'primary', onClick: async () => {
        for (let i = 0; i < nuevos.length; i += 300) await db.insert('beneficiarios', nuevos.slice(i, i + 300));
        for (const {ex, it} of cambios) {
          await db.insert('beneficiario_cuentas_hist', {beneficiario_id: ex.id, banco: ex.banco, cuenta: ex.cuenta, clabe: ex.clabe, origen: 'importación (cuenta anterior)'});
          await db.update('beneficiarios', ex.id, {banco: it.banco || ex.banco, cuenta: it.cuenta || ex.cuenta, clabe: it.clabe || ex.clabe, updated_at: new Date().toISOString()});
        }
        invalidate('beneficiarios'); toast(`Importados ${nuevos.length} nuevos, ${cambios.length} actualizados`); go('beneficiarios');
      }}]);
  });
}

/* ===================== cambios de cuenta (formulario público) ===================== */
VIEWS.cambios = async (main, arg) => {
  const estado = (arg && arg.estado) ?? 'pendiente';
  const [rows, bens] = await Promise.all([db.list('cambios_cuenta', {...(estado ? {eq: {estado}} : {}), order: 'created_at', desc: true}), cat('beneficiarios', true)]);
  const url = await getConfig('formularios_url', FORMS_URL);
  main.innerHTML = `<h2>Solicitudes de cambio de cuenta</h2>
    <div class="toolbar"><select id="ce">${[['pendiente', 'Pendientes'], ['aplicado', 'Aplicadas'], ['rechazado', 'Rechazadas'], ['', 'Todas']].map(([k, l]) => `<option value="${k}" ${k === estado ? 'selected' : ''}>${l}</option>`).join('')}</select>
      <span class="sp"></span>${url ? `<input readonly value="${esc(url)}/cambio.html" style="width:340px" id="cUrl"><button class="btn" id="cCopy">Copiar liga del formulario</button>` : '<span class="muted">Configura la dirección de los formularios en Ajustes para obtener la liga.</span>'}</div>
    <div class="note info">Nada se cambia solo: revisa cada solicitud, confirma con la persona por un medio conocido si hay duda, y presiona "Aplicar". La cuenta anterior queda en el historial.</div>
    <div class="tablewrap"><table><thead><tr><th>Fecha</th><th>Solicitante</th><th>Correo</th><th>Nueva cuenta</th><th>Validación</th><th>Coincide con</th><th>Estado</th></tr></thead><tbody>
    ${rows.map(r => { const n = [r.apellido_paterno, r.apellido_materno, r.nombres].join(' '); const v = Logica.validarClabe(r.clabe, r.banco); const m = Logica.matchNombre(n, bens);
      return `<tr class="click" data-id="${r.id}"><td>${fmtDate(r.created_at)}</td><td><b>${esc(n)}</b></td><td>${esc(r.correo || '')}</td><td>${esc(r.banco || '')}<div class="mono small">${esc(r.cuenta || '')} · ${esc(r.clabe || '')}</div></td>
        <td>${v.ok ? (v.aviso ? tag('warn', v.aviso) : tag('ok', 'CLABE válida')) : tag('bad', v.motivo)}</td><td>${m.tipo === 'ninguno' ? tag('bad', 'no está en la BD') : esc(m.ben.nombre) + (m.tipo === 'sugerido' ? ' ' + tag('warn', 'parecido') : '')}</td><td>${tag(r.estado)}</td></tr>`; }).join('') || '<tr><td colspan="7" class="empty">Sin solicitudes</td></tr>'}
    </tbody></table></div>`;
  $('#ce').onchange = () => go('cambios', {estado: val('ce')});
  if ($('#cCopy')) $('#cCopy').onclick = () => { navigator.clipboard.writeText($('#cUrl').value); toast('Liga copiada'); };
  $$('tr[data-id]').forEach(tr => tr.onclick = () => revisarCambio(rows.find(r => r.id === tr.dataset.id), bens));
};

function revisarCambio(r, bens) {
  const n = [r.apellido_paterno, r.apellido_materno, r.nombres].join(' '); const m = Logica.matchNombre(n, bens);
  let benId = r.beneficiario_id || (m.tipo !== 'ninguno' ? m.ben.id : null);
  const v = Logica.validarClabe(r.clabe, r.banco);
  modal('Cambio de cuenta: ' + n, `<div class="form"><label>Beneficiario</label><select id="rb"><option value="">— no está en la BD (crear nuevo) —</option>${[...new Map([...(m.candidatos || []), ...(benId ? [bens.find(b => b.id === benId)] : [])].filter(Boolean).map(b => [b.id, b])).values()].map(b => `<option value="${b.id}" ${b.id === benId ? 'selected' : ''}>${esc(b.nombre)}</option>`).join('')}</select></div>
    <div id="rc" style="margin-top:10px"></div>
    <p>${v.ok ? tag('ok', 'CLABE válida · ' + (v.banco || '')) + (v.aviso ? ' ' + tag('warn', v.aviso) : '') : tag('bad', v.motivo)}</p>
    <div class="form"><label>Notas</label><input id="rn" value="${esc(r.notas || '')}"></div>`, r.estado !== 'pendiente' ? [{label: 'Cerrar', cls: 'primary'}] : [
    {label: 'Rechazar', cls: 'bad', onClick: async () => { await db.update('cambios_cuenta', r.id, {estado: 'rechazado', notas: val('rn'), revisado_por: me.email, revisado_at: new Date().toISOString()}); toast('Solicitud rechazada'); refresh(); }},
    {label: 'Cancelar'},
    {label: 'Aplicar cambio', cls: 'primary', onClick: async () => {
      if (!v.ok && !(await confirmar('CLABE con error', esc(v.motivo) + '. ¿Aplicar de todos modos?'))) return false;
      const banco = Logica.normBanco(r.banco) || v.banco || null; const clabe = Logica.normalizarClabe(r.clabe);
      let id = val('rb');
      if (id) { const b = bens.find(x => x.id === id);
        await db.insert('beneficiario_cuentas_hist', {beneficiario_id: id, banco: b.banco, cuenta: b.cuenta, clabe: b.clabe, origen: 'cambio de cuenta (anterior)'});
        await db.update('beneficiarios', id, {banco, cuenta: r.cuenta, clabe, correo: b.correo || r.correo, updated_at: new Date().toISOString()});
      } else { const nb = await db.insert('beneficiarios', {nombre: n.toUpperCase().replace(/\s+/g, ' ').trim(), apellido_paterno: r.apellido_paterno, apellido_materno: r.apellido_materno, nombres: r.nombres, banco, cuenta: r.cuenta, clabe, correo: r.correo}); id = nb.id; }
      await db.update('cambios_cuenta', r.id, {estado: 'aplicado', beneficiario_id: id, notas: val('rn'), revisado_por: me.email, revisado_at: new Date().toISOString()});
      // dispersiones pendientes de esa persona toman la cuenta nueva
      const pend = await db.list('dispersiones', {eq: {beneficiario_id: id, estado: 'pendiente'}});
      for (const d of pend) await db.update('dispersiones', d.id, {banco, cuenta: r.cuenta, clabe});
      invalidate('beneficiarios'); toast('Cuenta actualizada' + (pend.length ? ` (y ${pend.length} dispersión(es) pendientes)` : '')); refresh();
    }}], {width: '720px'});
  const cmp = () => { const b = bens.find(x => x.id === val('rb'));
    const f = (k, a, c) => `<div>${k}</div><div class="mono">${esc(a || '—')}</div><div class="mono ${a !== c ? 'chg' : ''}">${esc(c || '—')}</div>`;
    $('#rc').innerHTML = `<div class="cmp"><div class="h"></div><div class="h">Actual en BD</div><div class="h">Solicitado</div>${f('Banco', b && b.banco, Logica.normBanco(r.banco))}${f('Cuenta', b && b.cuenta, r.cuenta)}${f('CLABE', b && b.clabe, Logica.normalizarClabe(r.clabe))}${f('Correo', b && b.correo, r.correo)}</div>`; };
  $('#rb').onchange = cmp; cmp();
}
