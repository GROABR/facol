/* ---------- FACOL: cuestionario de alta (invitaciones), revisión de respuestas y contratos ---------- */
VIEWS.contratos = async (main, arg) => {
  const tab = (arg && arg.tab) || 'respuestas';
  const [inv, cons, url] = await Promise.all([db.list('invitaciones', {order: 'created_at', desc: true}), db.list('contratos', {order: 'created_at', desc: true}), urlFormularios()]);
  const bens = await cat('beneficiarios');
  const resp = inv.filter(i => i.estado === 'respondida');
  main.innerHTML = `<h2>Altas y contratos</h2>
    <div class="toolbar"><span class="sp"></span><button class="btn" id="cMan">Contrato manual (sin cuestionario)</button><button class="btn primary" id="cNew">+ Enviar cuestionario</button></div>
    ${url ? '' : '<div class="note">Configura en Ajustes la dirección donde publicaste los formularios para poder generar las ligas.</div>'}
    <div class="tabs">${[['respuestas', 'Por revisar', resp.length], ['invitaciones', 'Todas las invitaciones'], ['contratos', `Contratos (${cons.length})`]].map(([k, l, b]) => `<button data-t="${k}" class="${k === tab ? 'on' : ''}">${l}${b ? `<span class="badge">${b}</span>` : ''}</button>`).join('')}</div><div id="cBody"></div>`;
  $$('[data-t]').forEach(b => b.onclick = () => go('contratos', {tab: b.dataset.t}));
  $('#cNew').onclick = () => nuevaInvitacion(url);
  $('#cMan').onclick = async () => { const b = await buscarBeneficiario('', bens); if (b) generarContrato(b); };
  const body = $('#cBody');
  const lista = rows => `<div class="tablewrap"><table><thead><tr><th>Enviada</th><th>Para</th><th>Contacto</th><th>Operación</th><th>Respondida</th><th>Estado</th><th></th></tr></thead><tbody>
    ${rows.map(i => `<tr data-id="${i.id}"><td>${fmtDate(i.created_at)}</td><td><b>${esc(i.respuesta ? Logica.valoresContrato(i.respuesta).SOCIO : i.nombre || '')}</b></td><td>${esc(i.correo || '')}<div class="small muted">${esc(i.celular || '')}</div></td>
      <td>${esc(i.asesor || '')}</td><td>${i.respondida_at ? fmtDate(i.respondida_at) : '—'}</td><td>${tag(i.estado)}</td>
      <td style="white-space:nowrap">${i.estado === 'enviada' ? `<button class="link" data-mail="${i.id}">Enviar por correo</button> · <button class="link" data-link="${i.id}">Liga</button> · <button class="link" data-can="${i.id}" style="color:var(--bad)">Cancelar</button>` : ''}${i.respuesta ? `<button class="link" data-rev="${i.id}">${i.estado === 'respondida' ? 'Revisar' : 'Ver'}</button>` : ''} · <button class="link" data-borra="${i.id}" style="color:var(--bad)">Borrar</button></td></tr>`).join('') || '<tr><td colspan="7" class="empty">Nada por aquí</td></tr>'}</tbody></table></div>`;
  if (tab === 'respuestas') body.innerHTML = lista(resp);
  if (tab === 'invitaciones') body.innerHTML = lista(inv);
  if (tab === 'contratos') tabContratos(body, cons, bens, arg || {});
  $$('[data-link]').forEach(b => b.onclick = () => mostrarLiga(inv.find(i => i.id === b.dataset.link), url));
  $$('[data-can]').forEach(b => b.onclick = async () => { if (await confirmar('Cancelar invitación', 'La liga dejará de funcionar.', 'Cancelar invitación', 'bad')) { await db.update('invitaciones', b.dataset.can, {estado: 'cancelada'}); refresh(); } });
  $$('[data-rev]').forEach(b => b.onclick = () => revisarAlta(inv.find(i => i.id === b.dataset.rev)));
  $$('[data-mail]').forEach(b => b.onclick = () => pedirCorreoYEnviar(inv.find(i => i.id === b.dataset.mail), url));
  $$('[data-borra]').forEach(b => b.onclick = () => borrarInvitacion(inv.find(i => i.id === b.dataset.borra)));
};

/* ===================== lista de contratos: buscador, orden y carga de contratos existentes ===================== */
const extDe = p => (String(p).match(/\.(\w+)$/) || [, 'docx'])[1].toLowerCase();
function tabContratos(body, cons, bens, arg) {
  const st = {q: arg.q || '', orden: arg.orden || 'az'};
  body.innerHTML = `<div class="toolbar"><input id="kq" placeholder="Buscar por nombre, RFC o CURP" style="width:300px" value="${esc(st.q)}">
      <select id="ko">${[['az', 'Nombre A → Z'], ['za', 'Nombre Z → A'], ['rec', 'Más recientes'], ['ant', 'Más antiguos']].map(([k, l]) => `<option value="${k}" ${k === st.orden ? 'selected' : ''}>${l}</option>`).join('')}</select>
      <span class="sp"></span><span class="muted small" id="kn"></span><button class="btn primary" id="kUp">Subir contratos que ya tengo</button></div>
    <div class="tablewrap tall"><table><thead><tr><th>Socio</th><th>Fecha inicio</th><th>Origen</th><th>Registrado</th><th></th></tr></thead><tbody id="kTb"></tbody></table></div>`;
  const filas = cons.map(c => ({c, b: bens.find(x => x.id === c.beneficiario_id)})).map(x => ({...x, nombre: x.b ? x.b.nombre : '(beneficiario borrado)'}));
  const draw = () => {
    const q = Logica.normName(val('kq')), o = val('ko');
    const rows = filas.filter(x => !q || Logica.normName([x.nombre, x.b && x.b.rfc, x.b && x.b.curp].join(' ')).includes(q));
    const cmp = {az: (a, b) => a.nombre.localeCompare(b.nombre, 'es'), za: (a, b) => b.nombre.localeCompare(a.nombre, 'es'),
      rec: (a, b) => String(b.c.created_at).localeCompare(String(a.c.created_at)), ant: (a, b) => String(a.c.created_at).localeCompare(String(b.c.created_at))}[o];
    rows.sort(cmp);
    $('#kn').textContent = `${rows.length} de ${filas.length} contrato(s)`;
    $('#kTb').innerHTML = rows.map(x => `<tr><td><b>${esc(x.nombre)}</b>${x.b && x.b.rfc ? `<div class="small muted">${esc(x.b.rfc)}</div>` : ''}</td><td>${fmtDay(x.c.fecha_inicio)}</td>
      <td>${/\/subido_/.test(x.c.path) ? tag('info', 'subido ' + extDe(x.c.path).toUpperCase()) : tag('ok', 'generado')}</td><td>${fmtDate(x.c.created_at)}<div class="small muted">${esc(x.c.created_by || '')}</div></td>
      <td style="white-space:nowrap"><button class="link" data-dl="${x.c.id}">Descargar</button> · <button class="link" data-del="${x.c.id}" style="color:var(--bad)">Quitar</button></td></tr>`).join('')
      || `<tr><td colspan="5" class="empty">${filas.length ? 'Sin resultados' : 'Sin contratos. Puedes subir los que ya tienes con "Subir contratos que ya tengo".'}</td></tr>`;
    $$('[data-dl]', body).forEach(b => b.onclick = () => { const x = filas.find(f => f.c.id === b.dataset.dl); guard(() => descargarArchivo(x.c.path, `${x.nombre}.${extDe(x.c.path)}`)); });
    $$('[data-del]', body).forEach(b => b.onclick = async () => { const x = filas.find(f => f.c.id === b.dataset.del);
      if (!(await confirmar('Quitar contrato', `¿Quitar el contrato de ${esc(x.nombre)}? El beneficiario no se borra.`, 'Quitar', 'bad'))) return;
      await guard(async () => { await db.removeFile(x.c.path).catch(() => {}); await db.remove('contratos', x.c.id); }, 'Contrato quitado'); go('contratos', {tab: 'contratos', q: val('kq'), orden: val('ko')}); });
  };
  $('#kq').oninput = draw; $('#ko').onchange = draw; draw();
  $('#kUp').onclick = async () => { const fs = await pickFiles('.docx,.pdf,.doc', true); if (fs.length) subirContratos(fs); };
}

/** Saca de un nombre de archivo el nombre de la persona: "CONTRATO PACHECO JUAREZ ONOFRE (1).docx" -> "PACHECO JUAREZ ONOFRE" */
const nombreDeArchivo = n => String(n).replace(/\.[^.]+$/, '').replace(/[_\-.]+/g, ' ').replace(/\(\d+\)/g, ' ')
  .replace(/\b(CONTRATO|CONTRATOS|FIRMADO|FIRMA|SCU|SOCIO|CARTA|CONFIDENCIALIDAD|COPIA|FINAL|V\d+)\b/gi, ' ').replace(/\d+/g, ' ').replace(/\s+/g, ' ').trim();
/** Para .docx hechos con el machote: lee "EL (LA) C. NOMBRE QUIEN…" del documento. */
async function nombreDeDocx(file) {
  try { const z = await JSZip.loadAsync(await file.arrayBuffer()); const x = await z.file('word/document.xml').async('string');
    const t = x.replace(/<\/w:p>/g, '\n').replace(/<[^>]+>/g, ''); const m = t.match(/C\.\s+([A-ZÁÉÍÓÚÑÜ ]{6,80}?)\s+QUIEN/); return m ? m[1].trim() : ''; } catch { return ''; }
}

async function subirContratos(files) {
  const bens = await cat('beneficiarios', true); const idx = Logica.indexBens(bens);
  const items = [];
  for (const f of files) {
    let base = nombreDeArchivo(f.name); let m = Logica.matchNombre(base, bens, idx);
    if (m.tipo !== 'exacto' && /\.docx$/i.test(f.name)) { const dn = await nombreDeDocx(f); if (dn) { const m2 = Logica.matchNombre(dn, bens, idx); if (m2.score >= m.score) { m = m2; base = dn; } } }
    items.push({f, base, m, benId: m.tipo === 'ninguno' ? null : m.ben.id, confirmado: m.tipo === 'exacto'});
  }
  const draw = () => {
    const sinB = items.filter(i => !i.benId).length, porC = items.filter(i => i.benId && !i.confirmado).length;
    $('#scB').innerHTML = `<p class="small" style="margin-top:0">${tag('ok', `${items.length - sinB - porC} ligados`)} ${porC ? tag('warn', `${porC} por parecido`) : ''} ${sinB ? tag('bad', `${sinB} sin beneficiario (no se suben)`) : ''}</p>
      <div class="tablewrap" style="max-height:52vh"><table><thead><tr><th>Archivo</th><th>Beneficiario</th><th></th></tr></thead><tbody>
      ${items.map((it, i) => { const opts = [...new Map([...(it.m.candidatos || []), ...(it.benId ? [bens.find(b => b.id === it.benId)] : [])].filter(Boolean).map(b => [b.id, b])).values()];
        return `<tr><td>📄 ${esc(it.f.name)}<div class="small muted">leído como: ${esc(it.base || '—')}</div></td>
          <td><select data-i="${i}" class="sb" style="max-width:300px"><option value="">— no subir —</option>${opts.map(b => `<option value="${b.id}" ${b.id === it.benId ? 'selected' : ''}>${esc(b.nombre)}</option>`).join('')}<option value="__buscar">Buscar otro…</option></select></td>
          <td>${!it.benId ? tag('bad', 'no encontrado') : it.confirmado ? tag('ok', 'ok') : tag('warn', `parecido ${Math.round(it.m.score * 100)}%`)}</td></tr>`; }).join('')}
      </tbody></table></div>`;
    $$('#scB .sb').forEach(s => s.onchange = async () => { const it = items[+s.dataset.i];
      if (s.value === '__buscar') { const b = await buscarBeneficiario(it.base, bens); reabrir(); if (b) { it.benId = b.id; it.m.candidatos = [b, ...(it.m.candidatos || [])]; it.confirmado = true; } return draw(); }
      it.benId = s.value || null; it.confirmado = true; draw(); });
  };
  const reabrir = () => {
    modal(`Subir contratos (${files.length})`, `<div class="form" style="margin-bottom:8px"><label>Fecha de inicio (opcional)</label><input type="date" id="scF"></div><div id="scB"></div>`, [
      {label: 'Cancelar'},
      {label: 'Guardar contratos', cls: 'primary', onClick: async () => {
        const ok = items.filter(i => i.benId); if (!ok.length) { toast('No hay contratos ligados a un beneficiario'); return false; }
        const fecha = val('scF') || null; let n = 0;
        for (const it of ok) {
          const path = `beneficiarios/${it.benId}/contrato/subido_${Date.now()}_${safePath(it.f.name)}`;
          await db.upload(path, it.f); await db.insert('contratos', {beneficiario_id: it.benId, fecha_inicio: fecha, path}); n++;
        }
        toast(`${n} contrato(s) guardado(s)${items.length - n ? `; ${items.length - n} sin beneficiario no se subieron` : ''}`); go('contratos', {tab: 'contratos'});
      }}], {width: '860px'});
    draw();
  };
  reabrir();
}

/* ----- envío del cuestionario por correo ----- */
const linkAlta = (i, url) => `${url}/alta.html?t=${i.token}`;
function correoAlta(i, url) {
  const nombre = String(i.nombre || '').trim();
  const link = linkAlta(i, url);
  const firma = esDisp(modulo) ? nombreModulo(modulo) : 'Colorines';   // firma del correo: el cliente del módulo (Colorines)
  return {to: i.correo, subject: `${firma} · Completa tu registro de alta`,
    html: `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:15px;color:#1d2733;max-width:560px">
      <p>Hola${nombre ? ' ' + esc(nombre) : ''}:</p>
      <p>Para darte de alta y preparar tu contrato, necesitamos que llenes el siguiente cuestionario con tus datos tal como aparecen en tu INE.
      Al final te pediremos tu firma, una foto de tu huella, de tu nombre escrito a mano y de tu INE (por ambos lados).</p>
      <p style="margin:26px 0"><a href="${link}" style="background:#1f5fbf;color:#fff;text-decoration:none;padding:12px 22px;border-radius:8px;font-weight:600">Llenar cuestionario</a></p>
      <p style="font-size:13px;color:#6b7785">Si el botón no abre, copia esta liga en tu navegador:<br>${link}</p>
      <p style="font-size:13px;color:#6b7785">La liga es personal; no la compartas. Cualquier duda, responde a este correo.</p>
      <p>Saludos,<br><b>${esc(firma)}</b></p></div>`};
}
/** Envía el correo con Outlook (app de escritorio). Si no se puede, abre el correo ya redactado. */
async function enviarCorreoAlta(i, url) {
  if (!i.correo) throw new Error('Falta el correo de la persona');
  const msg = correoAlta(i, url);
  if (window.desk && desk.enviarCorreo) { const r = await desk.enviarCorreo(msg); if (r && r.ok) return 'enviado'; }
  const texto = `Hola, para darte de alta y preparar tu contrato llena este cuestionario con tus datos: ${linkAlta(i, url)}`;
  window.open(`mailto:${encodeURIComponent(i.correo)}?subject=${encodeURIComponent(msg.subject)}&body=${encodeURIComponent(texto)}`);
  return 'redactado';
}
/** Correo con la liga del formulario de cambio de cuenta (es la misma liga para todos). */
async function enviarCorreoCambio(correo, nombre, url) {
  const firma = esDisp(modulo) ? nombreModulo(modulo) : 'Colorines'; const link = `${url}/cambio.html`;
  const msg = {to: correo, subject: `${firma} · Actualiza tu cuenta para pagos`,
    html: `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:15px;color:#1d2733;max-width:560px">
      <p>Hola${nombre ? ' ' + esc(nombre) : ''}:</p>
      <p>Si necesitas cambiar la cuenta bancaria donde recibes tus pagos, llena este formulario. Verifica muy bien tu número de cuenta y tu CLABE (18 dígitos).</p>
      <p style="margin:26px 0"><a href="${link}" style="background:#1f5fbf;color:#fff;text-decoration:none;padding:12px 22px;border-radius:8px;font-weight:600">Cambiar mi cuenta</a></p>
      <p style="font-size:13px;color:#6b7785">Si el botón no abre, copia esta liga en tu navegador:<br>${link}</p>
      <p style="font-size:13px;color:#6b7785">El cambio se aplica hasta que lo revisemos; mientras tanto tus pagos siguen a tu cuenta actual.</p>
      <p>Saludos,<br><b>${esc(firma)}</b></p></div>`};
  if (window.desk && desk.enviarCorreo) { const r = await desk.enviarCorreo(msg); if (r && r.ok) return 'enviado'; }
  window.open(`mailto:${encodeURIComponent(correo)}?subject=${encodeURIComponent(msg.subject)}&body=${encodeURIComponent('Para cambiar la cuenta donde recibes tus pagos llena este formulario: ' + link)}`);
  return 'redactado';
}
const avisoEnvio = (r, n = 1) => toast(r === 'enviado' ? `Cuestionario enviado por correo${n > 1 ? ` a ${n} personas` : ''}` : 'Se abrió tu correo con el mensaje listo: solo presiona Enviar');

function nuevaInvitacion(url) {
  modal('Enviar cuestionario de alta', `<div class="form"><label>Nombre</label><input id="iN" placeholder="Nombre de la persona"><label>Correo</label><input id="iC" type="email" placeholder="a donde se envía"><label>Celular (opcional)</label><input id="iT" placeholder="10 dígitos"></div>
    <p class="small muted">Se le envía por correo una liga personal para llenar el cuestionario. Cuando lo conteste aparecerá en "Por revisar".</p><div class="err" id="iErr"></div>`, [
    {label: 'Cancelar'},
    {label: 'Solo crear liga', onClick: async () => {
      const i = await db.insert('invitaciones', {nombre: val('iN').toUpperCase(), correo: val('iC').toLowerCase(), celular: val('iT').replace(/\D/g, '')});
      setTimeout(() => mostrarLiga(i, url), 50);
    }},
    {label: 'Enviar por correo', cls: 'primary', onClick: async () => {
      const correo = val('iC').toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) { $('#iErr').textContent = 'Escribe un correo válido'; return false; }
      const i = await db.insert('invitaciones', {nombre: val('iN').toUpperCase(), correo, celular: val('iT').replace(/\D/g, '')});
      avisoEnvio(await enviarCorreoAlta(i, url)); refresh();
    }}]);
}
/** Pide (o confirma) el correo de una invitación ya creada y la envía. */
function pedirCorreoYEnviar(i, url) {
  modal(`Enviar cuestionario a ${i.nombre || ''}`, `<div class="form"><label>Correo</label><input id="pcC" type="email" value="${esc(i.correo || '')}"></div><div class="err" id="pcE"></div>`, [
    {label: 'Cancelar'},
    {label: 'Enviar por correo', cls: 'primary', onClick: async () => {
      const correo = val('pcC').toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) { $('#pcE').textContent = 'Escribe un correo válido'; return false; }
      if (correo !== i.correo) { await db.update('invitaciones', i.id, {correo}); i.correo = correo; }
      avisoEnvio(await enviarCorreoAlta(i, url)); refresh();
    }}]);
}
async function borrarInvitacion(i) {
  const nombre = i.respuesta ? Logica.valoresContrato(i.respuesta).SOCIO : i.nombre || i.correo || '';
  if (!(await confirmar('Borrar invitación', `¿Borrar la invitación de <b>${esc(nombre)}</b>${i.respuesta ? ' y sus respuestas' : ''}? La liga deja de funcionar.${i.beneficiario_id ? '<br><span class="small muted">La persona sigue en Beneficiarios; para quitarla bórrala desde ahí.</span>' : ''}`, 'Borrar', 'bad'))) return;
  await guard(async () => {
    for (const p of Object.values((i.respuesta || {})._archivos || {})) if (p && !String(p).startsWith('data:')) await db.removeFile(p).catch(() => {});
    await db.remove('invitaciones', i.id);
  }, 'Invitación borrada'); refresh();
}
function mostrarLiga(i, url) {
  const link = `${url || '[configura la dirección en Ajustes]'}/alta.html?t=${i.token}`;
  const msg = `Hola${i.nombre ? ' ' + i.nombre.split(' ')[0] : ''}, por favor llena este cuestionario para tu alta: ${link}`;
  modal('Liga del cuestionario', `<p>Liga para <b>${esc(i.nombre || i.correo || '')}</b>:</p><input readonly id="lk" style="width:100%" value="${esc(link)}">
    <p class="small muted">La liga solo sirve para una persona y deja de funcionar al aprobar su alta o si la cancelas.</p>`, [
    {label: 'Enviar por correo', onClick: () => { setTimeout(() => pedirCorreoYEnviar(i, url), 50); }},
    ...(i.celular ? [{label: 'Abrir WhatsApp', onClick: () => { window.open(`https://wa.me/52${i.celular}?text=${encodeURIComponent(msg)}`); return false; }}] : []),
    {label: 'Copiar mensaje', onClick: () => { navigator.clipboard.writeText(msg); toast('Mensaje copiado'); return false; }},
    {label: 'Copiar liga', cls: 'primary', onClick: () => { navigator.clipboard.writeText(link); toast('Liga copiada'); refresh(); }}]);
}

async function imagenDe(path) {
  if (!path) return null;
  try { const blob = await db.download(path); const data = new Uint8Array(await blob.arrayBuffer());
    const ext = /png/i.test(blob.type || path) ? 'png' : 'jpeg';
    const dim = await new Promise(res => { const im = new Image(); im.onload = () => res({w: im.naturalWidth, h: im.naturalHeight}); im.onerror = () => res({w: 1, h: 1}); im.src = URL.createObjectURL(blob); });
    return {data, ext, ...dim, url: URL.createObjectURL(blob)};
  } catch { return null; }
}

async function revisarAlta(inv) {
  const d = inv.respuesta || {}; const archivos = d._archivos || {};
  const bens = await cat('beneficiarios', true);
  const v = Logica.valoresContrato(d);
  const existente = bens.find(b => (d.curp && b.curp === String(d.curp).toUpperCase()) || (d.rfc && b.rfc === String(d.rfc).toUpperCase())) || (() => { const m = Logica.matchNombre(v.SOCIO, bens); return m.tipo === 'exacto' ? m.ben : null; })();
  const errores = Logica.validarAlta(d);
  const imgs = {}; for (const [k, p] of Object.entries(archivos)) imgs[k] = await imagenDe(p);
  const cl = Logica.validarClabe(d.clabe, d.banco);
  modal('Alta: ' + v.SOCIO, `
    ${existente ? `<div class="note">Ya existe en la BD como <b>${esc(existente.nombre)}</b>. Al aprobar se actualizan sus datos (la cuenta anterior queda en el historial).</div>` : '<div class="note info">Es una persona nueva: al aprobar se da de alta en la BD de beneficiarios.</div>'}
    ${Object.keys(errores).length ? `<div class="note">Revisa: ${Object.entries(errores).map(([k, e]) => `${esc((Logica.CAMPOS_ALTA.find(c => c.k === k) || {}).l || k)}: ${esc(e)}`).join(' · ')}</div>` : ''}
    <div class="kv">${Logica.CAMPOS_ALTA.filter(c => c.k && !['file', 'firma'].includes(c.t)).map(c => `<div>${esc(c.l)}</div><div>${esc(c.t === 'date' ? fmtDay(d[c.k]) : d[c.k] || '—')}${c.k === 'clabe' ? ' ' + (cl.ok ? tag('ok', 'válida') + (cl.aviso ? ' ' + tag('warn', cl.aviso) : '') : tag('bad', cl.motivo)) : ''}</div>`).join('')}</div>
    <h3>Documentos</h3><div class="grid">${Logica.CAMPOS_ALTA.filter(c => ['file', 'firma'].includes(c.t)).map(c => `<div class="card"><div class="k">${esc(c.l)}</div>${imgs[c.k] ? `<img class="prev" src="${imgs[c.k].url}">` : archivos[c.k] ? `<button class="link" data-f="${esc(archivos[c.k])}">Descargar</button>` : '<span class="muted">—</span>'}</div>`).join('')}</div>`,
    inv.estado !== 'respondida' ? [{label: 'Generar contrato', onClick: async () => { const b = bens.find(x => x.id === inv.beneficiario_id); if (b) generarContrato(b, inv); else toast('No está ligada a un beneficiario'); return false; }}, {label: 'Cerrar', cls: 'primary'}] : [
      {label: 'Pedir corrección', onClick: async () => { await db.update('invitaciones', inv.id, {estado: 'enviada'}); toast('La liga vuelve a quedar abierta para que corrija'); refresh(); }},
      {label: 'Cancelar'},
      {label: 'Aprobar y dar de alta', cls: 'primary', onClick: async () => {
        if (Object.keys(errores).length && !(await confirmar('Hay datos con error', 'Hay campos con errores. ¿Aprobar de todos modos?'))) return false;
        const row = {nombre: v.SOCIO, apellido_paterno: v['APELLIDO PATERNO'], apellido_materno: v['APELLIDO MATERNO'], nombres: v.NOMBRES,
          banco: Logica.normBanco(d.banco) || cl.banco || null, cuenta: Logica.normalizarCuenta(d.numero_cuenta) || null, clabe: Logica.normalizarClabe(d.clabe) || null,
          rfc: v.RFC, curp: v.CURP, correo: d.correo, celular: d.celular, datos: {...d}, updated_at: new Date().toISOString()};
        let id;
        if (existente) {
          if (existente.clabe !== row.clabe || existente.cuenta !== row.cuenta) await db.insert('beneficiario_cuentas_hist', {beneficiario_id: existente.id, banco: existente.banco, cuenta: existente.cuenta, clabe: existente.clabe, origen: 'alta por cuestionario (anterior)'});
          await db.update('beneficiarios', existente.id, row); id = existente.id;
        } else { id = (await db.insert('beneficiarios', row)).id; await db.insert('beneficiario_cuentas_hist', {beneficiario_id: id, banco: row.banco, cuenta: row.cuenta, clabe: row.clabe, origen: 'alta por cuestionario'}); }
        await db.update('invitaciones', inv.id, {estado: 'aprobada', beneficiario_id: id});
        // liga sus dispersiones pendientes (las que no estaban en la BD o ya eran suyas) con la cuenta aprobada
        const pend = (await db.list('dispersiones', {eq: {estado: 'pendiente'}})).filter(d => d.beneficiario_id === id || (!d.beneficiario_id && Logica.nameScore(d.nombre_archivo, row.nombre) >= 0.86));
        for (const d of pend) await db.update('dispersiones', d.id, {beneficiario_id: id, nombre: row.nombre, banco: row.banco, cuenta: row.cuenta, clabe: row.clabe});
        invalidate('beneficiarios'); toast('Alta aprobada');
        const b = await db.get('beneficiarios', id);
        setTimeout(() => generarContrato(b, {...inv, beneficiario_id: id}), 50);
      }}], {width: '920px'});
  $$('[data-f]').forEach(a => a.onclick = () => guard(() => descargarArchivo(a.dataset.f, a.dataset.f.split('/').pop())));
}

/** Llena el machote con los datos del beneficiario (y de su cuestionario si existe). */
async function generarContrato(b, inv) {
  const d = {...(b.datos || {})};
  // si se dio de alta por importación, completa con lo que haya en la ficha
  d.apellido_paterno = d.apellido_paterno || b.apellido_paterno || ''; d.apellido_materno = d.apellido_materno || b.apellido_materno || ''; d.nombres = d.nombres || b.nombres || '';
  d.rfc = d.rfc || b.rfc || ''; d.curp = d.curp || b.curp || ''; d.celular = d.celular || b.celular || ''; d.correo = d.correo || b.correo || '';
  d.banco = b.banco || d.banco || ''; d.numero_cuenta = b.cuenta || d.numero_cuenta || ''; d.clabe = b.clabe || d.clabe || '';
  const tpl = await machote(); const campos = await Docx.campos(tpl);
  const vals0 = Logica.valoresContrato(d, today());
  if (!d.apellido_paterno && !d.nombres) vals0.SOCIO = b.nombre;
  const imgCampos = campos.filter(c => Docx.IMG_KEYS.includes(c.toUpperCase()));
  const txt = campos.filter(c => !imgCampos.includes(c));
  const arch = d._archivos || {}; const lugares = await Docx.lugaresImagenes(tpl).catch(() => ({firmas: 0, huellas: 0}));
  modal('Generar contrato: ' + (vals0.SOCIO || b.nombre), `<p class="small muted" style="margin-top:0">Revisa o completa los campos del machote. Lo que cambies aquí solo afecta este contrato.</p>
    <div class="form">${txt.map(c => `<label>[${esc(c)}]</label>` + (c === 'FECHA INICIO' ? `<input type="date" data-c="${esc(c)}" value="${today()}">` : `<input data-c="${esc(c)}" value="${esc(vals0[c] ?? '')}" class="${vals0[c] ? '' : 'bad'}">`)).join('')}</div>
    <p class="small">Firma: ${arch.firma ? tag('ok', `se pone en ${lugares.firmas} lugar(es) del machote`) : tag('warn', 'no hay firma del cuestionario')}
      · Huella: ${arch.huella ? tag('ok', `se pone en ${lugares.huellas} lugar(es)`) : tag('warn', 'no hay huella del cuestionario')}</p>`, [
    {label: 'Cancelar'}, {label: 'Generar .docx', cls: 'primary', onClick: async () => {
      const vals = {...vals0}; $$('#dlg [data-c]').forEach(i => vals[i.dataset.c] = i.type === 'date' ? fmtDay(i.value) : i.value.trim());
      const fechaIni = $('#dlg [data-c="FECHA INICIO"]') ? $('#dlg [data-c="FECHA INICIO"]').value : today();
      if (campos.includes('EDAD') && d.fecha_nacimiento) vals.EDAD = String(Logica.edad(d.fecha_nacimiento, new Date(fechaIni + 'T12:00:00')));
      // firma y huella siempre (el machote detecta dónde van); "nombre escrito a mano" solo si el machote lo pide
      const imagenes = {};
      const firma = await imagenDe(arch.firma); if (firma) imagenes.FIRMA = {...firma, maxW: 1.7};
      const huella = await imagenDe(arch.huella); if (huella) imagenes.HUELLA = {...huella, maxW: 0.85};
      if (imgCampos.some(c => /NOMBRE/i.test(c))) { const nm = await imagenDe(arch.nombre_mano); if (nm) for (const c of imgCampos.filter(c => /NOMBRE/i.test(c))) imagenes[c.toUpperCase()] = {...nm, maxW: 2}; }
      const out = await Docx.llenar(tpl, vals, imagenes);
      const nombre = `${vals.SOCIO || b.nombre}.docx`;
      const path = `beneficiarios/${b.id}/contrato/${Date.now()}_${safePath(nombre)}`;
      await db.upload(path, new Blob([out], {type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}));
      await db.insert('contratos', {beneficiario_id: b.id, invitacion_id: inv ? inv.id : null, fecha_inicio: fechaIni, path});
      if (inv) await db.update('invitaciones', inv.id, {estado: 'contrato'});
      const p = await guardarBytes(out, nombre);
      if (window.desk && p) { toast('Creando PDF con Word…'); const pdf = await desk.docxToPdf(p); toast(pdf ? 'Contrato guardado en Word y PDF' : 'Se guardó el Word; no se pudo crear el PDF (¿Word instalado?)'); }
      if (['contratos', 'operacion'].includes(view)) refresh();
    }}], {width: '760px'});
}
