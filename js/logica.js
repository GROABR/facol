/* ---------- FACOL: lógica de negocio pura (sin pantalla ni base de datos) ----------
   Se usa en la app, en los formularios públicos y en las pruebas (node tools/test-logica.js). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Logica = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const r2 = n => Math.round((+n || 0) * 100 + Number.EPSILON * 100) / 100;

  /* ----- nombres ----- */
  const normName = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  const tokens = s => normName(s).split(' ').filter(Boolean);
  const tokenKey = s => tokens(s).sort().join(' ');

  // distancia de edición pequeña para errores de dedo (MARAGRITA vs MARGARITA)
  function lev(a, b) {
    if (a === b) return 0; const m = a.length, n = b.length; if (!m || !n) return m || n;
    let prev = Array.from({length: n + 1}, (_, j) => j);
    for (let i = 1; i <= m; i++) {
      const cur = [i];
      for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
    return prev[n];
  }
  const tokSim = (a, b) => 1 - lev(a, b) / Math.max(a.length, b.length);

  // similitud entre dos nombres sin importar el orden (apellidos primero o al final)
  function nameScore(a, b) {
    const A = tokens(a), B = tokens(b); if (!A.length || !B.length) return 0;
    const used = new Set(); let sum = 0;
    for (const t of A) {
      let best = 0, bi = -1;
      B.forEach((u, i) => { if (used.has(i)) return; const s = tokSim(t, u); if (s > best) { best = s; bi = i; } });
      if (bi >= 0 && best >= 0.75) { used.add(bi); sum += best; }
    }
    return sum / Math.max(A.length, B.length);
  }

  /** Busca un nombre en la base. Devuelve {tipo: 'exacto'|'sugerido'|'ninguno', ben, score, candidatos} */
  function matchNombre(nombre, bens, index) {
    const key = tokenKey(nombre);
    const idx = index || indexBens(bens);
    const exact = idx.get(key);
    if (exact && exact.length === 1) return {tipo: 'exacto', ben: exact[0], score: 1, candidatos: exact};
    const cands = bens.map(b => ({b, s: Math.max(nameScore(nombre, b.nombre), b.nombre_alt ? nameScore(nombre, b.nombre_alt) : 0)}))
      .filter(x => x.s >= 0.6).sort((x, y) => y.s - x.s).slice(0, 5);
    if (exact && exact.length > 1) return {tipo: 'sugerido', ben: exact[0], score: 1, candidatos: exact, duplicado: true};
    if (cands.length && cands[0].s >= 0.86 && (cands.length === 1 || cands[0].s - cands[1].s > 0.08))
      return {tipo: 'sugerido', ben: cands[0].b, score: cands[0].s, candidatos: cands.map(c => c.b)};
    return {tipo: 'ninguno', ben: cands[0] ? cands[0].b : null, score: cands[0] ? cands[0].s : 0, candidatos: cands.map(c => c.b)};
  }
  function indexBens(bens) {
    const m = new Map();
    for (const b of bens) for (const n of [b.nombre, b.nombre_alt]) {
      if (!n) continue; const k = tokenKey(n); if (!k) continue;
      const arr = m.get(k) || []; if (!arr.includes(b)) arr.push(b); m.set(k, arr);
    }
    return m;
  }

  /* ----- CLABE / cuentas ----- */
  const BANCOS = {
    '002': 'BANAMEX', '006': 'BANCOMEXT', '012': 'BBVA', '014': 'SANTANDER', '019': 'BANJERCITO', '021': 'HSBC',
    '030': 'BAJIO', '036': 'INBURSA', '042': 'MIFEL', '044': 'SCOTIABANK', '058': 'BANREGIO', '059': 'INVEX',
    '060': 'BANSI', '062': 'AFIRME', '072': 'BANORTE', '106': 'BANK OF AMERICA', '112': 'BMONEX', '113': 'VE POR MAS',
    '127': 'BANCO AZTECA', '130': 'COMPARTAMOS', '133': 'ACTINVER', '136': 'INTERCAM', '137': 'BANCOPPEL',
    '143': 'CIBANCO', '152': 'BANCREA', '156': 'SABADELL', '166': 'BANCO DEL BIENESTAR', '167': 'HEY BANCO',
    '638': 'NU MEXICO', '646': 'STP', '722': 'MERCADO PAGO',
  };
  const ALIAS_BANCO = {BANCOMER: 'BBVA', 'BBVA BANCOMER': 'BBVA', CITIBANAMEX: 'BANAMEX', 'BANCO DEL BAJIO': 'BAJIO', NU: 'NU MEXICO', 'NU BANK': 'NU MEXICO', AZTECA: 'BANCO AZTECA', COPPEL: 'BANCOPPEL', HEY: 'HEY BANCO'};
  const normBanco = s => { const n = normName(s); return ALIAS_BANCO[n] || n; };
  const soloDigitos = v => String(v ?? '').replace(/^'/, '').replace(/\D/g, '');

  /** Convierte lo que venga de Excel (número, texto con ', con espacios) a CLABE de 18 dígitos.
      Excel suele borrar el 0 inicial (17 dígitos) -> se repone. */
  function normalizarClabe(v) {
    if (v === null || v === undefined || v === '') return '';
    let s = typeof v === 'number' ? BigInt(Math.round(v)).toString() : soloDigitos(v);
    if (s.length === 17) s = '0' + s;
    return s;
  }
  function digitoControlClabe(c17) {
    const w = [3, 7, 1]; let sum = 0;
    for (let i = 0; i < 17; i++) sum += (+c17[i] * w[i % 3]) % 10;
    return (10 - (sum % 10)) % 10;
  }
  /** {ok, motivo, banco} */
  function validarClabe(v, bancoDeclarado) {
    const c = normalizarClabe(v);
    if (!c) return {ok: false, motivo: 'Sin CLABE'};
    if (!/^\d{18}$/.test(c)) return {ok: false, motivo: `La CLABE debe tener 18 dígitos (tiene ${c.length})`};
    if (digitoControlClabe(c.slice(0, 17)) !== +c[17]) return {ok: false, motivo: 'Dígito verificador incorrecto (CLABE mal capturada)'};
    const banco = BANCOS[c.slice(0, 3)] || null;
    if (bancoDeclarado && banco && normBanco(bancoDeclarado) && normBanco(bancoDeclarado) !== banco)
      return {ok: true, banco, aviso: `La CLABE es de ${banco} pero dice ${String(bancoDeclarado).trim()}`};
    return {ok: true, banco};
  }
  const bancoDeClabe = v => BANCOS[normalizarClabe(v).slice(0, 3)] || '';
  const normalizarCuenta = v => typeof v === 'number' ? BigInt(Math.round(v)).toString() : soloDigitos(v);

  /* ----- comisiones ----- */
  /** lineas: [{id, pct, base_linea_id}] -> mismas líneas con monto. Las de base_linea_id son % sobre su línea padre. */
  function calcComisiones(base, lineas) {
    const byId = new Map(lineas.map(l => [l.id, {...l}]));
    const monto = (l, depth = 0) => {
      if (l._m !== undefined) return l._m;
      if (depth > 10) return 0;
      const p = l.base_linea_id && byId.get(l.base_linea_id);
      l._m = r2((p ? monto(p, depth + 1) : +base) * (+l.pct || 0) / 100);
      return l._m;
    };
    return lineas.map(l0 => { const l = byId.get(l0.id); const m = monto(l); const {_m, ...rest} = l; return {...rest, monto: m}; });
  }
  /** "Cuadre casa": los centavos de redondeo entre la comisión cobrada y la suma de líneas se suman a la línea CASA. */
  function ajustarCuadre(lineas, objetivo) {
    if (objetivo === undefined || objetivo === null) return lineas;
    const top = lineas.filter(l => !l.base_linea_id); if (!top.length) return lineas;
    const diff = r2(objetivo - top.reduce((s, l) => s + l.monto, 0));
    if (!diff || Math.abs(diff) > 1) return lineas;
    const casa = top.find(l => l.clase === 'casa') || top.reduce((a, b) => (b.monto > a.monto ? b : a));
    return lineas.map(l => l === casa ? {...l, monto: r2(l.monto + diff)} : l);
  }
  /** Esquema fijo de Colorines: CASA se queda con lo que resta del % total después de las terciarias;
      ABEL es un % de CASA (sale de CASA). */
  // Cada promotor tiene su % y sobre qué se calcula: 'total' (la comisión total), 'casa' o el nombre de otro promotor.
  const ESQUEMA_COLORINES = {pct: 11.5, casa: {nombre: 'CASA', tipo_pago: 'CASA'}, lineas: [
    {nombre: 'ABEL ORTIZ', pct: 35, base: 'casa', tipo_pago: 'TRANSFERENCIA'},
    {nombre: 'ARTURO', pct: 1.4, base: 'total', tipo_pago: 'EFECTIVO'},
    {nombre: 'OSCAR HERNANDEZ', pct: 1.6, base: 'total', tipo_pago: 'EFECTIVO'},
    {nombre: 'MANUEL', pct: 1, base: 'total', tipo_pago: 'TRANSFERENCIA/EFECTIVO'}]};
  /** Acepta el formato anterior (terciarias + abel) y lo pasa al nuevo (lineas con base). */
  function normalizarEsquema(e) {
    if (!e || !e.pct) return JSON.parse(JSON.stringify(ESQUEMA_COLORINES));
    if (Array.isArray(e.lineas)) return {...e, casa: typeof e.casa === 'string' ? {nombre: e.casa, id: e.casa_id} : (e.casa || {nombre: 'CASA'})};
    const lineas = [];
    if (e.abel && e.abel.nombre) lineas.push({nombre: e.abel.nombre, pct: +e.abel.pct || 0, base: 'casa', id: e.abel.id});
    for (const t of e.terciarias || []) lineas.push({nombre: t.nombre, pct: +t.pct || 0, base: 'total', id: t.id});
    return {pct: +e.pct, casa: {nombre: typeof e.casa === 'string' ? e.casa : 'CASA', id: e.casa_id}, lineas};
  }
  /** Líneas de comisión del esquema. CASA = % total − lo que va "sobre el total"; las demás salen de su base. */
  function lineasColorines(e0 = ESQUEMA_COLORINES) {
    const e = normalizarEsquema(e0); const ls = (e.lineas || []).filter(l => l.nombre);
    const enTotal = l => !l.base || l.base === 'total';
    const casaPct = Math.round((+e.pct - ls.filter(enTotal).reduce((s, l) => s + (+l.pct || 0), 0)) * 10000) / 10000;
    const casa = e.casa || {nombre: 'CASA'};
    const L = [{id: 'casa', nombre: casa.nombre || 'CASA', concepto: 'COMISIÓN CASA', clase: 'casa', pct: casaPct, comisionista_id: casa.id || null, tipo_pago: casa.tipo_pago || null}];
    ls.forEach((l, i) => L.push({id: 'l' + i, nombre: l.nombre, clase: enTotal(l) ? 'terciaria' : 'promotor', pct: +l.pct || 0, comisionista_id: l.id || null, tipo_pago: l.tipo_pago || null, _base: l.base}));
    const idDe = n => { const x = L.find(y => y.nombre === n && y.clase !== 'promotor'); return x ? x.id : null; };
    return L.map((l, i) => {
      const base_linea_id = l._base === 'casa' ? 'casa' : (l._base && l._base !== 'total') ? idDe(l._base) : null;
      const sobre = base_linea_id ? L.find(x => x.id === base_linea_id).nombre : null;
      const {_base, ...rest} = l;
      return {...rest, base_linea_id, orden: i, concepto: l.clase === 'casa' ? l.concepto : sobre ? `COMISIÓN ${l.nombre} (${l.pct}% DE ${sobre})` : `TERCIARIA ${l.nombre}`};
    });
  }
  /** Totales: comisión = suma de líneas sin base (las desglosadas salen de su línea padre). */
  function totalesComision(lineas) {
    const top = lineas.filter(l => !l.base_linea_id);
    return {total: r2(top.reduce((s, l) => s + l.monto, 0)), pct: r2(top.reduce((s, l) => s + (+l.pct || 0), 0) * 10000) / 10000};
  }
  /** Neto que recibe cada línea padre después de pagar sus terciarias. */
  function netoLinea(linea, lineas) {
    return r2(linea.monto - lineas.filter(l => l.base_linea_id === linea.id).reduce((s, l) => s + l.monto, 0));
  }

  /** Retorno: el cliente deposita X, se cobra pct% y se le devuelve el resto. */
  function calcRetorno(deposito, pct) {
    const comision = r2(deposito * pct / 100);
    return {comision, retorno: r2(deposito - comision)};
  }
  /** Colorines: cada persona debe recibir "monto"; el cliente deposita monto/(1-pct). */
  function grossUp(monto, pct) {
    const total = r2(+monto / (1 - pct / 100));
    return {monto: r2(monto), comision: r2(total - monto), total};
  }
  function resumenDispersion(lineas, pct, deposito) {
    const s = k => r2(lineas.reduce((a, l) => a + (+l[k] || 0), 0));
    const monto = s('monto'), total = s('total'), comision = s('comision');
    const porEstrategia = {};
    for (const l of lineas) { const e = l.estrategia || 'SCU'; porEstrategia[e] = r2((porEstrategia[e] || 0) + (+l.monto || 0)); }
    const dispersado = r2(lineas.filter(l => l.estado === 'dispersado').reduce((a, l) => a + (+l.monto || 0), 0));
    return {monto, comision, total, diferencia: deposito === undefined ? null : r2(deposito - total), porEstrategia,
      dispersado, pendiente: r2(monto - dispersado),
      nPendientes: lineas.filter(l => l.estado !== 'dispersado').length, n: lineas.length};
  }

  /* ----- lectura del archivo de personas a dispersar ----- */
  const ESTRATEGIAS = ['SCU', 'PAYPAL', 'EFECTIVO', 'USA'];
  const toNum = v => typeof v === 'number' ? v : +String(v ?? '').replace(/[$,\s]/g, '') || 0;
  /** rows: arreglo de filas (arreglos) como lo da XLSX.utils.sheet_to_json(ws, {header: 1}).
      Busca la fila de encabezados con NOMBRE y una columna de monto. */
  function parsePrenomina(rows) {
    let h = -1, cNom = -1, cMonto = -1, cModo = -1, cPuesto = -1;
    for (let i = 0; i < Math.min(rows.length, 40) && h < 0; i++) {
      const r = rows[i] || [];
      r.forEach((v, j) => {
        const t = normName(v);
        if (cNom < 0 && /^(NOMBRE|NOMBRES|BENEFICIARIO|NOMBRE COMPLETO|COLABORADOR|EMPLEADO)$/.test(t)) cNom = j;
        if (cMonto < 0 && /(MONTO|TOTAL|IMPORTE|DISPERS|NETO|PAGO)/.test(t)) cMonto = j;
        if (cModo < 0 && /^(MODO|ESTRATEGIA|TIPO)$/.test(t)) cModo = j;
        if (cPuesto < 0 && /^(PUESTO|AREA|DEPARTAMENTO|CARGO)$/.test(t)) cPuesto = j;
      });
      if (cNom >= 0 && cMonto >= 0) h = i; else { cNom = cMonto = cModo = cPuesto = -1; }
    }
    if (h < 0) throw new Error('No encontré en el archivo una fila de encabezados con "NOMBRE" y una columna de monto (MONTO, TOTAL, IMPORTE…).');
    const out = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || []; const nombre = String(r[cNom] ?? '').trim(); const monto = toNum(r[cMonto]);
      if (!nombre || !(monto > 0)) continue;
      if (/^(TOTAL|SUBTOTAL)/.test(normName(nombre))) continue;
      let estrategia = cModo >= 0 ? normName(r[cModo]) : '';
      if (!ESTRATEGIAS.includes(estrategia)) estrategia = (r.map(normName).find(v => ESTRATEGIAS.includes(v))) || 'SCU';
      let puesto = cPuesto >= 0 ? String(r[cPuesto] ?? '').trim() : '';
      if (!puesto) { // columna sin encabezado entre el nombre y el monto (como en SAGVE NOM SSD)
        for (let j = cNom + 1; j < cMonto; j++) { const v = String(r[j] ?? '').trim(); if (v && !ESTRATEGIAS.includes(normName(v)) && isNaN(toNum(v) || NaN)) { puesto = v; break; } }
      }
      out.push({nombre_archivo: nombre.replace(/\s+/g, ' '), monto: r2(monto), estrategia, puesto});
    }
    return out;
  }

  /* ----- importación de BD ALTAS (hoja BDCOLBCOL) ----- */
  function parseBdAltas(rows) {
    let h = -1;
    for (let i = 0; i < Math.min(rows.length, 20); i++) if ((rows[i] || []).some(v => normName(v) === 'BENEFICIARIO')) { h = i; break; }
    if (h < 0) throw new Error('No encontré la columna BENEFICIARIO en el archivo.');
    const H = rows[h].map(normName); const col = (...names) => H.findIndex(x => names.includes(x));
    const c = {idx: col('INDICE'), pre: col('PRENOMINA'), ben: col('BENEFICIARIO'), banco: col('BANCO BENEFICIARIO', 'BANCO'),
      cuenta: col('CUENTA'), clabe: col('CLABE INTERBANCARIA', 'CLABE'), swift: col('SWFIT CODE', 'SWIFT CODE', 'SWIFT'),
      aba: col('ABA NUMBER', 'ABA'), cn: col('CUENTA NUEVA'), kn: col('CLABE NUEVA'), obs: col('OBSERVACIONES')};
    const g = (r, k) => c[k] >= 0 ? r[c[k]] : '';
    const out = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || []; const nombre = String(g(r, 'ben') ?? '').toUpperCase().replace(/\s+/g, ' ').trim(); if (!normName(nombre)) continue;
      const nuevaClabe = normalizarClabe(g(r, 'kn')), nuevaCuenta = normalizarCuenta(g(r, 'cn'));
      out.push({
        indice: +g(r, 'idx') || null, nombre, nombre_alt: String(g(r, 'pre') || '').trim() || null,
        banco: normBanco(g(r, 'banco')) || bancoDeClabe(nuevaClabe || g(r, 'clabe')),
        cuenta: nuevaCuenta || normalizarCuenta(g(r, 'cuenta')), clabe: nuevaClabe || normalizarClabe(g(r, 'clabe')),
        swift: String(g(r, 'swift') || '').trim() || null, aba: String(g(r, 'aba') || '').trim() || null,
        observaciones: String(g(r, 'obs') || '').trim() || null,
      });
    }
    return out;
  }

  /* ----- importe con letra (para fichas) ----- */
  function numLetras(n) {
    const U = ['', 'UN', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE', 'DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE',
      'DIECISEIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE', 'VEINTE', 'VEINTIUN', 'VEINTIDOS', 'VEINTITRES', 'VEINTICUATRO', 'VEINTICINCO', 'VEINTISEIS', 'VEINTISIETE', 'VEINTIOCHO', 'VEINTINUEVE'];
    const D = ['', '', '', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA'];
    const C = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS'];
    const cien = x => { if (x === 100) return 'CIEN'; const c = Math.floor(x / 100), r = x % 100; let s = C[c];
      if (r) s += (s ? ' ' : '') + (r < 30 ? U[r] : D[Math.floor(r / 10)] + (r % 10 ? ' Y ' + U[r % 10] : '')); return s; };
    const mil = x => { const m = Math.floor(x / 1000), r = x % 1000; return (m ? (m === 1 ? 'MIL' : cien(m) + ' MIL') : '') + (r ? (m ? ' ' : '') + cien(r) : ''); };
    const ent = Math.floor(Math.abs(n)), cts = Math.round((Math.abs(n) - ent) * 100);
    const mm = Math.floor(ent / 1e6), rest = ent % 1e6;
    let s = (mm ? (mm === 1 ? 'UN MILLON' : mil(mm) + ' MILLONES') : '') + (rest ? (mm ? ' ' : '') + mil(rest) : '');
    if (!s) s = 'CERO'; if (mm && !rest) s += ' DE';
    return `${s} PESOS ${String(cts).padStart(2, '0')}/100 M.N.`;
  }

  function edad(fechaNac, ref = new Date()) {
    const d = new Date(fechaNac); if (isNaN(d)) return '';
    let e = ref.getFullYear() - d.getFullYear(); const m = ref.getMonth() - d.getMonth();
    if (m < 0 || (m === 0 && ref.getDate() < d.getDate())) e--; return e;
  }

  /* ----- cuestionario de alta (mismo para el formulario público y la app) ----- */
  const BANCOS_FORM = ['BBVA', 'SANTANDER', 'BANORTE', 'BANAMEX', 'HSBC', 'SCOTIABANK', 'BANREGIO', 'BANCO AZTECA', 'AFIRME', 'BAJIO', 'INBURSA', 'HEY BANCO', 'BANCOPPEL', 'NU MEXICO', 'MIFEL', 'OTRO'];
  // [clave, etiqueta, tipo, obligatorio, opciones | marcador del machote]
  const CAMPOS_ALTA = [
    {sec: 'Datos personales'},
    {k: 'apellido_paterno', l: 'Apellido paterno', req: 1, ph: 'APELLIDO PATERNO', up: 1},
    {k: 'apellido_materno', l: 'Apellido materno', req: 1, ph: 'APELLIDO MATERNO', up: 1},
    {k: 'nombres', l: 'Nombre(s)', req: 1, ph: 'NOMBRES', up: 1},
    {k: 'fecha_nacimiento', l: 'Fecha de nacimiento', t: 'date', req: 1, ph: 'FECHA NACIMIENTO'},
    {k: 'lugar_nacimiento', l: 'Lugar de nacimiento', req: 1, ph: 'LUGAR DE NACIMIENTO', up: 1},
    {k: 'sexo', l: 'Sexo', t: 'select', req: 1, o: ['MASCULINO', 'FEMENINO'], ph: 'SEXO'},
    {k: 'estado_civil', l: 'Estado civil', t: 'select', req: 1, o: ['SOLTERO(A)', 'CASADO(A) BIENES MANCOMUNADOS', 'CASADO(A) SEPARACIÓN DE BIENES', 'UNIÓN LIBRE', 'DIVORCIADO(A)', 'VIUDO(A)'], ph: 'ESTADO CIVIL'},
    {k: 'rfc', l: 'RFC (con homoclave)', req: 1, ph: 'RFC', up: 1, re: '^[A-ZÑ&]{4}\\d{6}[A-Z0-9]{3}$', err: 'El RFC debe tener 13 caracteres'},
    {k: 'curp', l: 'CURP', req: 1, ph: 'CURP', up: 1, re: '^[A-Z]{4}\\d{6}[HM][A-Z]{5}[A-Z0-9]\\d$', err: 'La CURP debe tener 18 caracteres'},
    {k: 'celular', l: 'Número celular (10 dígitos)', t: 'tel', req: 1, ph: 'NUMERO CELULAR', re: '^\\d{10}$', err: 'Escribe 10 dígitos'},
    {k: 'correo', l: 'Correo electrónico', t: 'email', req: 1, ph: 'CORREO ELECTRONICO'},
    {sec: 'Domicilio'},
    {k: 'calle', l: 'Calle', req: 1, ph: 'CALLE', up: 1},
    {k: 'no_exterior', l: 'No. exterior', req: 1, ph: 'No. EXTERIOR', up: 1},
    {k: 'no_interior', l: 'No. interior', ph: 'No. INTERIOR', up: 1},
    {k: 'colonia', l: 'Colonia', req: 1, ph: 'COLONIA', up: 1},
    {k: 'municipio', l: 'Delegación o municipio', req: 1, ph: 'MUNICIPIO', up: 1},
    {k: 'localidad', l: 'Localidad', req: 1, ph: 'LOCALIDAD', up: 1},
    {k: 'estado', l: 'Estado', req: 1, ph: 'ESTADO', up: 1},
    {k: 'cp', l: 'Código postal', req: 1, ph: 'CP', re: '^\\d{5}$', err: 'El C.P. tiene 5 dígitos'},
    {k: 'pais', l: 'País', req: 1, ph: 'PAÍS', up: 1, def: 'MÉXICO'},
    {sec: 'Estudios y experiencia'},
    {k: 'nivel_estudios', l: 'Nivel máximo de estudios', t: 'select', req: 1, o: ['Primaria', 'Secundaria', 'Preparatoria', 'Carrera técnica', 'Licenciatura', 'Maestría', 'Doctorado'], ph: 'NIVEL DE ESTUDIOS'},
    {k: 'titulado', l: '¿Titulado?', t: 'select', req: 1, o: ['Sí', 'No'], ph: 'TITULO'},
    {k: 'cedula', l: 'Cédula profesional', ph: 'CEDULA PROFESIONAL'},
    {k: 'experiencia', l: 'Experiencia profesional', t: 'textarea', req: 1, ph: 'EXPERIENCIA', up: 1},
    {sec: 'Datos bancarios'},
    {k: 'metodo_dispersion', l: 'Método de dispersión', t: 'select', req: 1, o: ['DEPOSITO BANCARIO', 'EFECTIVO', 'PAYPAL'], ph: 'METODO DE DISPERSIÓN'},
    {k: 'banco', l: 'Banco', t: 'select', req: 1, o: BANCOS_FORM, ph: 'BANCO'},
    {k: 'numero_cuenta', l: 'Número de cuenta', req: 1, ph: 'NUMERO DE CUENTA', re: '^\\d{8,12}$', err: 'Solo números (8 a 12 dígitos)'},
    {k: 'clabe', l: 'CLABE interbancaria (18 dígitos)', req: 1, ph: 'CLABE', t: 'clabe'},
    {sec: 'Documentos'},
    {k: 'firma', l: 'Firma (dibújala con el dedo o el mouse)', t: 'firma', req: 1, img: 'FIRMA'},
    {k: 'huella', l: 'Foto de tu huella digital (índice derecho)', t: 'file', req: 1, img: 'HUELLA'},
    {k: 'nombre_mano', l: 'Foto de tu nombre escrito a mano', t: 'file', req: 1, img: 'NOMBRE ESCRITO A MANO'},
    {k: 'ine_frente', l: 'INE por el frente', t: 'file', req: 1},
    {k: 'ine_reverso', l: 'INE por el reverso', t: 'file', req: 1},
  ];
  /** Valida las respuestas del cuestionario. Devuelve {campo: mensaje} */
  function validarAlta(d) {
    const e = {};
    for (const c of CAMPOS_ALTA) {
      if (!c.k || ['file', 'firma'].includes(c.t)) continue;
      const v = String(d[c.k] ?? '').trim();
      if (c.req && !v) { if (!(c.k === 'numero_cuenta' || c.k === 'clabe' || c.k === 'banco') || d.metodo_dispersion === 'DEPOSITO BANCARIO') e[c.k] = 'Obligatorio'; continue; }
      if (v && c.re && !new RegExp(c.re).test(c.up ? v.toUpperCase() : v)) e[c.k] = c.err || 'Formato incorrecto';
      if (v && c.t === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) e[c.k] = 'Correo no válido';
      if (v && c.t === 'clabe') { const r = validarClabe(v, d.banco); if (!r.ok) e[c.k] = r.motivo; }
    }
    return e;
  }
  /** Valores para llenar el machote a partir de las respuestas. */
  function valoresContrato(d, fechaInicio) {
    const fmt = s => { if (!s) return ''; const [y, m, dd] = String(s).slice(0, 10).split('-'); return dd ? `${dd}/${m}/${y}` : s; };
    const v = {};
    for (const c of CAMPOS_ALTA) if (c.ph) v[c.ph] = c.t === 'date' ? fmt(d[c.k]) : String(d[c.k] ?? '').trim()[c.up ? 'toUpperCase' : 'toString']();
    v['APELLIDO_PATERNO'] = v['APELLIDO PATERNO']; v['APELLIDO_MATERNO'] = v['APELLIDO MATERNO']; v['_NOMBRES_'] = v['NOMBRES'];
    v['SOCIO'] = [v['APELLIDO PATERNO'], v['APELLIDO MATERNO'], v['NOMBRES']].filter(Boolean).join(' ');
    v['FECHA INICIO'] = fmt(fechaInicio);
    v['EDAD'] = String(edad(d.fecha_nacimiento, fechaInicio ? new Date(fechaInicio + 'T12:00:00') : new Date()));
    v['CLABE'] = normalizarClabe(d.clabe);
    return v;
  }

  return {CAMPOS_ALTA, BANCOS_FORM, validarAlta, valoresContrato, r2, normName, tokens, nameScore, matchNombre, indexBens, BANCOS, normBanco, normalizarClabe, validarClabe,
    bancoDeClabe, normalizarCuenta, calcComisiones, ajustarCuadre, ESQUEMA_COLORINES, normalizarEsquema, lineasColorines, totalesComision, netoLinea, calcRetorno, grossUp, resumenDispersion,
    ESTRATEGIAS, parsePrenomina, parseBdAltas, numLetras, edad, toNum};
});
