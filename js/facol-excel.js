/* ---------- FACOL: libro de Excel de una operación Colorines (mismo formato que FACOL2354N) ----------
   Hojas: COMISIONES (cálculo y desglose), SSP (comisión por persona), PRENOMINA (formato por estrategia)
   y FORMATO DISP (datos del movimiento, formato de dispersión y pago a promotores). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FacolExcel = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const MONEY = '"$"#,##0.00'; const R = n => Math.round((+n || 0) * 100) / 100;
  const AZUL = 'FF0F2F5C', CLARO = 'FFE3EDFB', GRIS = 'FFF2F4F7';
  const fill = c => ({type: 'pattern', pattern: 'solid', fgColor: {argb: c}});
  const borde = {top: {style: 'thin', color: {argb: 'FFB7C0CC'}}, left: {style: 'thin', color: {argb: 'FFB7C0CC'}}, bottom: {style: 'thin', color: {argb: 'FFB7C0CC'}}, right: {style: 'thin', color: {argb: 'FFB7C0CC'}}};
  const fmtDay = d => { if (!d) return ''; const [y, m, dd] = String(d).slice(0, 10).split('-'); return `${dd}/${m}/${y}`; };
  const ESTR = ['SCU', 'PAYPAL', 'EFECTIVO', 'USA'];

  function titulo(ws, ref, text) { ws.mergeCells(ref); const c = ws.getCell(ref.split(':')[0]); c.value = text; c.font = {bold: true, size: 13, color: {argb: 'FFFFFFFF'}}; c.fill = fill(AZUL); c.alignment = {horizontal: 'center', vertical: 'middle'}; }
  function encabezado(row, from, to) { for (let c = from; c <= to; c++) { const x = row.getCell(c); x.font = {bold: true, color: {argb: 'FFFFFFFF'}}; x.fill = fill(AZUL); x.border = borde; x.alignment = {horizontal: 'center', vertical: 'middle', wrapText: true}; } }
  function celda(ws, addr, value, o = {}) {
    const c = ws.getCell(addr); c.value = value; c.border = borde;
    if (o.money) c.numFmt = MONEY; if (o.text) c.numFmt = '@'; if (o.bold) c.font = {bold: true}; if (o.fill) c.fill = fill(o.fill);
    return c;
  }

  /**
   * @param ExcelJS librería
   * @param datos {op, lineas: dispersiones, comisiones: operacion_comisiones, coms: catálogo de promotores}
   */
  async function generar(ExcelJS, {op, lineas, comisiones, coms = [], factura = null}) {
    const wb = new ExcelJS.Workbook(); wb.creator = 'FACOL'; wb.created = new Date();
    const pct = +op.pct_comision; const dep = +op.monto_deposito;
    const n = lineas.length;

    const cs = wb.addWorksheet('COMISIONES'), ssp = wb.addWorksheet('SSP'), pre = wb.addWorksheet('PRENOMINA'), fd = wb.addWorksheet('FORMATO DISP');
    /* ---------- SSP (por persona) ---------- */
    ssp.columns = [{width: 44}, {width: 11}, {width: 16}, {width: 14}, {width: 16}, {width: 3}, {width: 22}, {width: 14}];
    titulo(ssp, 'A1:E1', `SSP ${op.cliente_nombre || ''} · ${op.folio}`);
    const F = 6, L = F + n - 1; // filas de datos
    celda(ssp, 'B4', 'Subtotales', {bold: true});
    celda(ssp, 'C4', {formula: `SUM(C${F}:C${L})`, result: lineas.reduce((s, d) => s + +d.monto, 0)}, {money: 1, bold: 1, fill: CLARO});
    celda(ssp, 'D4', {formula: `SUM(D${F}:D${L})`, result: lineas.reduce((s, d) => s + +d.comision, 0)}, {money: 1, bold: 1, fill: CLARO});
    const totalSSP = lineas.reduce((s, d) => s + +d.total, 0);
    celda(ssp, 'E4', {formula: `SUM(E${F}:E${L})`, result: totalSSP}, {money: 1, bold: 1, fill: CLARO});
    ['NOMBRE', 'MODO', 'MONTO DESEADO', 'COMISIÓN', 'TOTAL'].forEach((h, i) => ssp.getRow(5).getCell(i + 1).value = h); encabezado(ssp.getRow(5), 1, 5);
    lineas.forEach((d, i) => { const r = F + i;
      celda(ssp, `A${r}`, d.nombre || d.nombre_archivo); celda(ssp, `B${r}`, d.estrategia);
      celda(ssp, `C${r}`, +d.monto, {money: 1}); celda(ssp, `D${r}`, +d.comision, {money: 1}); celda(ssp, `E${r}`, {formula: `C${r}+D${r}`, result: +d.total}, {money: 1}); });
    celda(ssp, 'G4', 'IMPORTE TRANSFERENCIA', {bold: 1, fill: GRIS}); celda(ssp, 'H4', 'DIFERENCIA', {bold: 1, fill: GRIS});
    celda(ssp, 'G5', dep, {money: 1}); celda(ssp, 'H5', {formula: 'G5-E4', result: Math.round((dep - totalSSP) * 100) / 100}, {money: 1, bold: 1});
    ssp.views = [{state: 'frozen', ySplit: 5}];

    /* ---------- COMISIONES (cálculo, como la hoja SCU) ---------- */
    cs.columns = [{width: 52}, {width: 18}, {width: 18}];
    titulo(cs, 'A1:C1', `CÁLCULO DE COMISIONES · ${op.folio}`);
    celda(cs, 'A2', 'MOVIMIENTO EN EFECTIVO', {bold: 1, fill: CLARO}); celda(cs, 'B2', 'PRENÓMINA', {bold: 1, fill: CLARO}); celda(cs, 'C2', 'DEPÓSITO', {bold: 1, fill: CLARO});
    ['DESCRIPCIÓN DEL MOVIMIENTO', 'IMPORTE', 'IMPORTE'].forEach((h, i) => cs.getRow(3).getCell(i + 1).value = h); encabezado(cs.getRow(3), 1, 3);
    const comTot = lineas.reduce((s, d) => s + +d.comision, 0);
    celda(cs, 'A4', 'DISPERSIÓN EFECTIVO'); celda(cs, 'B4', {formula: "SSP!E4", result: totalSSP}, {money: 1}); celda(cs, 'C4', dep, {money: 1});
    celda(cs, 'A5', `COMISIÓN A GRUPO COLORINES (${pct}%)`); celda(cs, 'B5', {formula: 'SSP!D4', result: comTot}, {money: 1}); celda(cs, 'C5', {formula: `C4*${pct / 100}`, result: R(dep * pct / 100)}, {money: 1});
    celda(cs, 'A6', 'TOTAL A DISPERSAR', {bold: 1}); celda(cs, 'B6', {formula: 'B4-B5', result: R(totalSSP - comTot)}, {money: 1, bold: 1}); celda(cs, 'C6', {formula: 'C4-C5', result: R(dep - dep * pct / 100)}, {money: 1, bold: 1});
    cs.mergeCells('A8:C8'); celda(cs, 'A8', 'DESGLOSE PAGO DE COMISIONES', {bold: 1, fill: CLARO});
    let r = 9; const top = comisiones.filter(l => !l.base_linea_id); const filasTop = [];
    for (const l of top) {
      celda(cs, `A${r}`, `${l.concepto || l.nombre} ${l.pct}%`); celda(cs, `B${r}`, +l.monto, {money: 1}); filasTop.push(r); r++;
      for (const h of comisiones.filter(x => x.base_linea_id === l.id)) { const c = celda(cs, `A${r}`, `      ↳ ${h.nombre} ${h.pct}% de ${l.nombre} (sale de ${l.nombre})`); c.font = {italic: true, color: {argb: 'FF555555'}}; celda(cs, `C${r}`, +h.monto, {money: 1}); r++; }
    }
    celda(cs, `A${r}`, 'CUADRE CASA'); celda(cs, `B${r}`, {formula: 'C4-B4', result: R(dep - totalSSP)}, {money: 1}); r++;
    celda(cs, `A${r}`, 'SUBTOTAL', {bold: 1, fill: GRIS});
    celda(cs, `B${r}`, {formula: filasTop.length ? filasTop.map(x => 'B' + x).join('+') : '0', result: top.reduce((s, l) => s + +l.monto, 0)}, {money: 1, bold: 1, fill: GRIS});
    celda(cs, `C${r}`, {formula: `B${r}-B5`, result: R(top.reduce((s, l) => s + +l.monto, 0) - comTot)}, {money: 1, fill: GRIS});

    /* ---------- PRENOMINA (formato de dispersión por estrategia, lado a lado) ---------- */
    // EFECTIVO, PAYPAL y los que se pagan con SWIFT van en FORMATO DISP; el resto (SCU…) va en PRENOMINA
    const enFormato = d => ['EFECTIVO', 'PAYPAL', 'USA'].includes(d.estrategia) || !!String(d.swift || '').trim();
    const prenom = lineas.filter(d => !enFormato(d)), formato = lineas.filter(enFormato);
    let col = 1;
    const estrPre = [...new Set([...ESTR, ...prenom.map(d => d.estrategia)])];
    if (!prenom.length) { pre.getColumn(1).width = 60; pre.getCell('A2').value = 'Sin colaboradores para prenómina en esta operación (todos van en FORMATO DISP).'; pre.getCell('A2').font = {italic: true, color: {argb: 'FF666666'}}; }
    for (const e of estrPre) {
      const ds = prenom.filter(d => d.estrategia === e); if (!ds.length) continue;
      const hdr = e === 'USA' ? ['ESTRATEGIA', 'BENEFICIARIO', 'BANCO BENEFICIARIO', 'MONTO A DISPERSAR', 'CUENTA', 'SWIFT CODE', 'ABA NUMBER'] : ['ESTRATEGIA', 'BENEFICIARIO', 'BANCO BENEFICIARIO', 'MONTO A DISPERSAR', 'CUENTA', 'CLABE INTERBANCARIA'];
      const widths = [12, 42, 18, 17, 15, 22, 14];
      hdr.forEach((_, i) => pre.getColumn(col + i).width = widths[i]); pre.getColumn(col + hdr.length).width = 3;
      const A = i => pre.getRow(1).getCell(col + i).address.replace(/\d+/, '');
      pre.mergeCells(`${A(0)}2:${A(hdr.length - 1)}2`); const t = pre.getCell(`${A(0)}2`); t.value = `INCIDENCIAS ${e} ${op.cliente_nombre || ''}`; t.font = {bold: true, color: {argb: 'FFFFFFFF'}}; t.fill = fill(AZUL); t.alignment = {horizontal: 'center'};
      const f0 = 6, f1 = f0 + ds.length - 1;
      celda(pre, `${A(2)}4`, 'A DISPERSAR', {bold: 1}); celda(pre, `${A(3)}4`, {formula: `SUM(${A(3)}${f0}:${A(3)}${f1})`, result: ds.reduce((s, d) => s + +d.monto, 0)}, {money: 1, bold: 1, fill: CLARO});
      celda(pre, `${A(4)}4`, 'COMISIÓN', {bold: 1}); celda(pre, `${A(5)}4`, ds.reduce((s, d) => s + +d.comision, 0), {money: 1, bold: 1, fill: CLARO});
      hdr.forEach((h, i) => pre.getRow(5).getCell(col + i).value = h); encabezado(pre.getRow(5), col, col + hdr.length - 1);
      ds.forEach((d, i) => { const rr = f0 + i;
        const vals = [d.estrategia, d.nombre || d.nombre_archivo, d.banco || '', +d.monto, d.cuenta || '', ...(e === 'USA' ? [d.swift || '', d.aba || ''] : [d.clabe || ''])];
        vals.forEach((v, j) => celda(pre, `${A(j)}${rr}`, v, {money: j === 3, text: j >= 4}));
        if (e !== 'EFECTIVO' && !d.beneficiario_id) pre.getCell(`${A(1)}${rr}`).fill = fill('FFFDE3E3');
      });
      col += hdr.length + 1;
    }
    pre.views = [{state: 'frozen', ySplit: 5}];

    /* ---------- FORMATO DISP ---------- */
    fd.columns = [{width: 3}, {width: 34}, {width: 32}, {width: 3}, {width: 12}, {width: 42}, {width: 17}, {width: 18}, {width: 15}, {width: 22}, {width: 13}, {width: 13}];
    titulo(fd, 'B1:C1', 'DATOS DEL MOVIMIENTO');
    titulo(fd, 'E1:L1', 'FORMATO DE DISPERSIÓN');
    const kv = [['FOLIO', op.folio], ['FECHA DE SOLICITUD', fmtDay(op.fecha_solicitud)], ['NOMBRE FISCAL DEL CLIENTE', op.razon_social || op.cliente_nombre], ['BANCO ORDENANTE', op.banco_ordenante],
      ['CUENTA ORDENANTE', op.cuenta_ordenante], ['FECHA DE DEPÓSITO', fmtDay(op.fecha_deposito)], ['EMPRESA QUE FACTURA', op.empresa_factura], ['BANCO RECEPTOR', op.banco_receptor],
      ['CUENTA RECEPTORA', op.cuenta_receptora], ['SOLICITANTE', op.solicitante], ['FECHA COMPROMISO DE PAGO', fmtDay(op.fecha_compromiso)], ['AUTORIZADO POR', op.autorizado_por],
      ['MONTO DEPOSITADO', dep, 1], ['COMISIÓN', pct + '%'], ['MONTO DE COMISIÓN', comTot, 1], ['TOTAL A DISPERSAR', lineas.reduce((s, d) => s + +d.monto, 0), 1]];
    kv.forEach(([k, v, m], i) => { celda(fd, `B${i + 3}`, k, {bold: 1, fill: GRIS}); celda(fd, `C${i + 3}`, v ?? '', {money: m}); });
    let rr = kv.length + 5;
    const ctaDe = l => coms.find(c => c.id === l.comisionista_id) || {};
    for (const l of [...top, ...comisiones.filter(x => x.base_linea_id)]) {
      if (!(+l.monto > 0)) continue;
      const par = l.base_linea_id && comisiones.find(x => x.id === l.base_linea_id);
      fd.mergeCells(`B${rr}:C${rr}`); const h = fd.getCell(`B${rr}`); h.value = l.clase === 'casa' ? 'COMISIÓN CASA' : par ? `COMISIÓN DESGLOSADA DE ${par.nombre}` : 'DATOS DEL COMISIONISTA O PROMOTOR'; h.font = {bold: true}; h.fill = fill(CLARO); rr++;
      const c = ctaDe(l);
      for (const [k, v, m] of [['FECHA DE ENTREGA', fmtDay(op.fecha_compromiso || op.fecha_deposito)], ['COMISIONISTA O PROMOTOR', l.nombre], ['TIPO DE PAGO', l.tipo_pago || ''], ['MONTO DE LA COMISIÓN A COBRAR', +l.monto, 1], ['CUENTA', c.cuenta || ''], ['BANCO', c.banco || ''], ['CLABE', c.clabe || '']]) {
        celda(fd, `B${rr}`, k, {fill: GRIS}); celda(fd, `C${rr}`, v, {money: m, text: ['CUENTA', 'CLABE'].includes(k)}); rr++;
      }
      rr++;
    }
    ['ESTRATEGIA', 'BENEFICIARIO', 'MONTO A DISPERSAR', 'BANCO BENEFICIARIO', 'CUENTA', 'CLABE INTERBANCARIA', 'SWIFT CODE', 'ABA NUMBER'].forEach((h, i) => fd.getRow(3).getCell(5 + i).value = h); encabezado(fd.getRow(3), 5, 12);
    formato.forEach((d, i) => { const r4 = 4 + i;
      [d.estrategia, d.nombre || d.nombre_archivo, +d.monto, d.banco || '', d.cuenta || '', d.clabe || '', d.swift || '', d.aba || ''].forEach((v, j) => celda(fd, fd.getRow(r4).getCell(5 + j).address, v, {money: j === 2, text: j >= 4})); });
    if (!formato.length) { fd.mergeCells('E4:L4'); const c = fd.getCell('E4'); c.value = 'Sin colaboradores en efectivo, PayPal o SWIFT (todos van en PRENOMINA).'; c.font = {italic: true, color: {argb: 'FF666666'}}; }
    const tr = 4 + Math.max(formato.length, 1); celda(fd, `F${tr}`, 'TOTAL', {bold: 1, fill: GRIS});
    celda(fd, `G${tr}`, {formula: formato.length ? `SUM(G4:G${tr - 1})` : '0', result: R(formato.reduce((s, d) => s + +d.monto, 0))}, {money: 1, bold: 1, fill: GRIS});

    /* ---------- FORMATO FACTURA (oculta): idéntica a la del FACOL de ejemplo ---------- */
    if (factura && factura.n > 0 && factura.plantilla) await formatoFactura(ExcelJS, wb, op, dep, factura);
    /* (respaldo si no se tiene el formato) SOLICITUD FACTURA sencilla */
    else if (factura && factura.n > 0) {
      const sf = wb.addWorksheet('SOLICITUD FACTURA'); const d = factura.datos || {};
      sf.columns = [{width: 3}, {width: 30}, {width: 46}, {width: 3}, {width: 12}, {width: 34}, {width: 18}, {width: 18}];
      const n = Math.max(1, Math.floor(factura.n));
      // reparto en centavos para que la suma sea exacta: subtotal = depósito / 1.16
      const totalC = Math.round(dep * 100), base = Math.floor(totalC / n);
      const totales = Array.from({length: n}, (_, i) => (i === n - 1 ? totalC - base * (n - 1) : base) / 100);
      let r = 1;
      totales.forEach((tot, i) => {
        const sub = R(tot / 1.16), iva = R(tot - sub);
        titulo(sf, `B${r}:H${r}`, `SOLICITUD DE FACTURA ${i + 1} DE ${n} · ${op.folio}`); r += 2;
        for (const [k, v] of [['EMPRESA QUE FACTURA', op.empresa_factura || d.empresa_factura], ['NOMBRE O RAZÓN SOCIAL CLIENTE', op.razon_social || op.cliente_nombre], ['R.F.C.', d.rfc || ''],
          ['FECHA', fmtDay(op.fecha_deposito)], ['FORMA DE PAGO', d.forma_pago_sat || '03 Transferencia electrónica de fondos'], ['MÉTODO DE PAGO', d.metodo_pago || 'PUE Pago en una sola exhibición'],
          ['USO CFDI', d.uso_cfdi || 'G03 Gastos en general'], ['MONEDA', 'MXN - PESO MEXICANO'], ['BANCO / CUENTA', [op.banco_receptor, op.cuenta_receptora].filter(Boolean).join(' / ')]]) {
          celda(sf, `B${r}`, k, {bold: 1, fill: GRIS}); sf.mergeCells(`C${r}:H${r}`); celda(sf, `C${r}`, v || ''); r++;
        }
        r++;
        sf.getRow(r).getCell(2).value = 'CANTIDAD'; sf.getRow(r).getCell(3).value = 'DESCRIPCIÓN'; sf.getRow(r).getCell(7).value = 'VALOR UNITARIO'; sf.getRow(r).getCell(8).value = 'IMPORTE';
        encabezado(sf.getRow(r), 2, 8); sf.mergeCells(`C${r}:F${r}`); r++;
        celda(sf, `B${r}`, 1); sf.mergeCells(`C${r}:F${r}`); celda(sf, `C${r}`, d.concepto_factura || 'COMISIONES MERCANTILES'); celda(sf, `G${r}`, sub, {money: 1}); celda(sf, `H${r}`, {formula: `B${r}*G${r}`, result: sub}, {money: 1}); const fr = r; r++;
        celda(sf, `G${r}`, 'SUBTOTAL', {bold: 1, fill: GRIS}); celda(sf, `H${r}`, {formula: `H${fr}`, result: sub}, {money: 1}); r++;
        celda(sf, `G${r}`, 'IVA 16%', {bold: 1, fill: GRIS}); celda(sf, `H${r}`, iva, {money: 1}); r++;
        celda(sf, `G${r}`, 'TOTAL', {bold: 1, fill: CLARO}); celda(sf, `H${r}`, {formula: `H${r - 2}+H${r - 1}`, result: tot}, {money: 1, bold: 1, fill: CLARO}); r += 3;
      });
      sf.state = 'hidden';
    }

    // la hoja de comisiones va oculta; ninguna hoja lleva líneas de cuadrícula
    cs.state = 'hidden';
    for (const w of wb.worksheets) {
      w.properties.tabColor = {argb: AZUL};
      const v = (w.views && w.views[0]) || {}; w.views = [{...v, showGridLines: false}];
    }
    wb.views = [{activeTab: 1, firstSheet: 1, visibility: 'visible'}];
    const buf = await wb.xlsx.writeBuffer();
    return new Uint8Array(buf);
  }
  /** Copia la hoja FORMATO FACTURA del formato (app/formatos/formato-factura.xlsx) y llena un bloque por factura.
      Cada bloque ocupa 40 filas (empieza en la 3, 43, 83…); la celda Q1 es "No. Fact". */
  async function formatoFactura(ExcelJS, wb, op, dep, factura) {
    // Formato estandarizado: se copia tal cual (colores, letras, bordes, textos y fórmulas) y solo se llenan los datos.
    const t = new ExcelJS.Workbook(); await t.xlsx.load(factura.plantilla);
    const src = t.worksheets[0]; const ws = wb.addWorksheet('FORMATO FACTURA');
    ws.model = Object.assign({}, src.model, {name: 'FORMATO FACTURA', mergeCells: src.model.merges});
    // las celdas dentro de una combinación pierden sus bordes/relleno al combinar: se copian uno por uno
    const dims = src.dimensions || {bottom: 200, right: 17};
    for (let r = 1; r <= Math.max(dims.bottom || 0, 200); r++) for (let c = 1; c <= Math.max(dims.right || 0, 17); c++) {
      const s = src.getCell(r, c).style; if (s && Object.keys(s).length) ws.getCell(r, c).style = JSON.parse(JSON.stringify(s));
    }
    // el tema de colores del formato original (sin esto los azules cambian de tono)
    if (t._themes && Object.keys(t._themes).length) wb._themes = t._themes;
    wb.calcProperties = Object.assign({}, wb.calcProperties, {fullCalcOnLoad: true});   // Excel recalcula al abrir, como en el original
    const d = factura.datos || {};
    const bloques = []; for (let r = 3; r < 400; r += 40) if (String(ws.getCell(`B${r + 1}`).value || '').startsWith('Empresa Que Factura')) bloques.push(r); else break;
    const n = Math.min(Math.max(1, Math.floor(factura.n)), bloques.length);
    ws.getCell('Q1').value = n;                                               // "No. Fact"
    const empresa = op.empresa_factura || d.empresa_factura || '', cliente = op.razon_social || op.cliente_nombre || '';
    const sub1 = Math.round(dep / 1.16 / n * 1e6) / 1e6;
    bloques.forEach((r, i) => {
      if (i >= n) return;                                                      // los bloques que no se usan quedan como en el formato
      ws.getCell(`F${r + 1}`).value = empresa;
      ws.getCell(`F${r + 3}`).value = cliente;
      if (d.rfc) ws.getCell(`D${r + 4}`).value = d.rfc;
    });
    // valor unitario igual que en el original: depósito (de FORMATO DISP) / 1.16 / No. Fact; los demás bloques lo toman con sus fórmulas
    ws.getCell(`K${bloques[0] + 16}`).value = {formula: `(((+'FORMATO DISP'!$C$15)/1.16))/$Q$1`, result: sub1};
    ws.state = 'hidden';
    return n;
  }
  return {generar, formatoFactura};
});
