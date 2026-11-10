/* ---------- FACOL: llenado del machote de contrato (.docx) ----------
   Reemplaza [CAMPO] en el documento, encabezados y pies, aunque Word haya partido el texto en
   varios fragmentos. Los campos de imagen ([HUELLA], [FIRMA]…) se sustituyen por la imagen. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('jszip'));
  else root.Docx = factory(root.JSZip);
})(typeof self !== 'undefined' ? self : this, function (JSZip) {
  const escXml = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const IMG_KEYS = ['HUELLA', 'HUELLA DIGITALIZADA', 'FIRMA', 'FIRMA DIGITALIZADA', 'NOMBRE ESCRITO A MANO', 'NOMBRE ESCRITO A MANO DIGITALIZADO'];

  /** Reemplaza en un XML de Word los marcadores [X] por valores[X] (texto) o por marcas §IMGn§ (imágenes). */
  function fillXml(xml, valores, imgMarks) {
    return xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
      const re = /(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g; const nodes = []; let m;
      while ((m = re.exec(p))) nodes.push({start: m.index, len: m[0].length, open: m[1], text: m[2], close: m[3]});
      if (!nodes.length) return p;
      const full = nodes.map(n => n.text).join('');
      if (!full.includes('[')) return p;
      // posiciones de cada nodo dentro del texto completo
      let pos = 0; for (const n of nodes) { n.a = pos; pos += n.text.length; n.b = pos; }
      const texts = nodes.map(n => n.text.split(''));
      const ph = /\[([^\[\]]{1,60})\]/g; const reps = []; let q;
      while ((q = ph.exec(full))) {
        const key = q[1].trim();
        let value;
        if (imgMarks && IMG_KEYS.includes(key.toUpperCase()) && imgMarks[key.toUpperCase()]) value = imgMarks[key.toUpperCase()];
        else if (Object.prototype.hasOwnProperty.call(valores, key)) value = escXml(valores[key]);
        else if (Object.prototype.hasOwnProperty.call(valores, key.toUpperCase())) value = escXml(valores[key.toUpperCase()]);
        else if (IMG_KEYS.includes(key.toUpperCase())) value = '';
        else continue;
        reps.push({a: q.index, b: q.index + q[0].length, value});
      }
      if (!reps.length) return p;
      for (const r of reps.reverse()) {
        // primer nodo recibe el valor, los demás pierden sus caracteres
        nodes.forEach((n, i) => {
          const s = Math.max(r.a, n.a), e = Math.min(r.b, n.b); if (s >= e) return;
          const ls = s - n.a, le = e - n.a;
          const isFirst = r.a >= n.a && r.a < n.b;
          texts[i].splice(ls, le - ls, ...(isFirst ? [r.value] : []));
        });
      }
      let out = '', last = 0;
      nodes.forEach((n, i) => {
        out += p.slice(last, n.start);
        const open = /xml:space/.test(n.open) ? n.open : n.open.replace('<w:t', '<w:t xml:space="preserve"');
        out += open + texts[i].join('') + n.close; last = n.start + n.len;
      });
      return out + p.slice(last);
    });
  }

  function drawingXml(rid, id, cx, cy) {
    return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Imagen${id}"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="Imagen${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  }

  /** Quita las copias "de respaldo" (mc:Fallback, formato viejo VML) de cuadros de texto y figuras:
      Word usa la versión moderna y así no se duplica ni se rompe la imagen de la huella. */
  function quitarFallbacks(xml) {
    const re = /<mc:Fallback>(?:(?!<mc:Fallback>)[\s\S])*?<\/mc:Fallback>/;
    let prev; do { prev = xml; xml = xml.replace(re, ''); } while (xml !== prev);
    return xml;
  }

  /**
   * Detecta en el machote dónde firma el socio y dónde va su huella, aunque no tenga [FIRMA]:
   *  - una línea de firma (______) seguida de "C. [SOCIO]" o "[SOCIO]"
   *  - un renglón "[SOCIO]" con "Firma" / "Huella" (ej. "[SOCIO] Nombre, Firma y Huella")
   *  - el recuadro de "REGISTRO DE FIRMA"
   * Ahí agrega [FIRMA] y/o [HUELLA] para que se llenen con las imágenes del cuestionario.
   * No toca líneas de firma de otras personas (representante legal, testigos, "Recibí").
   */
  function marcarFirmas(xml, {firma = true, huella = true} = {}) {
    const segs = []; const re = /<w:p[ >][\s\S]*?<\/w:p>/g; let m;
    while ((m = re.exec(xml))) {
      const texto = (m[0].match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join('');
      segs.push({start: m.index, xml: m[0], texto, t: texto.replace(/\s+/g, ' ').trim()});
    }
    const ya = new Set(); const marcas = new Map();   // índice de párrafo -> [FIRMA]/[HUELLA]
    const otraPersona = t => /HOIL|REPRESENTANTE|TESTIGO|RECIB[IÍ]|ADMINISTRADOR/i.test(t);
    const tiene = (i, k) => (marcas.get(i) || '').includes(k) || /\[(FIRMA|HUELLA)[^\]]*\]/i.test(segs[i].t) && segs[i].t.toUpperCase().includes(k);
    const poner = (i, k) => { if (!tiene(i, k)) marcas.set(i, (marcas.get(i) || '') + k); ya.add(i); };
    const sig = (i, n = 3) => { const out = []; for (let j = i + 1; j < segs.length && out.length < n; j++) if (segs[j].t) out.push(j); return out; };
    segs.forEach((s, i) => {
      if (/\[(FIRMA|HUELLA)/i.test(s.t)) return;                                         // el machote ya dice dónde
      // 1) línea de firma del socio
      if (/^_{8,}\.?$/.test(s.t)) { const [j] = sig(i, 1); if (j !== undefined && /\[SOCIO\]/.test(segs[j].t) && !otraPersona(segs[j].t) && j - i <= 4) { if (firma) poner(i, '[FIRMA]'); } return; }
      // 2) "[SOCIO] Nombre, Firma y Huella" en el mismo renglón o en el siguiente
      if (/\[SOCIO\]/.test(s.t) && !otraPersona(s.t) && s.t.length < 90) {
        const [j] = sig(i, 1); const pie = s.t + ' ' + (j !== undefined ? segs[j].t : '');
        if (/firma/i.test(pie) && firma && !ya.has(i - 1)) poner(i, '[FIRMA]');
        if (/huella/i.test(pie) && huella) poner(i, '[HUELLA]');
      }
      // 3) recuadro de registro de firma: primer renglón vacío después del título
      if (/REGISTRO DE FIRMA/i.test(s.t) && firma) { for (let j = i + 1; j < Math.min(segs.length, i + 6); j++) if (!segs[j].t) { poner(j, '[FIRMA]'); break; } }
    });
    if (!marcas.size) return xml;
    let out = '', last = 0;
    segs.forEach((s, i) => {
      if (!marcas.has(i)) return;
      const run = marcas.get(i).split(/(?=\[)/).filter(Boolean).map(k => `<w:r><w:t xml:space="preserve">${k} </w:t></w:r>`).join('') + '<w:r><w:br/></w:r>';
      const ppr = s.xml.match(/^<w:p(?: [^>]*)?>\s*(?:<w:pPr\/>|<w:pPr>(?:(?!<w:p[ >])[\s\S])*?<\/w:pPr>)/);
      const corte = ppr ? ppr[0].length : s.xml.indexOf('>') + 1;
      out += xml.slice(last, s.start) + s.xml.slice(0, corte) + run + s.xml.slice(corte); last = s.start + s.xml.length;
    });
    return out + xml.slice(last);
  }

  /** Lugares donde irían firma y huella (para mostrarlo en Ajustes). */
  async function lugaresImagenes(template) {
    const zip = await JSZip.loadAsync(template); const xml = quitarFallbacks(await zip.file('word/document.xml').async('string'));
    const marcado = marcarFirmas(xml); const t = marcado.replace(/<[^>]+>/g, ' ');
    return {firmas: (t.match(/\[FIRMA[^\]]*\]/gi) || []).length, huellas: (t.match(/\[HUELLA[^\]]*\]/gi) || []).length};
  }

  /**
   * @param template ArrayBuffer/Uint8Array del machote
   * @param valores {'SOCIO': 'PEREZ ...', 'RFC': ...}
   * @param imagenes {'HUELLA': {data: Uint8Array, ext: 'png'|'jpeg', w: px, h: px}}
   * @returns Uint8Array del .docx lleno
   */
  async function llenar(template, valores, imagenes = {}) {
    const zip = await JSZip.loadAsync(template);
    const imgMarks = {}; const imgs = [];
    let n = 0;
    for (const [k, img] of Object.entries(imagenes)) {
      if (!img || !img.data) continue;
      n++; const mark = `§IMG${n}§`; imgMarks[k.toUpperCase()] = mark; imgs.push({mark, img, n});
    }
    const files = Object.keys(zip.files).filter(f => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(f));
    for (const f of files) {
      let xml = await zip.file(f).async('string');
      if (f === 'word/document.xml' && imgs.length) {
        xml = quitarFallbacks(xml);
        xml = marcarFirmas(xml, {firma: !!imgMarks.FIRMA, huella: !!imgMarks.HUELLA});
      }
      zip.file(f, fillXml(xml, valores, f === 'word/document.xml' ? imgMarks : null));
    }
    if (imgs.length) {
      let doc = await zip.file('word/document.xml').async('string');
      let rels = await zip.file('word/_rels/document.xml.rels').async('string');
      let ct = await zip.file('[Content_Types].xml').async('string');
      let docPr = 0;   // cada imagen insertada lleva su propio id
      for (const {mark, img, n: i} of imgs) {
        const ext = img.ext === 'jpg' ? 'jpeg' : img.ext;
        const name = `facol_img${i}.${ext}`; const rid = `rIdFacol${i}`;
        zip.file(`word/media/${name}`, img.data);
        rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/></Relationships>`);
        if (!new RegExp(`Extension="${ext}"`, 'i').test(ct)) ct = ct.replace('<Types ', '<Types ').replace(/(<Types[^>]*>)/, `$1<Default Extension="${ext}" ContentType="image/${ext}"/>`);
        const maxW = img.maxW || 1.6, ratio = (img.h && img.w) ? img.h / img.w : 1;
        const cx = Math.round(maxW * 914400), cy = Math.round(maxW * ratio * 914400);
        doc = doc.replace(new RegExp(`(<w:t[^>]*>)([^<]*)${mark}([^<]*)(</w:t>)`, 'g'),
          (_, o, a, b, c) => `${o}${a}${c}</w:r><w:r>${drawingXml(rid, 9000 + (++docPr), cx, cy)}</w:r><w:r><w:t xml:space="preserve">${b}${c}`);
      }
      if (!/xmlns:wp=/.test(doc)) doc = doc.replace('<w:document ', '<w:document xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ');
      zip.file('word/document.xml', doc); zip.file('word/_rels/document.xml.rels', rels); zip.file('[Content_Types].xml', ct);
    }
    return zip.generateAsync({type: 'uint8array', compression: 'DEFLATE'});
  }

  /** Lista los [CAMPOS] que tiene un machote (para mostrar qué se va a llenar). */
  async function campos(template) {
    const zip = await JSZip.loadAsync(template); const set = new Set();
    for (const f of Object.keys(zip.files).filter(f => /^word\/(document|header\d*|footer\d*)\.xml$/.test(f))) {
      const xml = await zip.file(f).async('string');
      const text = xml.replace(/<\/w:p>/g, '\n').replace(/<[^>]+>/g, '');
      for (const m of text.matchAll(/\[([^\[\]\n]{1,60})\]/g)) set.add(m[1].trim());
    }
    return [...set];
  }

  return {llenar, campos, fillXml, IMG_KEYS, marcarFirmas, quitarFallbacks, lugaresImagenes};
});
