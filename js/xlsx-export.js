/* Genera el .xlsx rellenando la plantilla original (conserva logo, colores,
   fórmulas, formato condicional y configuración de impresión). */
(function () {
  const MAIN_NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const enc = new TextEncoder();
  const dec = new TextDecoder();

  /* ---------- utilidades de cálculo (las mismas reglas del Excel) ---------- */
  const ALL_ITEMS = () => window.SECCIONES.flatMap(s => s.items);
  const TOTAL = () => ALL_ITEMS().length; // 32

  function score(audit) {
    let c = 0, nc = 0;
    for (const it of ALL_ITEMS()) {
      const e = audit.items?.[it.n]?.estado;
      if (e === 'C') c++; else if (e === 'NC') nc++;
    }
    const total = TOTAL();
    return { c, nc, pend: total - c - nc, total, pct: c / total };
  }

  // =SI(F46<=0,25;E4+7;SI(F46<=0,5;E4+14;SI(F46<=0,8;E4+21;E4+30)))
  function diasProxima(pct) {
    if (pct <= 0.25) return 7;
    if (pct <= 0.5) return 14;
    if (pct <= 0.8) return 21;
    return 30;
  }

  function parseISO(iso) {
    if (!iso) return null;
    const [y, m, d] = iso.split('-').map(Number);
    if (!y || !m || !d) return null;
    return { y, m, d };
  }
  function excelSerial(iso) {
    const p = parseISO(iso);
    if (!p) return null;
    return Math.round((Date.UTC(p.y, p.m - 1, p.d) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  function fmtDMY(iso) {
    const p = parseISO(iso);
    if (!p) return '';
    return `${String(p.d).padStart(2, '0')}/${String(p.m).padStart(2, '0')}/${p.y}`;
  }
  function addDaysISO(iso, days) {
    const p = parseISO(iso);
    if (!p) return '';
    const dt = new Date(Date.UTC(p.y, p.m - 1, p.d + days));
    return dt.toISOString().slice(0, 10);
  }
  function proximaAuditoria(audit) {
    return addDaysISO(audit.fecha, diasProxima(score(audit).pct));
  }

  /* ---------- edición de XML ---------- */
  const esc = s => String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    // eliminar caracteres de control no válidos en XML
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

  const colNum = ref => {
    const letters = ref.match(/^[A-Z]+/)[0];
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  };

  function buildCell(ref, sAttr, val) {
    if (val == null || val === '') return `<c r="${ref}"${sAttr}/>`;
    if (typeof val === 'number') return `<c r="${ref}"${sAttr}><v>${val}</v></c>`;
    if (typeof val === 'object' && val.f) {
      const v = val.v == null ? '' : `<v>${val.v}</v>`;
      return `<c r="${ref}"${sAttr}><f>${esc(val.f)}</f>${v}</c>`;
    }
    return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${esc(val)}</t></is></c>`;
  }

  function setCell(xml, ref, val, forceStyle) {
    const re = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`);
    const m = xml.match(re);
    let sAttr = '';
    if (forceStyle != null) sAttr = ` s="${forceStyle}"`;
    else if (m) { const sm = m[1].match(/\ss="(\d+)"/); if (sm) sAttr = ` s="${sm[1]}"`; }
    const cell = buildCell(ref, sAttr, val);
    if (m) return xml.replace(m[0], () => cell);

    // La celda no existe: insertarla en su fila, en orden de columna.
    const rowNum = ref.match(/\d+$/)[0];
    const rowRe = new RegExp(`<row r="${rowNum}"([^>]*?)(/>|>([\\s\\S]*?)</row>)`);
    const rm = xml.match(rowRe);
    if (!rm) {
      // crear fila nueva antes de </sheetData>
      return xml.replace('</sheetData>', () => `<row r="${rowNum}">${cell}</row></sheetData>`);
    }
    if (rm[2] === '/>') return xml.replace(rm[0], () => `<row r="${rowNum}"${rm[1]}>${cell}</row>`);
    const inner = rm[3];
    const cells = inner.match(/<c r="[A-Z]+\d+"[\s\S]*?(?:\/>|<\/c>)/g) || [];
    const target = colNum(ref);
    let newInner = null;
    for (const c of cells) {
      const r = c.match(/r="([A-Z]+\d+)"/)[1];
      if (colNum(r) > target) { newInner = inner.replace(c, () => cell + c); break; }
    }
    if (newInner == null) newInner = inner + cell;
    return xml.replace(rm[0], () => `<row r="${rowNum}"${rm[1]}>${newInner}</row>`);
  }

  /* ---------- fotos en la hoja EVIDENCIAS ---------- */
  // Fila de cada criterio en la hoja EVIDENCIAS (el criterio 32 se agrega en la fila 39).
  function evidenciasRow(n) {
    if (n <= 8) return n + 2;      // 1-8   -> 3-10
    if (n <= 13) return n + 3;     // 9-13  -> 12-16
    if (n <= 19) return n + 4;     // 14-19 -> 18-23
    if (n <= 23) return n + 5;     // 20-23 -> 25-28
    if (n <= 28) return n + 6;     // 24-28 -> 30-34
    return n + 7;                  // 29-32 -> 36-39
  }

  const EMU = 9525;           // EMU por píxel
  const COL_PX = 80;          // ancho por defecto (11.43) de las columnas D..I
  const ROW_PX = 201;         // 150.75 pt
  const SLOT_W = COL_PX * 2;  // 3 fotos por fila, 2 columnas cada una
  const BOX_W = SLOT_W - 12, BOX_H = ROW_PX - 14;

  function anchorXml(idx, rId, rowIdx0, slot, w, h) {
    const k = Math.min(BOX_W / w, BOX_H / h, 1);
    const pw = Math.round(w * k), ph = Math.round(h * k);
    let xPx = slot * SLOT_W + Math.round((SLOT_W - pw) / 2);
    let col = 3 + Math.floor(xPx / COL_PX);
    const colOff = (xPx % COL_PX) * EMU;
    const rowOff = Math.round((ROW_PX - ph) / 2) * EMU;
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${rowIdx0}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:from><xdr:ext cx="${pw * EMU}" cy="${ph * EMU}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${idx + 2}" name="Evidencia ${idx + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${MAIN_NS_R}" r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${pw * EMU}" cy="${ph * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  }

  function b64ToU8(b64) {
    const bin = atob(b64);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
  }

  /**
   * audit: objeto de auditoría
   * photos: [{ item, bytes: Uint8Array, w, h, ext: 'jpeg'|'png' }]
   * -> Uint8Array con el .xlsx
   */
  function buildXlsx(audit, photos) {
    const T = window.XLSX_TEMPLATE;
    const files = {};
    for (const k of Object.keys(T)) files[k] = b64ToU8(T[k]);
    const getT = p => dec.decode(files[p]);
    const setT = (p, s) => { files[p] = enc.encode(s); };

    /* --- Hoja 1: FORMATO 5'S --- */
    let s1 = getT('xl/worksheets/sheet1.xml');
    s1 = setCell(s1, 'C3', `Proceso: ${audit.proceso || ''}`);
    s1 = setCell(s1, 'C4', `Fecha de auditoría: ${fmtDMY(audit.fecha)}`);
    const serial = excelSerial(audit.fecha);
    s1 = setCell(s1, 'E4', serial == null ? '' : serial); // celda auxiliar (oculta) con la fecha real

    const auds = (audit.auditores || []).map(x => x.trim()).filter(Boolean);
    const slots = ['G3', 'G4', 'G5', 'G6'];
    slots.forEach((ref, i) => {
      let v = auds[i] || '';
      if (i === slots.length - 1 && auds.length > slots.length) v = auds.slice(i).join(', ');
      s1 = setCell(s1, ref, v);
    });

    for (const sec of window.SECCIONES) {
      for (const it of sec.items) {
        const d = audit.items?.[it.n] || {};
        const r = it.row;
        s1 = setCell(s1, 'D' + r, d.estado === 'C' ? 1 : '');
        s1 = setCell(s1, 'E' + r, d.estado === 'NC' ? 1 : '');
        s1 = setCell(s1, 'F' + r, d.desc || '');
        s1 = setCell(s1, 'G' + r, d.contramedida || '');
        s1 = setCell(s1, 'H' + r, d.responsable || '');
        s1 = setCell(s1, 'I' + r, d.plazo ? fmtDMY(d.plazo) : '');
      }
    }

    // Valores calculados (Excel igual recalcula al abrir)
    const sc = score(audit);
    const sumRanges = 'SUM(D9:D16,D18:D22,D24:D29,D31:D34,D36:D40,D42:D45)';
    s1 = setCell(s1, 'D46', { f: sumRanges, v: sc.c });
    s1 = setCell(s1, 'E46', { f: sumRanges.replace(/D/g, 'E'), v: sc.nc });
    s1 = setCell(s1, 'F46', { f: 'D46/B46', v: sc.pct });
    s1 = setCell(s1, 'H46', {
      f: 'IF(F46<=0.25,E4+7,IF(F46<=0.5,E4+14,IF(F46<=0.8,E4+21,IF(F46>0.8,E4+30,0))))',
      v: serial == null ? null : serial + diasProxima(sc.pct),
    });
    setT('xl/worksheets/sheet1.xml', s1);

    /* --- Hoja 3: EVIDENCIAS (agrega el criterio 32 y las fotos) --- */
    let s3 = getT('xl/worksheets/sheet3.xml');
    if (!/<row r="39"/.test(s3)) {
      const c32 = window.SECCIONES.flatMap(s => s.items).find(i => i.n === 32);
      // La fila 38 tiene el borde inferior de cierre: pasa a la nueva fila 39,
      // y la 38 toma el estilo "intermedio" de la 37.
      const st = ref => (s3.match(new RegExp(`<c r="${ref}" s="(\\d+)"`)) || [])[1];
      const b37 = st('B37'), c37 = st('C37'), b38 = st('B38'), c38 = st('C38');
      if (b37) s3 = s3.replace(`<c r="B38" s="${b38}"`, `<c r="B38" s="${b37}"`);
      if (c37) s3 = s3.replace(`<c r="C38" s="${c38}"`, `<c r="C38" s="${c37}"`);
      const cells = ['D', 'E', 'F', 'G', 'H', 'I'].map(c => `<c r="${c}39" s="${st('D38') || 54}"/>`).join('');
      const row39 = `<row r="39" spans="2:9" ht="150.75" customHeight="1"><c r="B39" s="${b38 || 20}"><v>32</v></c><c r="C39" s="${c38 || 5}" t="inlineStr"><is><t xml:space="preserve">${esc(c32.texto)}</t></is></c>${cells}</row>`;
      s3 = s3.replace('</sheetData>', () => row39 + '</sheetData>');
      s3 = s3.replace('<dimension ref="B1:I38"/>', '<dimension ref="B1:I39"/>');
      s3 = s3.replace(/<mergeCells count="(\d+)">/, (_, n) => `<mergeCells count="${+n + 1}"><mergeCell ref="D39:I39"/>`);
    }

    const pics = (photos || []).filter(p => p.bytes && p.w && p.h);
    if (pics.length) {
      const perItem = {};
      let anchors = '';
      let rels = '';
      pics.forEach((p, i) => {
        const slot = (perItem[p.item] = (perItem[p.item] ?? -1) + 1);
        if (slot > 2) return; // máximo 3 por criterio en el Excel
        const ext = p.ext === 'png' ? 'png' : 'jpeg';
        const name = `evidencia${i + 1}.${ext}`;
        files['xl/media/' + name] = p.bytes;
        const rId = 'rId' + (i + 1);
        rels += `<Relationship Id="${rId}" Type="${MAIN_NS_R}/image" Target="../media/${name}"/>`;
        anchors += anchorXml(i, rId, evidenciasRow(p.item) - 1, slot, p.w, p.h);
      });
      files['xl/drawings/drawing2.xml'] = enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors}</xdr:wsDr>`);
      files['xl/drawings/_rels/drawing2.xml.rels'] = enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);

      let r3 = getT('xl/worksheets/_rels/sheet3.xml.rels');
      r3 = r3.replace('</Relationships>', `<Relationship Id="rIdEv1" Type="${MAIN_NS_R}/drawing" Target="../drawings/drawing2.xml"/></Relationships>`);
      setT('xl/worksheets/_rels/sheet3.xml.rels', r3);
      s3 = s3.replace(/(<pageSetup[^>]*\/>)/, '$1<drawing r:id="rIdEv1"/>');

      let ct = getT('[Content_Types].xml');
      if (!/Extension="jpeg"/.test(ct)) ct = ct.replace('<Default Extension="png"', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png"');
      ct = ct.replace('</Types>', '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
      setT('[Content_Types].xml', ct);
    }
    setT('xl/worksheets/sheet3.xml', s3);

    /* --- Libro: cálculo automático, sin calcChain --- */
    let wb = getT('xl/workbook.xml');
    wb = wb.replace(/<calcPr[^>]*\/>/, '<calcPr calcId="191028" fullCalcOnLoad="1"/>');
    wb = wb.replace("EVIDENCIAS!$B$1:$I$38", "EVIDENCIAS!$B$1:$I$39");
    setT('xl/workbook.xml', wb);

    let wrels = getT('xl/_rels/workbook.xml.rels');
    wrels = wrels.replace(/<Relationship [^>]*calcChain[^>]*\/>/, '');
    setT('xl/_rels/workbook.xml.rels', wrels);
    let ct2 = getT('[Content_Types].xml');
    ct2 = ct2.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
    setT('[Content_Types].xml', ct2);

    // [Content_Types].xml debe ir primero en el paquete
    const order = Object.keys(files).sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0));
    return window.Zip.zipStore(order.map(name => ({ name, data: files[name] })));
  }

  function fileName(audit) {
    const clean = s => (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '_');
    const proc = clean(audit.proceso) || 'Proceso';
    return `Auditoria_5S_${proc}_${audit.fecha || 'sin-fecha'}.xlsx`;
  }

  window.Xlsx5S = { buildXlsx, fileName, score, diasProxima, proximaAuditoria, fmtDMY, addDaysISO, excelSerial, setCell };
})();
