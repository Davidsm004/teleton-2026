/*!
 * Teletón 2026 · Generador de informes en PowerPoint (.pptx) y PDF con el formato de marca.
 *
 * Uso en el navegador:
 *   <script src="teleton-informes-recursos.js"></script>
 *   <script src="teleton-informes.js"></script>
 *   TeletonInforme.descargar(datos, 'pptx');   // o 'pdf'
 *
 * Uso en Node (servidor):
 *   const TeletonInforme = require('./teleton-informes.js');
 *   const buffer = await TeletonInforme.generar(datos, 'pdf');   // Buffer listo para enviar o guardar
 *
 * Todo lo que se ve en el informe (títulos, cifras, tablas, gráficos) sale de "datos":
 * cambiando los datos, el archivo descargado cambia solo. Ver el formato al final de este archivo.
 */
(function (raiz, fabrica) {
  if (typeof module !== 'undefined' && module.exports) module.exports = fabrica(true);
  else raiz.TeletonInforme = fabrica(false);
})(typeof self !== 'undefined' ? self : this, function (enNode) {
  'use strict';

  // ===== Marca =====
  const COL = { vino: 'C800A1', naranja: 'FE5000', rosa: 'E2235F', amarillo: 'FF9E1B', vinoOsc: '5A0049', vinoMed: '7A005F', tinta: '22001C', gris: '6E5F6B', crema: 'FFF5F0', linea: 'F6E0E4', blanco: 'FFFFFF', durazno: 'FFE7D6' };
  const SERIES = [COL.vino, COL.naranja, COL.amarillo, COL.vinoOsc, COL.rosa, COL.vinoMed];
  const TARJETAS = [COL.vino, COL.naranja, COL.vinoOsc, COL.rosa];
  const W = 13.333, H = 7.5;                     // lienzo 16:9 en pulgadas
  const F_TIT = 'Nexa Heavy', F_TXT = 'Nexa Extra Light';
  const FILAS_POR_PAGINA = 9;

  // ===== Utilidades =====
  const texto = v => (v === null || v === undefined) ? '' : String(v);
  const grupo = n => Math.round(Math.abs(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  const fmt = {
    colones: n => (n < 0 ? '-' : '') + '₡' + Math.abs(Math.round(n * 100) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),   // ₡1,340,899.99
    numero: n => (n < 0 ? '-' : '') + grupo(n),
    porcentaje: n => (Math.round(n * 10) / 10).toString().replace('.', ',') + '%',
    millones: n => '₡' + (Math.round(n / 1e5) / 10).toString().replace('.', ',') + ' M'
  };
  const formatear = (v, f) => (typeof v === 'number' && f && fmt[f]) ? fmt[f](v) : texto(v);
  const esNumero = v => typeof v === 'number' || /^[₡$-]?\s?[\d\s.,]+%?\s?M?$/.test(texto(v).trim());
  // ancho aproximado de un texto en pulgadas (Nexa): sirve para achicar títulos largos igual en PPTX y PDF
  let medidor = null;                             // en PDF se mide con la fuente real; en PPTX, aproximado
  const ancho = (s, pt, gruesa) => medidor ? medidor(texto(s), pt, gruesa) : texto(s).length * pt / 72 * (gruesa ? 0.64 : 0.52);
  function lineas(s, pt, w, gruesa) {
    const pal = texto(s).split(/\s+/); let n = 1, act = 0;
    pal.forEach(p => { const a = ancho(p + ' ', pt, gruesa); if (act + a > w && act > 0) { n++; act = a; } else act += a; });
    return n;
  }
  function ajustar(s, pt, w, h, gruesa, min) {      // baja el tamaño hasta que el texto quepa en la caja
    let t = pt; min = min || 12;
    while (t > min && lineas(s, t, w, gruesa) * t / 72 * 1.18 > h) t -= 1;
    return t;
  }
  const unaLinea = (s, pt, w, gruesa, min) => { let t = pt; while (t > (min || 12) && ancho(s, t, gruesa) > w) t -= 1; return t; };   // cifras: nunca se parten
  const slug = s => texto(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w]+/g, '-').replace(/^-|-$/g, '').toLowerCase();
  const recursos = () => {
    const r = (typeof TELETON_RECURSOS !== 'undefined') ? TELETON_RECURSOS : (enNode ? require('./teleton-informes-recursos.js') : null);
    if (!r) throw new Error('Falta cargar teleton-informes-recursos.js');
    return r;
  };

  // ===== Carga de librerías =====
  // Librerías alojadas en el mismo sitio (sin depender de servidores externos)
  const LIB = () => (typeof window !== 'undefined' && window.TL_LIB_BASE) || 'lib/';
  const CDN = { get pptx() { return LIB() + 'pptxgen.bundle.js'; }, get pdf() { return LIB() + 'jspdf.umd.min.js'; } };
  function cargarScript(src) {
    return new Promise((ok, mal) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => mal(new Error('No se pudo cargar ' + src)); document.head.appendChild(s); });
  }
  async function libPptx() {
    if (enNode) return require('pptxgenjs');
    if (!window.PptxGenJS) await cargarScript(CDN.pptx);
    return window.PptxGenJS;
  }
  async function libPdf() {
    if (enNode) return require('jspdf').jsPDF;
    if (!(window.jspdf && window.jspdf.jsPDF)) await cargarScript(CDN.pdf);
    return window.jspdf.jsPDF;
  }

  // =====================================================================
  //  Lienzo: las mismas instrucciones de dibujo sirven para PPTX y PDF
  // =====================================================================
  function lienzoPptx(pres) {
    let s = null;
    return {
      nueva(fondo, numero) {
        s = pres.addSlide({ masterName: 'TT_' + fondo });
        if (numero) s.addText(String(numero), { x: 12.25, y: 6.92, w: 0.7, h: 0.32, fontSize: 10, color: COL.vinoMed, fontFace: F_TXT, align: 'right', margin: 0, isTextBox: true });
      },
      notas(t) { if (t) s.addNotes(t); },
      img(data, x, y, w, h, opc) { s.addImage(Object.assign({ data, x, y, w, h }, opc || {})); },
      rect(x, y, w, h, color, radio, sombra, rot) {
        s.addShape(radio ? 'roundRect' : 'rect', Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' }, rotate: rot || 0 },
          radio ? { rectRadius: radio } : {}, sombra ? { shadow: { type: 'outer', color: '3D0032', opacity: 0.3, blur: 10, offset: 5, angle: 90 } } : {}));
      },
      txt(t, x, y, w, h, o) {
        s.addText(Array.isArray(t) ? t.map(r => ({ text: r.text, options: Object.assign({}, r.options || {}) })) : texto(t), {
          x, y, w, h, isTextBox: true, margin: o.margen || 0, fontSize: o.pt, color: o.color, fontFace: o.gruesa ? F_TIT : F_TXT, bold: false,
          align: o.alinear || 'left', valign: o.v || 'top', rotate: o.rot || 0, paraSpaceAfter: o.espacio || 0, lineSpacingMultiple: 1.05
        });
      },
      vinetas(items, x, y, w, h, pt, color) {
        s.addText(items.map((it, i) => ({ text: texto(it), options: { bullet: { indent: 18 }, breakLine: i < items.length - 1 } })),
          { x, y, w, h, isTextBox: true, margin: 0, fontSize: pt, color, fontFace: F_TXT, valign: 'top', paraSpaceAfter: 10 });
      },
      tabla(cab, filas, x, y, w, anchos, o) {
        const celda = (v, extra) => ({ text: texto(v), options: Object.assign({ fontFace: F_TXT, fontSize: o.pt, color: COL.tinta, valign: 'middle', margin: [3, 8, 3, 8] }, extra) });
        const rows = [cab.map((c, i) => celda(c, { fontFace: F_TIT, color: COL.blanco, fill: { color: COL.vinoOsc }, align: o.alin[i] }))];
        filas.forEach((f, r) => rows.push(f.cells.map((c, i) => celda(c, Object.assign({ align: o.alin[i], fill: { color: f.total ? COL.durazno : (r % 2 ? COL.crema : COL.blanco) } }, f.total ? { fontFace: F_TIT, color: COL.vinoOsc } : {})))));
        s.addTable(rows, { x, y, w, colW: anchos, rowH: o.alto, border: { type: 'solid', pt: 0.75, color: COL.linea } });
      },
      grafico(g, x, y, w, h) {
        const tipo = { barra: 'bar', columna: 'bar', linea: 'line', dona: 'doughnut' }[g.tipo] || 'bar';
        const data = (g.series || []).map(se => ({ name: texto(se.nombre), labels: g.etiquetas.map(texto), values: se.valores.map(Number) }));
        const f = g.formato === 'colones' ? '"₡"#,##0.00' : g.formato === 'millones' ? '"₡"#,##0.0,," M"' : g.formato === 'porcentaje' ? '0.0"%"' : '#,##0';
        const opc = {
          x, y, w, h, chartColors: (tipo === 'doughnut' ? SERIES.concat(SERIES) : SERIES).slice(0, tipo === 'doughnut' ? g.etiquetas.length : data.length),
          barDir: g.tipo === 'barra' ? 'bar' : 'col', barGapWidthPct: 60,
          showLegend: data.length > 1 || tipo === 'doughnut', legendPos: 'b', legendFontFace: F_TXT, legendFontSize: 12, legendColor: COL.tinta,
          showValue: tipo !== 'line', dataLabelFormatCode: f, dataLabelFontFace: F_TIT, dataLabelFontSize: 11, dataLabelColor: tipo === 'doughnut' ? COL.blanco : COL.vinoOsc,
          dataLabelPosition: tipo === 'doughnut' ? 'ctr' : 'outEnd', valAxisLabelFormatCode: f,
          catAxisLabelFontFace: F_TXT, valAxisLabelFontFace: F_TXT, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelColor: COL.gris, valAxisLabelColor: COL.gris,
          valGridLine: { color: COL.linea, size: 1 }, catGridLine: { style: 'none' }, holeSize: 58, lineSize: 3, lineDataSymbolSize: 9
        };
        if (tipo === 'doughnut') { opc.showPercent = true; opc.showValue = false; opc.dataLabelFormatCode = '0.0%'; }
        s.addChart(tipo, data, opc);
      }
    };
  }

  function lienzoPdf(doc, R) {
    const rgb = h => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
    const pt2in = pt => pt / 72;
    let primera = true;
    function fuente(gruesa, pt, color) { doc.setFont(gruesa ? 'NexaHeavy' : 'NexaExtraLight', 'normal'); doc.setFontSize(pt); doc.setTextColor.apply(doc, rgb(color)); }
    function partir(t, pt, w, gruesa) { fuente(gruesa, pt, COL.tinta); return doc.splitTextToSize(texto(t), w); }
    const api = {
      medir(s, pt, gruesa) { fuente(gruesa, pt, COL.tinta); return doc.getTextWidth(s); },
      nueva(fondo, numero) {
        if (!primera) doc.addPage([W, H], 'landscape'); primera = false;
        doc.addImage(R.fondos[fondo === 'portada' || fondo === 'seccion' || fondo === 'imagen' || fondo === 'cierre' ? fondo : 'claro'], 'JPEG', 0, 0, W, H, 'fondo-' + fondo, 'FAST');
        if (numero) { fuente(false, 10, COL.vinoMed); doc.text(String(numero), 12.95, 7.13, { align: 'right' }); }
      },
      notas() {},
      img(data, x, y, w, h) { doc.addImage(data, data.indexOf('image/png') > 0 ? 'PNG' : 'JPEG', x, y, w, h, undefined, 'FAST'); },
      rect(x, y, w, h, color, radio, sombra, rot) {
        if (rot) {                                        // rectángulo inclinado (etiquetas y bloques de nombre)
          const a = rot * Math.PI / 180, cx = x + w / 2, cy = y + h / 2, c = Math.cos(a), s = Math.sin(a);
          const p = [[-w / 2, -h / 2], [w / 2, -h / 2], [w / 2, h / 2], [-w / 2, h / 2]].map(([px, py]) => [cx + px * c - py * s, cy + px * s + py * c]);
          if (sombra) { doc.setFillColor(60, 0, 40); doc.setGState(new doc.GState({ opacity: 0.18 })); doc.lines(p.slice(1).map((q, i) => [q[0] - p[i][0], q[1] - p[i][1]]), p[0][0] + 0.04, p[0][1] + 0.06, [1, 1], 'F', true); doc.setGState(new doc.GState({ opacity: 1 })); }
          doc.setFillColor.apply(doc, rgb(color));
          doc.lines(p.slice(1).map((q, i) => [q[0] - p[i][0], q[1] - p[i][1]]), p[0][0], p[0][1], [1, 1], 'F', true);
          return;
        }
        if (sombra) { doc.setFillColor(60, 0, 40); doc.setGState(new doc.GState({ opacity: 0.16 })); radio ? doc.roundedRect(x + 0.03, y + 0.07, w, h, radio, radio, 'F') : doc.rect(x + 0.03, y + 0.07, w, h, 'F'); doc.setGState(new doc.GState({ opacity: 1 })); }
        doc.setFillColor.apply(doc, rgb(color));
        radio ? doc.roundedRect(x, y, w, h, radio, radio, 'F') : doc.rect(x, y, w, h, 'F');
      },
      txt(t, x, y, w, h, o) {
        const runs = Array.isArray(t) ? t : [{ text: t }];
        // varias "corridas" (por ejemplo "¡Hagamos que" + salto + "suene!"): cada una en su línea si trae breakLine
        let filas = [], act = '';
        runs.forEach(r => { act += texto(r.text); if (r.options && r.options.breakLine) { filas.push(act); act = ''; } });
        if (act) filas.push(act);
        let ls = []; filas.forEach(f => { ls = ls.concat(partir(f, o.pt, w - 2 * (o.margen || 0) / 72, o.gruesa)); });
        const lh = pt2in(o.pt) * 1.12, alto = ls.length * lh;
        let ty = o.v === 'middle' ? y + (h - alto) / 2 : o.v === 'bottom' ? y + h - alto : y;
        fuente(o.gruesa, o.pt, o.color);
        const ax = o.alinear === 'center' ? x + w / 2 : o.alinear === 'right' ? x + w - (o.margen || 0) / 72 : x + (o.margen || 0) / 72;
        if (o.rot) {
          const a = -o.rot, cx = x + w / 2, cy = y + h / 2;
          doc.text(ls.join(' '), cx, cy + pt2in(o.pt) * 0.32, { align: 'center', angle: a });
          return;
        }
        ls.forEach((l, i) => doc.text(l, ax, ty + lh * i + pt2in(o.pt) * 0.82, { align: o.alinear || 'left' }));
      },
      vinetas(items, x, y, w, h, pt, color) {
        let ty = y; const lh = pt2in(pt) * 1.15;
        items.forEach(it => {
          const ls = partir(it, pt, w - 0.3, false);
          doc.setFillColor.apply(doc, rgb(COL.naranja)); doc.circle(x + 0.08, ty + lh * 0.55, 0.055, 'F');
          fuente(false, pt, color); ls.forEach((l, i) => doc.text(l, x + 0.3, ty + lh * i + pt2in(pt) * 0.82));
          ty += ls.length * lh + 0.14;
        });
      },
      tabla(cab, filas, x, y, w, anchos, o) {
        let ty = y;
        const fila = (cells, bg, gruesa, color) => {
          let tx = x;
          doc.setFillColor.apply(doc, rgb(bg)); doc.rect(x, ty, w, o.alto, 'F');
          doc.setDrawColor.apply(doc, rgb(COL.linea)); doc.setLineWidth(0.01); doc.line(x, ty + o.alto, x + w, ty + o.alto);
          cells.forEach((c, i) => {
            fuente(gruesa, o.pt, color);
            let s = texto(c); while (s.length > 1 && doc.getTextWidth(s) > anchos[i] - 0.22) s = s.slice(0, -2) + '…';
            const al = o.alin[i] || 'left', px = al === 'right' ? tx + anchos[i] - 0.11 : al === 'center' ? tx + anchos[i] / 2 : tx + 0.11;
            doc.text(s, px, ty + o.alto / 2 + pt2in(o.pt) * 0.32, { align: al });
            tx += anchos[i];
          });
          ty += o.alto;
        };
        fila(cab, COL.vinoOsc, true, COL.blanco);
        filas.forEach((f, r) => fila(f.cells, f.total ? COL.durazno : (r % 2 ? COL.crema : COL.blanco), !!f.total, f.total ? COL.vinoOsc : COL.tinta));
      },
      grafico(g, x, y, w, h) { dibujarGraficoPdf(doc, g, x, y, w, h, fuente, rgb); }
    };
    return api;
  }

  function dibujarGraficoPdf(doc, g, x, y, w, h, fuente, rgb) {
    const series = g.series || [], et = g.etiquetas || [], f = v => formatear(v, g.formato);
    const leyenda = (nombres, colores, ly) => {
      let tot = 0; fuente(false, 12, COL.tinta); nombres.forEach(n => { tot += 0.3 + doc.getTextWidth(texto(n)) + 0.35; });
      let lx = x + (w - tot) / 2;
      nombres.forEach((n, i) => { doc.setFillColor.apply(doc, rgb(colores[i % colores.length])); doc.roundedRect(lx, ly - 0.13, 0.2, 0.2, 0.04, 0.04, 'F'); fuente(false, 12, COL.tinta); doc.text(texto(n), lx + 0.3, ly + 0.03); lx += 0.3 + doc.getTextWidth(texto(n)) + 0.35; });
    };
    if (g.tipo === 'dona') {
      const vals = (series[0] || { valores: [] }).valores.map(Number), tot = vals.reduce((a, b) => a + b, 0) || 1;
      const r = Math.min(w * 0.42, h - 0.9) / 2, cx = x + w / 2, cy = y + (h - 0.6) / 2, ri = r * 0.58;
      let a0 = -Math.PI / 2;
      vals.forEach((v, i) => {
        const a1 = a0 + v / tot * Math.PI * 2, pasos = Math.max(4, Math.ceil((a1 - a0) / 0.05)), pts = [];
        for (let k = 0; k <= pasos; k++) { const a = a0 + (a1 - a0) * k / pasos; pts.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]); }
        for (let k = pasos; k >= 0; k--) { const a = a0 + (a1 - a0) * k / pasos; pts.push([cx + ri * Math.cos(a), cy + ri * Math.sin(a)]); }
        doc.setFillColor.apply(doc, rgb(SERIES[i % SERIES.length]));
        doc.lines(pts.slice(1).map((q, j) => [q[0] - pts[j][0], q[1] - pts[j][1]]), pts[0][0], pts[0][1], [1, 1], 'F', true);
        const am = (a0 + a1) / 2, rm = (r + ri) / 2;
        if (v / tot > 0.04) { fuente(true, 12, COL.blanco); doc.text(fmt.porcentaje(v / tot * 100), cx + rm * Math.cos(am), cy + rm * Math.sin(am) + 0.06, { align: 'center' }); }
        a0 = a1;
      });
      if (g.centro) { fuente(true, 22, COL.vinoOsc); doc.text(texto(g.centro), cx, cy + 0.1, { align: 'center' }); }
      leyenda(et, SERIES, y + h - 0.15);
      return;
    }
    const todos = series.reduce((a, s) => a.concat(s.valores.map(Number)), []);
    let max = Math.max.apply(null, todos.concat([0])) || 1;
    const p = Math.pow(10, Math.floor(Math.log10(max))); max = Math.ceil(max / p * (max / p < 3 ? 2 : 1)) * p / (max / p < 3 ? 2 : 1);
    const izq = 1.35, abajo = series.length > 1 ? 0.95 : 0.55, gx = x + izq, gy = y + 0.2, gw = w - izq - 0.1, gh = h - abajo - 0.2;
    doc.setDrawColor.apply(doc, rgb(COL.linea)); doc.setLineWidth(0.012);
    for (let k = 0; k <= 4; k++) {
      const yy = gy + gh - gh * k / 4; doc.line(gx, yy, gx + gw, yy);
      fuente(false, 11, COL.gris); doc.text(f(max * k / 4), gx - 0.12, yy + 0.05, { align: 'right' });
    }
    const n = et.length, banda = gw / Math.max(1, n);
    et.forEach((e, i) => { fuente(false, 12, COL.gris); doc.text(doc.splitTextToSize(texto(e), banda - 0.1), gx + banda * (i + 0.5), gy + gh + 0.28, { align: 'center' }); });
    if (g.tipo === 'linea') {
      series.forEach((s, k) => {
        const c = SERIES[k % SERIES.length], pts = s.valores.map((v, i) => [gx + banda * (i + 0.5), gy + gh - gh * Number(v) / max]);
        doc.setDrawColor.apply(doc, rgb(c)); doc.setLineWidth(0.04);
        for (let i = 1; i < pts.length; i++) doc.line(pts[i - 1][0], pts[i - 1][1], pts[i][0], pts[i][1]);
        doc.setFillColor.apply(doc, rgb(c)); pts.forEach(q => doc.circle(q[0], q[1], 0.06, 'F'));
        if (series.length === 1) pts.forEach((q, i) => { fuente(true, 11, COL.vinoOsc); doc.text(f(s.valores[i]), q[0], q[1] - 0.14, { align: 'center' }); });
      });
    } else {
      const bw = banda * 0.62 / series.length;
      series.forEach((s, k) => s.valores.forEach((v, i) => {
        const bh = gh * Number(v) / max, bx = gx + banda * i + banda * 0.19 + bw * k;
        doc.setFillColor.apply(doc, rgb(SERIES[k % SERIES.length])); doc.roundedRect(bx, gy + gh - bh, bw, bh, Math.min(0.06, bw / 3), Math.min(0.06, bw / 3), 'F');
        if (bh > 0.01) doc.rect(bx, gy + gh - Math.min(bh, 0.07), bw, Math.min(bh, 0.07), 'F');
        fuente(true, series.length > 2 ? 9 : 11, COL.vinoOsc); doc.text(f(v), bx + bw / 2, gy + gh - bh - 0.08, { align: 'center' });
      }));
    }
    if (series.length > 1) leyenda(series.map(s => s.nombre), SERIES, y + h - 0.15);
  }

  // =====================================================================
  //  Diapositivas / páginas (mismo diseño en ambos formatos)
  // =====================================================================
  function construir(L, datos, R) {
    medidor = L.medir || null;
    let pagina = 0;
    const tituloClaro = (t) => { const pt = ajustar(t, 36, 11.2, 1.05, true, 22); L.txt(t, 0.7, 0.45, 11.2, 1.05, { pt, color: COL.vinoOsc, gruesa: true, v: 'middle' }); };
    const pie = () => L.img(R.logoColor, 0.7, 6.75, 1.2, 0.556);
    const claro = (t) => { pagina++; L.nueva('claro', pagina + 1); tituloClaro(t); pie(); };

    // --- Portada ---
    L.nueva('portada');
    L.img(R.logoBlanco, 0.7, 0.55, 2.2, 1.02);
    L.img(R.mascota, 9.35, 0.9, 2.96, 6.2);
    const tit = texto(datos.titulo || 'Informe Teletón 2026');
    L.txt(tit, 0.7, 1.9, 8.3, 2.4, { pt: ajustar(tit, 54, 8.3, 2.4, true, 30), color: COL.blanco, gruesa: true, v: 'bottom' });
    if (datos.subtitulo) L.txt(datos.subtitulo, 0.7, 4.45, 8.0, 1.1, { pt: ajustar(datos.subtitulo, 22, 8, 1.1, false, 14), color: COL.blanco });
    const pill = texto(datos.fecha || new Date().toLocaleDateString('es-CR', { day: 'numeric', month: 'long', year: 'numeric' }));
    const pw = Math.min(8, ancho(pill, 18, true) + 0.5);
    L.rect(0.7, 5.9, pw, 0.6, COL.blanco, 0.3, true);
    L.txt(pill, 0.7, 5.9, pw, 0.6, { pt: 18, color: COL.vinoOsc, gruesa: true, alinear: 'center', v: 'middle' });
    L.notas(datos.notas);

    (datos.secciones || []).forEach(sec => {
      const t = texto(sec.titulo);
      switch (sec.tipo) {
        case 'seccion': {
          pagina++; L.nueva('seccion');
          if (sec.numero) L.txt(sec.numero, 0.9, 1.3, 5, 1.7, { pt: 96, color: COL.amarillo, gruesa: true, v: 'bottom' });
          L.txt(t, 0.9, 3.15, 10.5, 1.6, { pt: ajustar(t, 48, 10.5, 1.6, true, 28), color: COL.blanco, gruesa: true });
          if (sec.texto) L.txt(sec.texto, 0.9, 4.85, 8.5, 1.1, { pt: 20, color: COL.blanco });
          L.img(R.logoBlanco, 11.0, 6.55, 1.6, 0.742);
          break;
        }
        case 'texto': {
          claro(t);
          if (sec.parrafo) L.txt(sec.parrafo, 0.7, 1.7, 11.9, 1.2, { pt: 20, color: COL.tinta });
          if (sec.puntos) L.vinetas(sec.puntos, 0.7, sec.parrafo ? 3.0 : 1.75, 11.9, 4.8, ajustar(sec.puntos.join(' '), 20, 11.9 * 2.2, 4.8, false, 14), COL.tinta);
          break;
        }
        case 'cifras': {
          claro(t);
          const cs = (sec.cifras || []).slice(0, 4), n = cs.length || 1, gap = 0.3, cw = (11.9 - gap * (n - 1)) / n;
          const y0 = sec.nota ? 1.8 : 2.0;
          cs.forEach((c, i) => {
            const x = 0.7 + i * (cw + gap), val = formatear(c.valor, c.formato), col = c.color || TARJETAS[i % TARJETAS.length];
            L.rect(x, y0, cw, 3.6, col, 0.25, true);
            L.txt(val, x + 0.2, y0 + 0.45, cw - 0.4, 1.5, { pt: unaLinea(val, n > 3 ? 54 : 66, cw - 0.5, true, 20), color: COL.blanco, gruesa: true, alinear: 'center', v: 'middle' });
            L.txt(c.etiqueta, x + 0.25, y0 + 2.05, cw - 0.5, 1.3, { pt: 18, color: COL.blanco, alinear: 'center' });
          });
          if (sec.nota) L.txt(sec.nota, 0.7, 5.75, 11.9, 0.7, { pt: 14, color: COL.gris });
          break;
        }
        case 'progreso': {
          claro(t);
          const meta = Number(sec.meta) || 1, act = Number(sec.recaudado) || 0, pct = Math.max(0, Math.min(1, act / meta));
          L.txt([{ text: fmt.colones(act) }], 0.7, 1.75, 11.9, 1.2, { pt: unaLinea(fmt.colones(act), 66, 11.8, true, 30), color: COL.vino, gruesa: true });
          L.txt('recaudados de una meta de ' + fmt.colones(meta), 0.7, 2.95, 11.9, 0.6, { pt: 20, color: COL.tinta });
          L.rect(0.7, 3.85, 11.9, 0.62, COL.linea, 0.31);
          if (pct > 0) L.rect(0.7, 3.85, Math.max(0.62, 11.9 * pct), 0.62, COL.naranja, 0.31);
          const etq = fmt.porcentaje(pct * 100);
          const ex = Math.min(0.7 + 11.9 * pct - 1.3, 11.3);
          L.txt(etq, Math.max(0.75, ex), 3.85, 1.2, 0.62, { pt: 18, color: pct > 0.12 ? COL.blanco : COL.vinoOsc, gruesa: true, alinear: pct > 0.12 ? 'right' : 'left', v: 'middle' });
          const falta = Math.max(0, meta - act);
          const extras = (sec.cifras || [{ valor: falta, formato: 'colones', etiqueta: 'faltan para la meta' }]).slice(0, 3);
          extras.forEach((c, i) => {
            const cw = (11.9 - 0.3 * (extras.length - 1)) / extras.length, x = 0.7 + i * (cw + 0.3), v = formatear(c.valor, c.formato);
            L.rect(x, 4.95, cw, 1.45, COL.crema, 0.2);
            L.txt(v, x + 0.25, 5.05, cw - 0.5, 0.75, { pt: unaLinea(v, 32, cw - 0.55, true, 16), color: TARJETAS[i % TARJETAS.length], gruesa: true, v: 'middle' });
            L.txt(c.etiqueta, x + 0.25, 5.8, cw - 0.5, 0.5, { pt: 14, color: COL.tinta });
          });
          break;
        }
        case 'tabla': {
          const cols = sec.columnas || [], filas = (sec.filas || []).map(f => ({ cells: f.map((v, i) => formatear(v, (sec.formatos || [])[i])) }));
          if (sec.totales) filas.push({ cells: sec.totales.map((v, i) => formatear(v, (sec.formatos || [])[i])), total: true });
          const alin = cols.map((c, i) => (sec.alinear && sec.alinear[i]) || (filas.length && esNumero(filas[0].cells[i]) ? 'right' : 'left'));
          const pesos = sec.anchos || cols.map((c, i) => Math.max(ancho(c, 14, true), ...filas.map(f => ancho(f.cells[i], 14, false))) + 0.4);
          const sum = pesos.reduce((a, b) => a + b, 0), anchos = pesos.map(p => p / sum * 11.9);
          const porPag = sec.filasPorPagina || FILAS_POR_PAGINA, paginas = Math.max(1, Math.ceil(filas.length / porPag));
          for (let p = 0; p < paginas; p++) {
            claro(t + (paginas > 1 ? ' (' + (p + 1) + '/' + paginas + ')' : ''));
            const trozo = filas.slice(p * porPag, (p + 1) * porPag);
            L.tabla(cols, trozo, 0.7, 1.75, 11.9, anchos, { pt: 14, alto: 0.46, alin });
            if (sec.nota && p === paginas - 1) L.txt(sec.nota, 0.7, 1.75 + 0.46 * (trozo.length + 1) + 0.2, 11.9, 0.6, { pt: 12, color: COL.gris });
          }
          break;
        }
        case 'grafico': {
          claro(t);
          const conTexto = !!sec.comentario;
          L.grafico(Object.assign({}, sec, { tipo: sec.grafico || 'columna' }), 0.7, 1.6, conTexto ? 8.2 : 11.9, 4.95);
          if (conTexto) {
            L.rect(9.2, 1.75, 3.4, 4.6, COL.crema, 0.25);
            L.txt(sec.comentario, 9.45, 2.0, 2.9, 4.1, { pt: ajustar(sec.comentario, 18, 2.9 * 1.6, 4.1, false, 12), color: COL.tinta });
          }
          break;
        }
        case 'columnas': {
          claro(t);
          [sec.izquierda, sec.derecha].forEach((c, i) => {
            if (!c) return; const x = i ? 6.9 : 0.7;
            L.txt(c.titulo, x, 1.75, 5.75, 0.55, { pt: 22, color: i ? COL.naranja : COL.vino, gruesa: true });
            if (c.puntos) L.vinetas(c.puntos, x, 2.45, 5.75, 4.1, 18, COL.tinta);
            else if (c.texto) L.txt(c.texto, x, 2.45, 5.75, 4.1, { pt: 18, color: COL.tinta });
          });
          break;
        }
        case 'afiche': {
          pagina++; L.nueva('imagen');
          L.rect(0.85, 0.95, 4.9, 5.6, COL.blanco, 0.25, true);
          if (sec.foto) L.img(sec.foto, 1.0, 1.1, 4.6, 5.3, { sizing: { type: 'cover', w: 4.6, h: 5.3 } });
          else L.img(R.mascota, 1.95, 1.2, 2.47, 5.17);
          if (sec.etiqueta) { const ew = Math.min(4, ancho(sec.etiqueta, 14, true) + 0.5); L.rect(0.6, 0.62, ew, 0.5, COL.amarillo, 0, true, -5); L.txt(texto(sec.etiqueta).toUpperCase(), 0.6, 0.62, ew, 0.5, { pt: 14, color: '3D0032', gruesa: true, alinear: 'center', v: 'middle', rot: -5 }); }
          L.txt(t, 6.5, 1.2, 6.1, 1.7, { pt: ajustar(t, 44, 6.1, 1.7, true, 26), color: COL.blanco, gruesa: true, v: 'bottom' });
          if (sec.texto) L.txt(sec.texto, 6.5, 3.05, 6.0, 2.2, { pt: ajustar(sec.texto, 18, 6 * 2.5, 2.2, false, 13), color: COL.blanco });
          if (sec.dato) {
            const d1 = texto(sec.dato), w1 = Math.min(6, ancho(d1, 26, true) + 0.5);
            L.rect(6.5, 5.35, w1, 0.62, COL.blanco, 0, true, -2);
            L.txt(d1, 6.5, 5.35, w1, 0.62, { pt: 26, color: COL.vino, gruesa: true, alinear: 'center', v: 'middle', rot: -2 });
          }
          L.img(R.logoBlanco, 11.0, 6.55, 1.6, 0.742);
          break;
        }
        default: throw new Error('Tipo de sección desconocido: ' + sec.tipo);
      }
      L.notas(sec.notas);
    });

    // --- Cierre ---
    const ci = datos.cierre || {};
    L.nueva('cierre');
    L.img(R.logoBlanco, 5.57, 0.6, 2.2, 1.02);
    const ct = texto(ci.titulo || '¡Hagamos que suene!');
    L.txt(ct, 1.2, 2.0, 10.9, 2.6, { pt: ajustar(ct, 60, 10.9, 2.6, true, 30), color: COL.blanco, gruesa: true, alinear: 'center', v: 'middle' });
    L.txt(ci.texto || 'Gracias por sumarte a la Teletón 2026', 2.2, 4.75, 8.9, 0.9, { pt: 22, color: COL.blanco, alinear: 'center' });
    const cc = texto(ci.contacto || 'teletoncr.com'), cw = Math.min(8, ancho(cc, 18, true) + 0.7);
    L.rect((W - cw) / 2, 5.9, cw, 0.6, COL.blanco, 0.3, true);
    L.txt(cc, (W - cw) / 2, 5.9, cw, 0.6, { pt: 18, color: COL.vino, gruesa: true, alinear: 'center', v: 'middle' });
  }

  // =====================================================================
  //  API pública
  // =====================================================================
  async function generar(datos, formato) {
    const R = recursos(); formato = (formato || 'pdf').toLowerCase();
    if (formato === 'pptx' || formato === 'powerpoint') {
      const PptxGenJS = await libPptx(), pres = new PptxGenJS();
      pres.layout = 'LAYOUT_WIDE'; pres.title = texto(datos.titulo); pres.author = 'Teletón Costa Rica';
      pres.theme = { headFontFace: F_TIT, bodyFontFace: F_TXT };
      ['portada', 'seccion', 'claro', 'imagen', 'cierre'].forEach(k => pres.defineSlideMaster({ title: 'TT_' + k, background: { data: R.fondos[k] } }));
      construir(lienzoPptx(pres), datos, R);
      return pres.write({ outputType: enNode ? 'nodebuffer' : 'blob' });
    }
    if (formato === 'pdf') {
      const jsPDF = await libPdf(), doc = new jsPDF({ orientation: 'landscape', unit: 'in', format: [W, H], compress: true });
      doc.addFileToVFS('NexaHeavy.ttf', R.fuentes.NexaHeavy); doc.addFont('NexaHeavy.ttf', 'NexaHeavy', 'normal');
      doc.addFileToVFS('NexaExtraLight.ttf', R.fuentes.NexaExtraLight); doc.addFont('NexaExtraLight.ttf', 'NexaExtraLight', 'normal');
      doc.setProperties({ title: texto(datos.titulo), author: 'Teletón Costa Rica', creator: 'Teletón 2026' });
      construir(lienzoPdf(doc, R), datos, R);
      return enNode ? Buffer.from(doc.output('arraybuffer')) : doc.output('blob');
    }
    throw new Error('Formato no soportado: ' + formato + ' (usá "pptx" o "pdf")');
  }

  function nombreArchivo(datos, formato) {
    const base = datos.archivo || (slug(datos.titulo || 'informe-teleton') + '-' + new Date().toISOString().slice(0, 10));
    return base.replace(/\.(pptx|pdf)$/i, '') + (formato === 'pdf' ? '.pdf' : '.pptx');
  }

  async function descargar(datos, formato) {
    formato = (formato || 'pdf').toLowerCase() === 'pdf' ? 'pdf' : 'pptx';
    const blob = await generar(datos, formato), url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = nombreArchivo(datos, formato); document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  return { generar, descargar, nombreArchivo, formato: fmt };
});

/* =======================================================================
 * FORMATO DE LOS DATOS
 * -----------------------------------------------------------------------
 * {
 *   titulo: 'Informe de recaudación',              // portada (se achica solo si es largo)
 *   subtitulo: 'Teletón 2026 · corte semanal',
 *   fecha: '18 de octubre de 2026',                 // si no viene, usa la fecha de hoy
 *   archivo: 'recaudacion-semana-42',               // opcional: nombre del archivo descargado
 *   secciones: [
 *     { tipo: 'seccion',   numero: '01', titulo: 'Recaudación', texto: 'Resumen al corte' },
 *     { tipo: 'progreso',  titulo: 'Avance hacia la meta', meta: 500000000, recaudado: 186450000,
 *                          cifras: [ { valor: 313550000, formato: 'colones', etiqueta: 'faltan' }, ... ] },   // cifras opcional
 *     { tipo: 'cifras',    titulo: 'Resumen', cifras: [ { valor: 1520, formato: 'numero', etiqueta: 'donantes' }, ... ], nota: '...' },  // hasta 4
 *     { tipo: 'grafico',   titulo: 'Recaudación por canal', grafico: 'columna' | 'barra' | 'linea' | 'dona',
 *                          etiquetas: ['Sinpe', 'Place to Pay'], series: [ { nombre: 'Colones', valores: [1, 2] } ],
 *                          formato: 'colones', comentario: 'Texto opcional a la derecha' },
 *     { tipo: 'tabla',     titulo: 'Detalle por ejecutiva', columnas: ['Ejecutiva', 'Monto'], filas: [['Ana', 1500000]],
 *                          formatos: [null, 'colones'], totales: ['Total', 1500000], filasPorPagina: 9, nota: '...' },  // pagina sola
 *     { tipo: 'texto',     titulo: 'Conclusiones', parrafo: '...', puntos: ['...', '...'] },
 *     { tipo: 'columnas',  titulo: 'Patrocinios', izquierda: { titulo: 'Firmados', puntos: [...] }, derecha: { titulo: 'Pendientes', texto: '...' } },
 *     { tipo: 'afiche',    etiqueta: 'Embajadora 2026', titulo: 'Lady Hidalgo Monge', texto: '...', dato: '¡Gracias!', foto: 'data:image/jpeg;base64,...' }
 *   ],
 *   cierre: { titulo: '¡Hagamos que suene!', texto: 'Gracias', contacto: 'teletoncr.com' }
 * }
 * Formatos de números: 'colones' (₡1 234 567), 'numero' (1 234), 'porcentaje' (45,3%), 'millones' (₡186,5 M).
 * ======================================================================= */
