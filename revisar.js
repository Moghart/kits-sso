/* Revisar documentos existentes: lectura de la MIPER (.xlsm/.xlsx), del procedimiento (.docx/.pdf) y del mapa,
   revisiones automáticas del formato y del criterio, e informe PDF de cambios. Todo corre en el navegador. */
(function(){
'use strict';
const colNum = s => s.split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0);
const textoNodo = n => [...n.getElementsByTagName('t')].map(t => t.textContent).join('');
const P = s => new DOMParser().parseFromString(s, 'application/xml');
const norm = s => String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9 ]/g,' ').replace(/\s+/g,' ').trim();
const sinNum = s => String(s||'').replace(/^\s*\d+(\.\d+)*\.?\s+/,'');
const clave = s => norm(sinNum(s)).slice(0,40);
const OCULTAS = ['RIESGOS','CRITERIOS','NIVELES','CONTROLES','CAUSAS-AMENAZA','Anexo N° 1','MIPER de EJEMPLO'];

/* ---------- MIPER SIGO-F-006 ---------- */
window.leerMiperDoc = async function(buf){
  const files = await window.leerZip(new Uint8Array(buf));
  const dec = new TextDecoder(); const get = n => { const f = files.find(x => x[0] === n); return f ? dec.decode(f[1]) : null; };
  const ss = get('xl/sharedStrings.xml'); const shared = ss ? [...P(ss).getElementsByTagName('si')].map(textoNodo) : [];
  const wbx = P(get('xl/workbook.xml')); const rels = P(get('xl/_rels/workbook.xml.rels'));
  const hojas = [...wbx.getElementsByTagName('sheet')].filter(s => !s.getAttribute('state') && !OCULTAS.includes(s.getAttribute('name')));
  if (!hojas.length) throw new Error('No encontré la hoja de la MIPER.');
  let mejor = null;
  for (const s of hojas){
    const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    let target = [...rels.getElementsByTagName('Relationship')].find(r => r.getAttribute('Id') === rid).getAttribute('Target');
    target = target.replace(/^\/?xl\//,'').replace(/^\//,'');
    const x = get('xl/' + target); if (!x) continue;
    if (!mejor || x.length > mejor.x.length) mejor = { x, nombre: s.getAttribute('name') };
  }
  const sheet = P(mejor.x); const G = {};
  for (const c of sheet.getElementsByTagName('c')){
    const m = /^([A-Z]+)(\d+)$/.exec(c.getAttribute('r')||''); if (!m) continue;
    const t = c.getAttribute('t'); const v = c.getElementsByTagName('v')[0];
    let val = t === 's' ? shared[Number(v && v.textContent)] : t === 'inlineStr' ? textoNodo(c) : (v ? v.textContent : '');
    if (!t && v && /^-?\d+\.\d{6,}$/.test(val)) val = String(Number(Number(val).toPrecision(12)));
    val = String(val == null ? '' : val).trim(); if (!val) continue;
    (G[m[2]] = G[m[2]] || {})[m[1]] = val;
  }
  const g = (r, c) => (G[r] || {})[c] || '';
  const filas = []; const trans = []; let tarea = '', cargo = '', dot = '', bloqueTrans = false, n = 0;
  const maxR = Math.max(...Object.keys(G).map(Number));
  for (let r = 10; r <= maxR; r++){
    if (/REALIZADO/i.test(g(r,'D'))) break;
    if (/transversal/i.test(g(r,'D')) || /^todas las tareas/i.test(g(r,'F'))) bloqueTrans = true;
    if (g(r,'F') && !/^todas las tareas/i.test(g(r,'F'))){ bloqueTrans = false; tarea = g(r,'F').split('\n')[0].replace(/\s+/g,' ').trim(); cargo = g(r,'G'); dot = g(r,'H'); n++; }
    if (!g(r,'K') && !g(r,'I')) continue;
    const d = {}; for (const col of 'DEFGHIJKLMNOPQRSTUV') d[col] = g(r, col);
    if (bloqueTrans){ trans.push(d); continue; }
    if (!g(r,'K')) continue;
    filas.push({ fila: r, tn: String(n), tarea, cargo, dot, peligro: d.I, cat: d.J, cod: d.K, P: d.L, C: d.M, vep: d.N, niv: d.O, med: d.P, jer: d.Q, Pr: d.R, Cr: d.S, vepr: d.T, nivr: d.U, resp: d.V });
  }
  return { hoja: mejor.nombre, hdr: { F: g(7,'T'), M: g(7,'U'), NB: g(7,'V') }, proceso: g(10,'D') && !/transversal/i.test(g(10,'D')) ? g(10,'D') : (Object.keys(G).map(Number).filter(r => r >= 10).map(r => g(r,'D')).find(x => x && !/transversal/i.test(x)) || ''),
    actividad: Object.keys(G).map(Number).filter(r => r >= 10).map(r => g(r,'E')).find(x => x && !/^todas/i.test(x)) || '', filas, trans };
};

/* ---------- Procedimiento (.docx o .pdf) ---------- */
let pdfjs = null;
async function cargarPdfJs(){
  if (pdfjs) return pdfjs;
  await new Promise((ok, ko) => { const s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js'; s.onload = ok; s.onerror = () => ko(new Error('No se pudo cargar el lector de PDF.')); document.head.appendChild(s); });
  pdfjs = window.pdfjsLib; pdfjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  return pdfjs;
}
window.leerProcDoc = async function(buf, nombre){
  let lineas = [];
  if (/\.pdf$/i.test(nombre)){
    const lib = await cargarPdfJs();
    const doc = await lib.getDocument({ data: new Uint8Array(buf) }).promise;
    for (let p = 1; p <= doc.numPages; p++){
      const tc = await (await doc.getPage(p)).getTextContent(); let y = null, l = '';
      for (const it of tc.items){ const yy = Math.round(it.transform[5]); if (y !== null && Math.abs(yy - y) > 3){ lineas.push(l); l = ''; } l += (l && !l.endsWith(' ') && !it.str.startsWith(' ') ? ' ' : '') + it.str; y = yy; }
      if (l) lineas.push(l);
    }
  } else {
    const files = await window.leerZip(new Uint8Array(buf));
    const f = files.find(x => x[0] === 'word/document.xml'); if (!f) throw new Error('El archivo no es un .docx válido.');
    const x = P(new TextDecoder().decode(f[1]));
    for (const p of x.getElementsByTagName('w:p')) lineas.push([...p.getElementsByTagName('w:t')].map(t => t.textContent).join(''));
  }
  lineas = lineas.map(l => l.replace(/\s+/g,' ').trim()).filter(Boolean);
  const texto = lineas.join('\n');
  const cuenta = {}; for (const m of texto.replace(/\s+/g,'').match(/VELTB-P-[A-Z0-9]+-[A-Z]{2,3}-\d{4}/g) || []){ const c = m.replace(/\s+/g,''); cuenta[c] = (cuenta[c]||0) + 1; }
  const cod = (Object.entries(cuenta).sort((a,b) => b[1]-a[1])[0] || [''])[0];
  let rev = (texto.match(/Revisi[oó]n\s*(?:N[°º]\s*)?:?\s*(\d{1,2})\b/i) || [])[1];
  const iH = lineas.findIndex(l => /historial del cambio/i.test(l));
  if (iH >= 0){ const fin = lineas.findIndex((l, i) => i > iH && /tabla de contenido|^1\.?\s*objetivo/i.test(l)); const ns = lineas.slice(iH, fin > 0 ? fin : iH + 80).filter(l => /^\d{1,2}$/.test(l)).map(Number); if (ns.length) rev = Math.max(...ns); }
  const rfs = [...new Set((texto.match(/\bRF\s?-?\s?(\d{2})\b/g) || []).map(s => 'RF' + s.replace(/\D/g,'')))].sort();
  return { texto, lineas, codigo: cod, revision: rev != null ? Number(rev) : null, rfs };
};

/* ---------- Revisiones automáticas ---------- */
const RC = cod => { const m = String(cod||'').match(/_R[CF](\d\d)(_|$)/); return m ? 'RF' + m[1] : null; };
window.revisionAutomatica = function(M, PR, MAPA, D){
  const H = []; const add = (prioridad, doc, ubic, problema, porque, como, fuente, extra) => H.push({ prioridad, doc, ubic, problema, porque, como, fuente, ...(extra||{}) });
  const CODES = D.catalogo.map(c => c.cod); const ENT = D.entorno.map(e => e.codigo);
  const tareas = [...new Map(M.filas.map(f => [f.tn, f])).values()];
  if (M.trans.length) add('FORMA','MIPER','bloque de riesgos transversales (' + M.trans.length + ' filas)', 'La MIPER trae filas transversales (psicosocial, emergencias, personas sensibles) al inicio.', 'La MIPER guía del área (kit MH-0001) parte directo con las tareas y no lleva ese bloque.', 'Quitar el bloque transversal; la MIPER corregida ya sale sin él. Si el área decide mantenerlo, rechaza este hallazgo.', 'MH-0001', { cambio: { tipo: 'transversales' } });
  const fueraLista = M.filas.filter(f => !CODES.includes(f.cod) && !CODES.some(c => c.split('_')[0] === String(f.cod).split('_')[0]));
  if (fueraLista.length) add('MAYOR','MIPER','filas ' + fueraLista.slice(0,6).map(f => f.fila).join(', ') + (fueraLista.length > 6 ? '…' : ''), fueraLista.length + ' filas con un código que no está en la lista del formato SIGO-F-006 (ej.: "' + fueraLista[0].cod + '").', 'El formato solo acepta los códigos de su lista desplegable; uno distinto rompe la trazabilidad con los RF.', 'Elegir el código de la lista que corresponde al peligro.', 'formato');
  const vepMal = M.filas.filter(f => { const p = Number(f.P), c = Number(f.C); return p && c && (Number(f.vep) !== p*c); });
  if (vepMal.length) add('MAYOR','MIPER','filas ' + vepMal.slice(0,6).map(f => f.fila).join(', '), vepMal.length + ' filas donde el valor esperado no es P × C.', 'El nivel de riesgo se calcula como P × C; un valor distinto cambia la prioridad del control.', 'La MIPER corregida recalcula P × C y el nivel en todas las filas.', 'formato', { aplicado: true });
  const rfAlto = M.filas.filter(f => RC(f.cod) && Number(f.Pr)*Number(f.Cr) >= 8);
  if (rfAlto.length) add('CRITICO','MIPER','filas ' + rfAlto.slice(0,6).map(f => f.fila).join(', '), rfAlto.length + ' filas de riesgo de fatalidad con residual IMPORTANTE o mayor (ej.: ' + RC(rfAlto[0].cod) + ' ' + rfAlto[0].Pr + '×' + rfAlto[0].Cr + ').', 'Un RF no puede quedar sobre MODERADO después de aplicar los controles críticos.', 'Agregar o reforzar controles críticos hasta bajar la probabilidad residual.', 'G-012', { filas: rfAlto.map(f => f.fila) });
  const sinCCP = M.filas.filter(f => RC(f.cod) && !/\bCC[PM]\s?\d/i.test(f.med));
  if (sinCCP.length) add('MAYOR','MIPER','filas ' + sinCCP.slice(0,6).map(f => f.fila).join(', ') + (sinCCP.length > 6 ? '…' : ''), sinCCP.length + ' filas de RF sin controles críticos (CCP/CCM) en las medidas.', 'La SIGO-G-012 exige declarar los controles críticos preventivos y mitigadores de cada RF expuesto.', 'Agregar los CCP y CCM de la SIGO-G-012 que aplican a la tarea.', 'G-012', { filas: sinCCP.map(f => f.fila) });
  const rfM = [...new Set(M.filas.map(f => RC(f.cod)).filter(Boolean))].sort();
  const noDecl = rfM.filter(rf => D.rfcat[rf] && !D.rfcat[rf].declarado);
  if (noDecl.length) add('CRITICO','MIPER','RF ' + noDecl.join(', '), 'La MIPER expone RF que BERLIAM no tiene declarados: ' + noDecl.join(', ') + '.', 'El RESSO v11 descuenta 5% por exponerse a un RF no declarado.', 'Confirmar si la tarea realmente expone ese RF. Si es así, declararlo; si no, cambiar el código.', 'formato');
  const faltaEnt = tareas.map(t => ({ t, falta: ENT.filter(c => !M.filas.some(f => f.tn === t.tn && f.cod.split('_')[0] === c.split('_')[0])) })).filter(x => x.falta.length);
  const cuenta = {}; for (const x of faltaEnt) for (const c of x.falta){ const e = D.entorno.find(e => e.codigo === c); const nom = String(c).split('_').slice(1).join(' ').replace(/^R[CF]\d\d /,''); cuenta[nom] = (cuenta[nom]||0) + 1; }
  if (faltaEnt.length) add('MAYOR','MIPER', faltaEnt.length + ' de ' + tareas.length + ' tareas', 'Faltan peligros de entorno: ' + Object.entries(cuenta).sort((a,b) => b[1]-a[1]).map(([n,c]) => n.toLowerCase() + ' (' + c + ' tareas)').join(', ') + '.', 'La guía MH-0001 los evalúa en cada tarea porque están presentes en todo el sector.', 'Agregar los peligros de entorno que faltan en cada tarea.', 'MH-0001', { cambio: { tipo: 'entorno', tareas: faltaEnt.map(x => x.t.tn) } });
  if (PR){
    const sinProc = rfM.filter(rf => !PR.rfs.includes(rf));
    if (sinProc.length) add('CRITICO','Procedimiento','tablas de RF', 'RF evaluados en la MIPER sin control en el procedimiento: ' + sinProc.join(', ') + '.', 'Cada RF expuesto en la MIPER debe tener su tabla de controles críticos en el procedimiento.', 'Agregar la tabla de controles del RF en el punto de peligros y controles del procedimiento.', 'G-012');
    const sobran = PR.rfs.filter(rf => !rfM.includes(rf) && D.rfcat[rf]);
    if (sobran.length) add('MAYOR','Procedimiento','tablas de RF', 'RF desarrollados en el procedimiento que la MIPER no evalúa: ' + sobran.join(', ') + '.', 'El procedimiento y la MIPER deben coincidir; si el RF existe, falta en la MIPER, y si no, sobra en el procedimiento.', 'Evaluar el RF en la tarea que corresponde o quitarlo del procedimiento.', 'formato');
    const tp = norm(PR.texto);
    const noEsta = tareas.filter(t => { const w = norm(sinNum(t.tarea)).split(' ').filter(x => x.length > 3).slice(0,4); return w.length && w.filter(x => tp.includes(x)).length < Math.min(3, w.length); });
    if (noEsta.length) add('MAYOR','Procedimiento','paso a paso', noEsta.length + ' tareas de la MIPER que no aparecen como paso del procedimiento: ' + noEsta.slice(0,4).map(t => '"' + t.tarea.slice(0,60) + '"').join('; ') + '.', 'Cada tarea evaluada debe estar en el paso a paso, en el mismo orden.', 'Agregar el paso con el mismo texto de la tarea.', 'formato');
  }
  if (MAPA){
    const tm = MAPA.actividades.flatMap(a => a.tareas);
    const falta = tm.filter(t => !tareas.some(x => clave(x.tarea) === clave(t.tarea)));
    if (falta.length) add('CRITICO','Mapa/MIPER', falta.length + ' tareas', 'Tareas del mapa sin evaluación en la MIPER: ' + falta.slice(0,4).map(t => '"' + t.tarea.slice(0,60) + '"').join('; ') + '.', 'Toda tarea del mapa SIGO-F-011 debe estar evaluada en la MIPER con el mismo texto.', 'Agregar la tarea en la MIPER y evaluar sus peligros.', 'formato');
    const dif = tareas.filter(x => { const t = tm.find(t => clave(t.tarea) === clave(x.tarea)); return t && String(t.dotacion) !== String(x.dot); });
    if (dif.length) add('MAYOR','Mapa/MIPER', dif.length + ' tareas', 'La dotación por tarea no coincide entre el mapa y la MIPER.', 'El mapa y la MIPER deben declarar las mismas personas expuestas.', 'Usar la dotación del mapa en la MIPER.', 'formato');
  }
  return H;
};

/* ---------- Informe PDF ---------- */
async function cargar(src){ await new Promise((ok, ko) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => ko(new Error('No se pudo cargar ' + src)); document.head.appendChild(s); }); }
const W = s => String(s == null ? '' : s).replace(/[→⇒]/g,'->').replace(/[≤]/g,'<=').replace(/[≥]/g,'>=').replace(/[“”]/g,'"').replace(/[‘’]/g,"'").replace(/[–—]/g,'-').replace(/…/g,'...').replace(/[^\x00-\xFF]/g,'');
window.informeRevisionPdf = async function(k, quien){
  if (!window.jspdf) await cargar('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  if (!window.jspdf.jsPDF.API.autoTable) await cargar('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js');
  const R = k.revisionExt; const doc = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const naranjo = [196, 74, 26];
  doc.setFont('helvetica','bold'); doc.setFontSize(15); doc.text(W('Informe de revisión de documentos'), 14, 16);
  doc.setFont('helvetica','normal'); doc.setFontSize(9.5);
  const H = R.hallazgos; const ac = H.filter(x => x.estado === 'aceptado'), re = H.filter(x => x.estado === 'rechazado'), pe = H.filter(x => !x.estado || x.estado === 'pendiente');
  const meta = [
    'BERLIAM · Codelco División Radomiro Tomic · Contrato 4600031257',
    'Kit: ' + (k.codigo || '-') + ' · ' + (k.titulo || '') ,
    'Revisión del documento: ' + (R.revOriginal != null ? R.revOriginal : '-') + ' -> ' + (k.revision),
    'Archivos revisados: ' + [R.archivos.miper, R.archivos.proc, R.archivos.mapa].filter(Boolean).join(' · '),
    'Fecha: ' + new Date().toLocaleDateString('es-CL') + (quien ? ' · Revisado por: ' + quien : ''),
    'Resultado: ' + H.length + ' hallazgos (' + H.filter(x=>x.prioridad==='CRITICO').length + ' críticos, ' + H.filter(x=>x.prioridad==='MAYOR').length + ' mayores, ' + H.filter(x=>x.prioridad==='FORMA').length + ' de forma) · ' + ac.length + ' aceptados, ' + re.length + ' rechazados, ' + pe.length + ' pendientes · ' + (R.cambios||[]).length + ' cambios aplicados a la MIPER.'
  ];
  meta.forEach((l, i) => doc.text(W(l), 14, 24 + i * 5));
  const fuenteTxt = x => x.fuente === 'MH-0001' ? 'Guía inicial MH-0001' : x.fuente === 'G-012' ? 'SIGO-G-012' : x.fuente === 'formato' ? 'Formato / RESSO' : /^regla:/.test(x.fuente||'') ? 'Regla del equipo' + (x.reglaAutor ? ' (' + x.reglaAutor + ')' : '') : 'Criterio de prevención';
  const dec = x => x.estado === 'aceptado' ? 'Aceptado' + (x.aplicado ? ' (aplicado)' : '') : x.estado === 'rechazado' ? 'Rechazado' + (x.motivo ? ': ' + x.motivo : '') : x.aplicado ? 'Aplicado' : 'Pendiente';
  doc.autoTable({ startY: 56, head: [['N°','Prioridad','Documento / ubicación','Qué está mal','Por qué','Cómo corregirlo','Origen','Decisión'].map(W)],
    body: H.map((x, i) => [i+1, x.prioridad, x.doc + (x.ubic ? ' · ' + x.ubic : ''), x.problema, x.porque, x.como, fuenteTxt(x), dec(x)].map(W)),
    styles: { fontSize: 7.5, cellPadding: 1.6, valign: 'top' }, headStyles: { fillColor: naranjo, fontSize: 8 },
    columnStyles: { 0: { cellWidth: 8 }, 1: { cellWidth: 18 }, 2: { cellWidth: 32 }, 3: { cellWidth: 55 }, 4: { cellWidth: 50 }, 5: { cellWidth: 55 }, 6: { cellWidth: 25 }, 7: { cellWidth: 26 } },
    didParseCell: d => { if (d.section === 'body' && d.column.index === 1){ d.cell.styles.fontStyle = 'bold'; d.cell.styles.textColor = d.cell.raw === 'CRITICO' ? [176,32,32] : d.cell.raw === 'MAYOR' ? [176,110,0] : [80,80,80]; } } });
  const C = R.cambios || [];
  if (C.length){
    doc.addPage(); doc.setFont('helvetica','bold'); doc.setFontSize(12); doc.text(W('Cambios aplicados a la MIPER'), 14, 16);
    doc.autoTable({ startY: 22, head: [['Tarea','Qué cambió','Antes','Después','Hallazgo'].map(W)],
      body: C.map(c => [c.tarea, c.que, c.antes, c.despues, c.h != null ? 'N° ' + (c.h + 1) : ''].map(W)),
      styles: { fontSize: 7.5, cellPadding: 1.6, valign: 'top' }, headStyles: { fillColor: naranjo, fontSize: 8 },
      columnStyles: { 0: { cellWidth: 40 }, 1: { cellWidth: 35 }, 2: { cellWidth: 85 }, 3: { cellWidth: 85 }, 4: { cellWidth: 20 } } });
  }
  const n = doc.getNumberOfPages();
  for (let i = 1; i <= n; i++){ doc.setPage(i); doc.setFontSize(7.5); doc.setTextColor(110); doc.text(W('Borrador para revisión de un prevencionista · ' + (k.codigo||'') + ' · Página ' + i + ' de ' + n), 14, 203); doc.setTextColor(0); }
  return new Uint8Array(doc.output('arraybuffer'));
};
})();
