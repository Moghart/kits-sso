/* MIPER oficial SIGO-F-006 v7 DRT (.xlsm con macros) armada sobre la plantilla del formato,
   editando el XML de la hoja 8 (mismo método que build_miper.py del proyecto). */
(function(){
'use strict';
async function cargarPlantilla(){
  if (!window.PLANTILLA_MIPER){
    await new Promise((ok, ko) => { const s = document.createElement('script'); s.src = 'plantilla_miper.js?v=803afbd9'; s.onload = ok; s.onerror = () => ko(new Error('plantilla')); document.head.appendChild(s); });
  }
  const bin = atob(window.PLANTILLA_MIPER); const u = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return u;
}
async function inflar(data){
  const ds = new DecompressionStream('deflate-raw');
  return new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(ds)).arrayBuffer());
}
async function unzip(u){
  const dv = new DataView(u.buffer, u.byteOffset, u.byteLength);
  let e = u.length - 22; while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  const n = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true);
  const dec = new TextDecoder(); const out = [];
  for (let i = 0; i < n; i++){
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true), nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
    const name = dec.decode(u.subarray(p + 46, p + 46 + nlen));
    const lnl = dv.getUint16(off + 26, true), lxl = dv.getUint16(off + 28, true);
    const raw = u.subarray(off + 30 + lnl + lxl, off + 30 + lnl + lxl + csize);
    out.push([name, method === 8 ? await inflar(raw) : raw.slice()]);
    p += 46 + nlen + xlen + clen;
  }
  return out;
}
const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const istr = (ref, st, t) => (t == null || t === '') ? `<c r="${ref}" s="${st}"/>` : `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${esc(t)}</t></is></c>`;
const num = (ref, st, v) => `<c r="${ref}" s="${st}"><v>${v}</v></c>`;
const val = (ref, st, v) => typeof v === 'number' ? num(ref, st, v) : istr(ref, st, v);
const ST = {D:247,E:247,F:247,G:247,H:248,I:204,J:184,K:151,L:150,M:150,N:150,O:201,P:80,Q:222,R:203,S:203,T:203,U:201,V:200};
const COLS = 'DEFGHIJKLMNOPQRSTUV'.split('');
const lines = (t, w) => String(t == null ? '' : t).split('\n').reduce((a, l) => a + Math.max(1, Math.ceil(l.length / w)), 0);

const JER = [[/elimin/i,'ELIMINACIÓN'],[/sustitu/i,'SUSTITUCIÓN'],[/ingenier/i,'CONTROLES DE INGENIERÍA'],[/administ/i,'CONTROLES ADMINISTRATIVOS'],[/epp|protecci[oó]n personal/i,'EQUIPOS DE PROTECCIÓN PERSONAL']];
const jerOficial = t => { const x = JER.find(([re]) => re.test(String(t||''))); return x ? x[1] : String(t||''); };
/* Texto de medidas parejo: sin líneas en blanco de más ni espacios al inicio, y sin secciones que no aplican
   (en P5 solo queda la parte de UV de origen solar, que es la que se ve en la MIPER guía; láser e infrarroja no aplican). */
function limpiarMedidas(t, cod){
  let x = String(t == null ? '' : t).replace(/\r/g, '');
  if (/^\(P5\)/.test(cod || '')){ const i = x.search(/\n\s*Radiaci[oó]n UV\s*\n/); if (i > 0) x = x.slice(0, i); }
  return x.split('\n').map(l => l.replace(/\s+$/,'').replace(/^\s+(?=\S)/, m => m.length > 1 ? '' : m)).join('\n').replace(/\n{2,}/g, '\n').trim();
}
function celdasFila(f){
  const d = { I: f.peligro, J: f.cat, K: f.codigo, P: limpiarMedidas(f.medidas, f.codigo), Q: jerOficial(f.jerarquia), V: f.resp || 'ADC / SUPERVISIÓN' };
  if (f.P && f.C){ Object.assign(d, { L: f.P, M: f.C, N: f.vep, O: f.niv }); }
  else Object.assign(d, { L: 'NA', M: 'NA', N: f.vep || '', O: f.niv || '' });
  if (f.Pr && f.Cr){ Object.assign(d, { R: f.Pr, S: f.Cr, T: f.vepr, U: f.nivr }); }
  else Object.assign(d, { R: 'NA', S: 'NA', T: f.vepr || '', U: f.nivr || '' });
  return d;
}

window.leerZip = unzip;
window.armarMiperOficial = async function(k, meta, transversales){
  const files = await unzip(await cargarPlantilla());
  const get = n => new TextDecoder().decode(files.find(f => f[0] === n)[1]);
  const put = (n, s) => { files.find(f => f[0] === n)[1] = new TextEncoder().encode(s); };
  // estilo propio para la columna P (medidas): mismo formato que el original pero alineado arriba e izquierda, para que todas las filas se vean iguales
  let stP = ST.P;
  { const sx = get('xl/styles.xml'); const m = /<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/.exec(sx);
    if (m){ const n = Number(m[1]); const xf = '<xf numFmtId="0" fontId="26" fillId="0" borderId="37" xfId="2" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>';
      put('xl/styles.xml', sx.replace(m[0], `<cellXfs count="${n+1}">${m[2]}${xf}</cellXfs>`)); stP = n; } }
  const s8 = get('xl/worksheets/sheet8.xml');
  const i0 = s8.indexOf('<sheetData>'), i1 = s8.indexOf('</sheetData>');
  const rows = {};
  for (const m of s8.slice(i0, i1).matchAll(/<row [^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)) rows[Number(/<row r="(\d+)"/.exec(m[0])[1])] = m[0];
  const cells = rx => [...rx.matchAll(/<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)].map(m => [m[1], m[2], m[3]]);

  const data = []; const merges = []; let r = 10;
  const t0 = r;
  transversales.forEach((h, i) => { const d = { ...h }; if (i === 0) Object.assign(d, { D: 'Riesgos transversales', E: 'Todas las actividades', F: 'Todas las tareas', G: 'Todos los puestos de trabajo', H: meta.nper }); data.push([r, d]); r++; });
  if (r > t0) merges.push(`D${t0}:D${r-1}`, `E${t0}:E${r-1}`, `F${t0}:F${r-1}`, `G${t0}:G${r-1}`, `H${t0}:H${r-1}`);
  const a0 = r;
  for (const t of k.listado.tareas){
    const fs = k.miper.filter(f => String(f.tarea_n) === String(t.n)); if (!fs.length) continue;
    const f0 = r;
    fs.forEach((f, i) => { const d = celdasFila(f); if (i === 0) Object.assign(d, { F: t.n + '. ' + t.tarea, G: (t.cargos||[]).join(' / '), H: Number(t.dotacion) || '' }); if (r === a0) Object.assign(d, { D: meta.proceso, E: meta.actividad }); data.push([r, d]); r++; });
    if (r - 1 > f0) merges.push(`F${f0}:F${r-1}`, `G${f0}:G${r-1}`, `H${f0}:H${r-1}`);
  }
  merges.push(`D${a0}:D${r-1}`, `E${a0}:E${r-1}`);
  const DATA_END = r - 1;
  while (r <= 89){ data.push([r, {}]); r++; }   // conserva la lista oculta AH/AI del formato (filas 10-89)
  const LAST = r - 1, SHIFT = LAST - 89;

  const rowXml = (rn, d) => {
    let ht = Object.keys(d).length ? Math.min(Math.max(lines(d.P, 78), lines(d.I, 36), lines(d.K, 22), lines(d.F || '', 30), 4) * 15 + 12, 409) : 15;
    const cs = COLS.map(c => val(c + rn, c === 'P' && Object.keys(d).length ? stP : ST[c], d[c] == null ? '' : d[c]));
    if (rows[rn] && rn <= 89) for (const [col, attrs, inner] of cells(rows[rn])) if (col === 'AH' || col === 'AI') cs.push(inner ? `<c r="${col}${rn}"${attrs}>${inner}</c>` : `<c r="${col}${rn}"${attrs}/>`);
    return `<row r="${rn}" spans="3:35" ht="${ht.toFixed(1)}" customHeight="1">` + cs.join('') + '</row>';
  };
  const shiftRow = (rx, dn) => { const n = Number(/<row r="(\d+)"/.exec(rx)[1]); return rx.replace(/<row r="\d+"/, `<row r="${n + dn}"`).replace(/(<c r="[A-Z]+)(\d+)"/g, (m, a, b) => `${a}${Number(b) + dn}"`); };
  const setCell = (rx, ref, text) => {
    const re = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`); const m = re.exec(rx); if (!m) return rx;
    const sty = /s="(\d+)"/.exec(m[1])[1];
    return rx.slice(0, m.index) + (typeof text === 'number' ? num(ref, sty, text) : istr(ref, sty, text)) + rx.slice(m.index + m[0].length);
  };
  const H = {}; for (const n of Object.keys(rows).map(Number).sort((a,b)=>a-b)) if (n <= 9) H[n] = rows[n];
  H[6] = setCell(H[6], 'I6', 'GERENCIA: MANTENIMIENTO');
  H[6] = setCell(H[6], 'O6', 'SUPERINTENDENCIA: MANTENIMIENTO / ÁREA: ' + (meta.area || 'CHANCADO PRIMARIO Y CORREAS').toUpperCase() + ' / EE.CC. BERLIAM - CONTRATO N° 4600031257');
  ['T','U','V'].forEach((c, i) => { const v = meta.counts[i]; H[7] = setCell(H[7], c + '7', v === '' || v == null ? '' : Number(v)); });
  let body = Object.keys(H).map(Number).sort((a,b)=>a-b).map(n => H[n]).join('');
  body += data.map(([rn, d]) => rowXml(rn, d)).join('');
  const FOOT = {}; for (let n = 90; n <= 106; n++) if (rows[n]) FOOT[n] = shiftRow(rows[n], SHIFT);
  const fset = (n, col, text) => { FOOT[n] = setCell(FOOT[n], col + (n + SHIFT), text); };
  ['D','G','J'].forEach((col, i) => { const [nm, cg] = meta.firmas[i] || ['','']; fset(101, col, nm || ''); fset(102, col, cg || ''); fset(103, col, ''); });
  body += Object.keys(FOOT).map(Number).sort((a,b)=>a-b).map(n => FOOT[n]).join('');
  const FOOT_END = 105 + SHIFT;

  let sh = s8.slice(0, i0) + '<sheetData>' + body + s8.slice(i1);
  sh = sh.replace('<dimension ref="A1:AK1046"/>', `<dimension ref="A1:AK${FOOT_END + 1}"/>`);
  sh = sh.replace('tabSelected="1" topLeftCell="A12" zoomScale="40" zoomScaleNormal="40"', 'tabSelected="1" topLeftCell="A1" zoomScale="50" zoomScaleNormal="50"');
  sh = sh.replace(/<selection [^>]*\/>/g, '<selection activeCell="I10" sqref="I10"/>');
  const keep = [];
  for (const m of sh.matchAll(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g)){
    const r1 = Number(m[2]), r2 = Number(m[4]);
    if (r2 <= 9) keep.push(`${m[1]}${r1}:${m[3]}${r2}`); else if (r1 >= 90) keep.push(`${m[1]}${r1 + SHIFT}:${m[3]}${r2 + SHIFT}`);
  }
  const allm = keep.concat(merges);
  sh = sh.replace(/<mergeCells count="\d+">[\s\S]*?<\/mergeCells>/, `<mergeCells count="${allm.length}">` + allm.map(m => `<mergeCell ref="${m}"/>`).join('') + '</mergeCells>');
  sh = sh.replace('sqref="K95:K100 K106:K1048576 K1:K93"', `sqref="K1:K${LAST}"`);
  for (const c of 'QLNMTRS') sh = sh.split(`sqref="${c}10:${c}89"`).join(`sqref="${c}10:${c}${LAST}"`);
  sh = sh.replace('<xm:sqref>J10:J89</xm:sqref>', `<xm:sqref>J10:J${LAST}</xm:sqref>`).replace('<xm:sqref>P10:P89</xm:sqref>', `<xm:sqref>P10:P${LAST}</xm:sqref>`);
  sh = sh.replace('sqref="T10:U89 N10:O89"', `sqref="T10:U${LAST} N10:O${LAST}"`).split('sqref="U10:U89 O10:O89"').join(`sqref="U10:U${LAST} O10:O${LAST}"`);
  sh = sh.replace(/<colBreaks [\s\S]*?<\/colBreaks>/, '');
  sh = sh.replace('<pageSetup paperSize="9" fitToWidth="0"', '<pageSetup paperSize="8" fitToWidth="1" fitToHeight="0"');
  put('xl/worksheets/sheet8.xml', sh);

  const hoja = meta.hoja.replace(/[\[\]:*?\/\\']/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 31) || 'MIPER';
  let wb = get('xl/workbook.xml').split('Mtto y cambio motor Media Tensi').join(hoja.replace(/&/g,'&amp;'));
  wb = wb.replace(/<definedName name="_xlnm.Print_Area" localSheetId="7">[^<]*<\/definedName>/,
    `<definedName name="_xlnm.Print_Area" localSheetId="7">'${hoja}'!$C$2:$V$${FOOT_END}</definedName><definedName name="_xlnm.Print_Titles" localSheetId="7">'${hoja}'!$8:$9</definedName>`);
  put('xl/workbook.xml', wb);
  if (files.find(f => f[0] === 'xl/calcChain.xml')) put('xl/calcChain.xml', get('xl/calcChain.xml').replace(/<c [^>]*i="8"[^>]*\/>/g, ''));
  return window.zipArchivos(files, 'application/vnd.ms-excel.sheet.macroEnabled.12');
};
})();
