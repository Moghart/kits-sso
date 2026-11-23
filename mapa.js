/* Lectura del mapa de procesos SIGO-F-011 (.xlsx) en el navegador: sección D (actividades y tareas operativas). */
(function(){
'use strict';
const colNum = s => s.split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0);
function textoNodo(n){ return [...n.getElementsByTagName('t')].map(t => t.textContent).join(''); }
window.leerMapa = async function(buf){
  const files = await window.leerZip(new Uint8Array(buf));
  const dec = new TextDecoder(); const get = n => { const f = files.find(x => x[0] === n); return f ? dec.decode(f[1]) : null; };
  const P = s => new DOMParser().parseFromString(s, 'application/xml');
  const ss = get('xl/sharedStrings.xml'); const shared = ss ? [...P(ss).getElementsByTagName('si')].map(textoNodo) : [];
  const wbx = P(get('xl/workbook.xml')); const rels = P(get('xl/_rels/workbook.xml.rels'));
  const first = wbx.getElementsByTagName('sheet')[0]; const rid = first.getAttribute('r:id') || first.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
  let target = [...rels.getElementsByTagName('Relationship')].find(r => r.getAttribute('Id') === rid).getAttribute('Target');
  target = target.replace(/^\/?xl\//,'').replace(/^\//,''); const sheet = P(get('xl/' + target));
  const G = {};  // G[row][col] = texto
  for (const c of sheet.getElementsByTagName('c')){
    const ref = c.getAttribute('r'); const m = /^([A-Z]+)(\d+)$/.exec(ref); if (!m) continue;
    const t = c.getAttribute('t'); const v = c.getElementsByTagName('v')[0];
    let val = t === 's' ? shared[Number(v && v.textContent)] : t === 'inlineStr' ? textoNodo(c) : (v ? v.textContent : '');
    if (!t && v && /^-?\d+\.\d{6,}$/.test(val)) val = String(Number(Number(val).toPrecision(12)));
    val = String(val == null ? '' : val).trim(); if (!val) continue;
    (G[m[2]] = G[m[2]] || {})[colNum(m[1])] = val;
  }
  const R = Object.keys(G).map(Number).sort((a,b) => a - b);
  const cel = (r, c) => (G[r] || {})[c] || '';
  const meta = {};
  for (const r of R){ const row = G[r]; for (const c in row){ const t = row[c].toLowerCase();
    if (t === 'gerencia') meta.gerencia = cel(r+1, Number(c)); if (t.startsWith('superintendencia')) meta.superintendencia = cel(r+1, Number(c));
    if (t === 'área' || t === 'area') meta.area = cel(r+1, Number(c)); if (t.startsWith('ee.cc')) meta.contrato = cel(r+1, Number(c)); } }
  const hr = R.find(r => Object.values(G[r]).some(v => /^tareas de la actividad/i.test(v)));
  if (!hr) throw new Error('No encontré la tabla "Tareas de la Actividad" del SIGO-F-011.');
  const H = {}; for (const c in G[hr]){ const t = G[hr][c].toLowerCase();
    if (/^n°\s*a/.test(t)) H.na = +c; else if (/^actividades/.test(t)) H.act = +c; else if (/^n°\s*t/.test(t)) H.nt = +c; else if (/^tareas/.test(t)) H.tarea = +c;
    else if (/r\s*\/\s*nr/.test(t)) H.rnr = +c; else if (/^cargo/.test(t)) H.cargo = +c; else if (/^lugar/.test(t)) H.lugar = +c; else if (/^n°\s*de personas/.test(t)) H.dot = +c; }
  const sub = G[hr+1] || {}; for (const c in sub){ const t = sub[c].toUpperCase(); if (t === 'M') H.m = +c; if (t === 'F') H.f = +c; if (t === 'NB') H.nb = +c; if (/por tarea/i.test(sub[c])) H.dot = +c; }
  // sección C: lista de procesos operativos (N° | Procesos Operativos) para saber si la actividad es mecánica, eléctrica, etc.
  meta.procesos = {};
  const hc = R.find(r => Object.values(G[r]).some(v => /^procesos operativos$/i.test(v)));
  if (hc) for (const r of R.filter(r => r > hc && r < hr)){ const n = cel(r, 1) || cel(r, 2); const nom = Object.entries(G[r]).filter(([c]) => +c > 1).map(([,v]) => v).find(v => /[a-záéíóú]/i.test(v) && v !== n); if (/^\d+$/.test(n) && nom) meta.procesos[n] = nom; else if (/^[A-Z]\.$/.test(cel(r,1))) break; }
  const acts = []; let cur = null;
  for (const r of R.filter(r => r > hr + 1)){
    const a = cel(r, 1); if (/^[A-Z]\.$/.test(a)) break;            // siguiente sección (E. procesos administrativos)
    const nombre = H.act ? cel(r, H.act) : ''; const tarea = H.tarea ? cel(r, H.tarea) : '';
    if (nombre){ cur = { n: (H.na && cel(r, H.na)) || String(acts.length + 1), nombre, tareas: [] }; acts.push(cur); }
    if (!tarea) continue;
    if (!cur){ cur = { n: '1', nombre: 'Actividad', tareas: [] }; acts.push(cur); }
    const cargos = (cel(r, H.cargo) || '').split(/\n|;/).map(x => x.replace(/^\s*\d+\s*[.)-]?\s*/, '').trim()).filter(Boolean);
    const rnr = /n\s*\/?\s*r/i.test(cel(r, H.rnr)) ? 'NR' : 'R';
    cur.tareas.push({ cod: cel(r, H.nt), tarea: tarea.replace(/\s+/g,' ').trim(), rnr, cargos, lugar: cel(r, H.lugar), dotacion: Number(cel(r, H.dot)) || '', m: cel(r, H.m), f: cel(r, H.f), nb: cel(r, H.nb) });
  }
  return { meta, actividades: acts.filter(a => a.tareas.length) };
};
})();
