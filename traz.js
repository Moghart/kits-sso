/* Planilla de trazabilidad del kit (.xlsx propio, sin librerías): mapa ↔ procedimiento ↔ MIPER ↔ RF/CC ↔ ECF. */
(function(){
'use strict';
const esc = s => String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'');
const colL = i => { let s = ''; i++; while (i > 0){ const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const NIV = {TOLERABLE:3, MODERADO:4, IMPORTANTE:5, INTOLERABLE:6};
function hoja(rows, widths, levelCols){
  const cols = '<cols>' + widths.map((w,i) => `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('') + '</cols>';
  const body = rows.map((r, ri) => `<row r="${ri+1}">` + r.map((v, ci) => {
    const ref = colL(ci) + (ri + 1);
    let st = ri === 0 ? 1 : 2; if (ri > 0 && levelCols.includes(ci) && NIV[v]) st = NIV[v];
    if (typeof v === 'number') return `<c r="${ref}" s="${st}"><v>${v}</v></c>`;
    return `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  }).join('') + '</row>').join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>${cols}<sheetData>${body}</sheetData><autoFilter ref="A1:${colL(rows[0].length-1)}${rows.length}"/></worksheet>`;
}
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts>
<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F8B8D"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF66FF66"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF9900"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FF999999"/></left><right style="thin"><color rgb="FF999999"/></right><top style="thin"><color rgb="FF999999"/></top><bottom style="thin"><color rgb="FF999999"/></bottom></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="7"><xf/>
<xf fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf>
<xf fontId="0" fillId="0" borderId="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf fontId="0" fillId="3" borderId="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf fontId="0" fillId="4" borderId="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf fontId="0" fillId="5" borderId="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf fontId="0" fillId="6" borderId="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
</cellXfs></styleSheet>`;
window.armarXlsx = function(hojas){
  const n = hojas.length;
  const files = [
    ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${hojas.map((h,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`],
    ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${hojas.map((h,i)=>`<sheet name="${esc(h.nombre)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets><definedNames>${hojas.map((h,i)=>`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${h.nombre}'!$A$1:$${colL(h.filas[0].length-1)}$${h.filas.length}</definedName>`).join('')}</definedNames></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hojas.map((h,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="rId${n+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', STYLES],
    ...hojas.map((h,i) => [`xl/worksheets/sheet${i+1}.xml`, hoja(h.filas, h.anchos, h.nivel || [])])
  ];
  return window.zipArchivos(files, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
};
})();
