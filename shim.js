/* Kits SSO · versión independiente (Apps Script + Google Sheets).
   Reemplaza window.claude (db, sample, downloads) por el servidor de Apps Script
   y agrega el inicio de sesión y la administración de usuarios. */
(function(){
'use strict';
const API = window.KITSSO_API;
let token = null, yo = null, listo; const prom = new Promise(r => listo = r);
try { token = localStorage.getItem('kitsso-token'); } catch (_) {}

function call(fn, ...args){
  return fetch(API, { method: 'POST', body: JSON.stringify({ f: fn, a: args }), credentials: 'omit', redirect: 'follow' })
    .catch(() => { throw { code: 'red', message: 'Sin conexión con el servidor. Revisa tu internet e intenta de nuevo.' }; })
    .then(r => r.json())
    .then(o => {
      if (o && o.ok) return o.r;
      const m = (o && o.e) || 'Error del servidor';
      if (m === 'SESION'){ salir(true); throw { code: 'not_granted', message: 'Tu sesión expiró.' }; }
      if (m === 'SIN_CLAVE') throw { code: 'sin_clave', message: m };
      if (m === 'LIMITE') throw { code: 'rate_limited', message: m };
      if (m === 'PROMPT_GRANDE') throw { code: 'prompt_too_large', message: m };
      throw { code: /^API_/.test(m) ? m.toLowerCase() : 'error', message: m };
    });
}
const T = (fn, ...a) => call(fn, token, ...a);

/* ---------- base de datos con la misma forma que la app espera ---------- */
const oyentes = {};
const snap = rows => ({ docs: rows.map(r => ({ id: r.id, exists: true, data: () => r.data })) });
async function avisar(col){ const l = oyentes[col]; if (!l || !l.length) return; try { const s = snap(await T('dbList', col)); for (const [cb] of l) cb(s); } catch (e) { for (const [, er] of l) er && er(e); } }
const pendientes = new Set(); let tmr = null;
function cambio(col){ pendientes.add(col); clearTimeout(tmr); tmr = setTimeout(() => { for (const c of pendientes) avisar(c); pendientes.clear(); }, 600); }
const db = {
  collection(col){ return { onSnapshot(cb, er){ (oyentes[col] = oyentes[col] || []).push([cb, er]); avisar(col); return () => { oyentes[col] = oyentes[col].filter(x => x[0] !== cb); }; } }; },
  doc(path){ const col = path.slice(0, path.indexOf('/')); return {
    async get(){ const d = await T('dbGet', path); return { exists: !!d.exists, data: () => d.data }; },
    async set(obj){ await T('dbSet', path, obj); cambio(col); },
    async delete(){ await T('dbDel', path); cambio(col); }
  }; }
};
setInterval(() => { if (token && !document.hidden) for (const c of Object.keys(oyentes)) avisar(c); }, 60000);

/* ---------- IA por el servidor (la clave nunca llega al navegador) ---------- */
async function pedirIA(input, opts){
  opts = opts || {};
  const prompt = typeof input === 'string' ? input : input.map(t => t.content).join('\n\n');
  const t0 = Date.now();
  const reloj = setInterval(() => { const el = document.getElementById('busyN'); if (el) el.textContent = 'Procesando… ' + Math.round((Date.now() - t0) / 1000) + ' s (puede tardar hasta 3 minutos)'; }, 1000);
  const cancel = new Promise((_, rej) => opts.signal && opts.signal.addEventListener('abort', () => rej({ code: 'cancelled' })));
  try {
    const r = await Promise.race([T('ia', prompt, opts.modelTier || 'default'), cancel]);
    if (opts.onText) opts.onText({ text: r.text, delta: r.text });
    return { text: r.text, truncated: r.stop === 'max_tokens' };
  } finally { clearInterval(reloj); }
}
const sample = async (input, opts) => await pedirIA(input, opts);
sample.json = async (input, opts) => {
  const { text } = await pedirIA(input, opts);
  const t = text.replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '');
  try { return JSON.parse(t.slice(t.indexOf('{'), t.lastIndexOf('}') + 1)); } catch { throw { code: 'invalid_json', text }; }
};
sample.limits = async () => ({ images: false });

/* ---------- descargas directas del navegador ---------- */
const downloads = { async save({ filename, data }){
  const blob = data instanceof Blob ? data : new Blob([data]);
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  return { status: 'saved' };
} };

window.claude = { use: async (n) => { await prom; return n === 'db' ? db : n === 'sample' ? sample : n === 'downloads' ? downloads : null; } };

/* ---------- pantalla de ingreso ---------- */
const css = document.createElement('style');
css.textContent = `
#kxLogin{position:fixed;inset:0;z-index:9999;background:var(--bg,#f6f5f2);display:flex;align-items:center;justify-content:center;padding:16px;font:15px/1.5 Inter,system-ui,sans-serif}
#kxLogin .bx{width:100%;max-width:380px;background:#fff;border:1px solid #e3e0d8;border-radius:14px;padding:28px;box-shadow:0 8px 30px rgba(0,0,0,.06);color:#1d1d1b}
#kxLogin h1{font-size:22px;margin:0 0 4px} #kxLogin p{margin:0 0 16px;color:#5f5d57;font-size:14px}
#kxLogin label{display:block;font-size:13px;font-weight:600;margin:12px 0 4px}
#kxLogin input{width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cfcbc1;border-radius:8px;font:inherit}
#kxLogin,#kxAdm{color-scheme:light} #kxLogin input,#kxAdm input{background:#fff;color:#1d1d1b} #kxLogin input[readonly]{background:#f3f2ee;color:#5f5d57}
#kxLogin button{width:100%;margin-top:18px;padding:11px;border:0;border-radius:8px;background:#c2410c;color:#fff;font:600 15px Inter,system-ui,sans-serif;cursor:pointer}
#kxLogin button:disabled{opacity:.6} #kxLogin .er{color:#b91c1c;font-size:13px;margin-top:10px;min-height:1em}
#kxLogin a{color:#c2410c;font-size:13px;cursor:pointer;display:inline-block;margin-top:14px}
#kxBar{position:fixed;right:12px;bottom:12px;z-index:9000;display:flex;gap:8px;font:13px Inter,system-ui,sans-serif}
#kxBar button{padding:7px 12px;border-radius:8px;border:1px solid #d6d2c8;background:#fff;color:#1d1d1b;cursor:pointer}
#kxAdm{position:fixed;inset:0;z-index:9500;background:rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center;padding:16px;font:14px/1.5 Inter,system-ui,sans-serif}
#kxAdm .bx{width:100%;max-width:560px;max-height:90vh;overflow:auto;background:#fff;border-radius:14px;padding:24px;color:#1d1d1b}
#kxAdm table{width:100%;border-collapse:collapse;margin:10px 0} #kxAdm td,#kxAdm th{border-bottom:1px solid #eee;padding:6px 4px;text-align:left;font-size:13px}
#kxAdm input{padding:8px;border:1px solid #cfcbc1;border-radius:8px;font:inherit;width:100%;box-sizing:border-box;margin-top:4px}
#kxAdm .row{display:flex;gap:8px;flex-wrap:wrap;align-items:end} #kxAdm .row>*{flex:1 1 160px}
#kxAdm button{padding:8px 12px;border-radius:8px;border:1px solid #d6d2c8;background:#fff;cursor:pointer} #kxAdm button.pri{background:#c2410c;color:#fff;border-color:#c2410c}
`;
document.head.appendChild(css);
const el = (t, a, ...k) => { const e = document.createElement(t); for (const [x, v] of Object.entries(a || {})){ if (x === 'text') e.textContent = v; else if (x.startsWith('on')) e.addEventListener(x.slice(2), v); else e.setAttribute(x, v); } e.append(...k); return e; };
function salir(expiro){ if (!expiro && token) call('logout', token).catch(() => {}); token = null; try { localStorage.removeItem('kitsso-token'); } catch (_) {} location.reload(); }
function pantallaLogin(primerAdmin, aviso){
  document.getElementById('kxLogin')?.remove();
  const em = el('input', { type: 'email', autocomplete: 'username', id: 'kxEm' });
  const pw = el('input', { type: 'password', autocomplete: primerAdmin ? 'new-password' : 'current-password', id: 'kxPw' });
  const er = el('div', { class: 'er', role: 'alert', text: aviso || '' });
  const bt = el('button', { type: 'submit', text: primerAdmin ? 'Crear administrador y entrar' : 'Entrar' });
  const f = el('form', {}, el('label', { for: 'kxEm', text: 'Correo' }), em, el('label', { for: 'kxPw', text: primerAdmin ? 'Contraseña (mínimo 8 caracteres)' : 'Contraseña' }), pw, bt, er);
  f.addEventListener('submit', async (ev) => {
    ev.preventDefault(); er.textContent = ''; bt.disabled = true; bt.textContent = 'Entrando…';
    try {
      const r = await call(primerAdmin ? 'primerAdmin' : 'login', em.value, pw.value);
      token = r.token; try { localStorage.setItem('kitsso-token', token); } catch (_) {}
      yo = r; entrar();
    } catch (e) { er.textContent = e.message || 'No se pudo entrar.'; bt.disabled = false; bt.textContent = primerAdmin ? 'Crear administrador y entrar' : 'Entrar'; }
  });
  document.body.appendChild(el('div', { id: 'kxLogin' }, el('div', { class: 'bx' },
    el('h1', { text: 'Kits SSO' }),
    el('p', { text: primerAdmin ? 'Primera vez: crea la cuenta del administrador. Después, el administrador agrega a los demás usuarios.' : 'BERLIAM · Codelco DRT. Entra con el correo y la contraseña que te dio el administrador.' }), f,
    primerAdmin ? '' : el('p', { style: 'margin:14px 0 0;text-align:center' }, el('a', { href: '#', id: 'kxOlv', text: '¿Olvidaste tu contraseña?', onclick: (ev) => { ev.preventDefault(); pantallaOlvide(em.value); } })))));
  em.focus();
}
function pantallaOlvide(correo){
  document.getElementById('kxLogin')?.remove();
  const em = el('input', { type: 'email', autocomplete: 'username', id: 'kxEm' }); em.value = correo || '';
  const co = el('input', { type: 'text', inputmode: 'numeric', autocomplete: 'one-time-code', id: 'kxCo', maxlength: '6' });
  const pw = el('input', { type: 'password', autocomplete: 'new-password', id: 'kxPw' });
  const er = el('div', { class: 'er', role: 'alert' });
  const paso2 = el('div', { hidden: '' }, el('label', { for: 'kxCo', text: 'Código que llegó a tu correo' }), co, el('label', { for: 'kxPw', text: 'Contraseña nueva (mínimo 8 caracteres)' }), pw);
  const bt = el('button', { type: 'submit', text: 'Enviarme un código' });
  const txt = el('p', { text: 'Escribe tu correo. Te enviaremos un código de 6 números para crear una contraseña nueva.' });
  const f = el('form', {}, el('label', { for: 'kxEm', text: 'Correo' }), em, paso2, bt, er);
  let enviado = false;
  f.addEventListener('submit', async (ev) => {
    ev.preventDefault(); er.textContent = ''; bt.disabled = true;
    try {
      if (!enviado){ bt.textContent = 'Enviando…'; await call('olvide', em.value);
        enviado = true; paso2.hidden = false; em.readOnly = true; txt.textContent = 'Si el correo tiene cuenta, te llegó un código (revisa también Spam). Vence en 15 minutos.'; bt.textContent = 'Guardar y entrar'; co.focus(); }
      else { bt.textContent = 'Guardando…'; const r = await call('restablecer', em.value, co.value, pw.value);
        token = r.token; try { localStorage.setItem('kitsso-token', token); } catch (_) {} yo = r; entrar(); }
    } catch (e) { er.textContent = e.message || 'No se pudo.'; bt.textContent = enviado ? 'Guardar y entrar' : 'Enviarme un código'; }
    bt.disabled = false;
  });
  document.body.appendChild(el('div', { id: 'kxLogin' }, el('div', { class: 'bx' }, el('h1', { text: 'Contraseña nueva' }), txt, f,
    el('p', { style: 'margin:14px 0 0;text-align:center' }, el('a', { href: '#', text: 'Volver a ingresar', onclick: (ev) => { ev.preventDefault(); pantallaLogin(false); } })))));
  (em.value ? (bt.focus(), null) : em.focus());
}
function entrar(){ document.getElementById('kxLogin')?.remove(); document.getElementById('kxHide')?.remove(); window.KITSSO_YO = { email: yo.email, admin: !!yo.admin }; barra(); listo(); }
function barra(){
  document.getElementById('kxBar')?.remove();
  const b = el('div', { id: 'kxBar' });
  if (yo.admin) b.append(el('button', { text: 'Usuarios', onclick: panelUsuarios }));
  b.append(el('button', { text: 'Cambiar contraseña', onclick: panelClave }));
  b.append(el('button', { text: 'Salir (' + yo.email + ')', onclick: () => salir(false) }));
  document.body.appendChild(b);
}
async function panelUsuarios(recien){
  document.getElementById('kxAdm')?.remove();
  const cont = el('div', { class: 'bx' }, el('h2', { text: 'Usuarios' }), el('p', { text: 'Cargando…' }));
  const ov = el('div', { id: 'kxAdm', onclick: (e) => { if (e.target === ov) ov.remove(); } }, cont); document.body.appendChild(ov);
  let j; try { j = await T('usuarios'); } catch (e) { cont.lastChild.textContent = e.message; return; }
  const tb = el('tbody');
  for (const u of j.usuarios.filter(x => x.activo)){ const x = j.uso[u.email] || { n: 0, tin: 0, tout: 0 };
    tb.append(el('tr', {}, el('td', { text: u.email }), el('td', { text: u.admin ? 'Admin' : 'Usuario' }), el('td', { text: x.n + ' análisis · ' + Math.round((x.tin + x.tout) / 1000) + 'k tokens' }),
      el('td', { style: 'white-space:nowrap' }, el('button', { text: 'Nueva contraseña', onclick: () => resetear(u) }), ' ', u.email === yo.email ? '' : el('button', { text: 'Quitar', onclick: async () => { try { await T('quitarUsuario', u.email); panelUsuarios(); } catch (e) { alertar(e.message); } } })))); }
  const em = el('input', { type: 'email', placeholder: 'correo@empresa.cl', 'aria-label': 'Correo' });
  const pw = el('input', { type: 'text', placeholder: 'Contraseña inicial (8+)', 'aria-label': 'Contraseña inicial' });
  const tel = el('input', { type: 'tel', placeholder: '+56 9 1234 5678', 'aria-label': 'WhatsApp (opcional)' });
  const er = el('div', { style: 'color:#b91c1c;font-size:13px;min-height:1em' }); const alertar = m => { er.textContent = m; };
  const envio = el('div');
  const mostrarEnvio = (email, clave, nueva) => envio.replaceChildren(cajaEnvio(email, clave, tel.value, nueva));
  const resetear = async (u) => { er.textContent = ''; const clave = claveAzar();
    try { await T('crearUsuario', u.email, clave, !!u.admin); mostrarEnvio(u.email, clave, true); } catch (e) { er.textContent = e.message; } };
  cont.replaceChildren(el('h2', { text: 'Usuarios' }),
    el('p', { text: 'Uso de los últimos 30 días. Agrega usuarios con una contraseña inicial y entrégasela tú; no se envían correos.' }),
    el('table', {}, el('thead', {}, el('tr', {}, el('th', { text: 'Correo' }), el('th', { text: 'Rol' }), el('th', { text: 'Uso' }), el('th'))), tb),
    el('div', { class: 'row' }, el('label', {}, 'Correo', em), el('label', {}, 'Contraseña', pw), el('label', {}, 'WhatsApp (opcional)', tel),
      el('button', { class: 'pri', text: 'Agregar', onclick: async () => { er.textContent = ''; const correo = em.value.trim().toLowerCase(), clave = pw.value;
        try { await T('crearUsuario', correo, clave, false); await panelUsuarios({ email: correo, clave, tel: tel.value }); } catch (e) { er.textContent = e.message; } } })),
    er, envio, el('div', { style: 'text-align:right;margin-top:10px' }, el('button', { text: 'Cerrar', onclick: () => ov.remove() })));
  if (recien && recien.email) { tel.value = recien.tel || ''; mostrarEnvio(recien.email, recien.clave, false); }
  pw.value = claveAzar();
}
function claveAzar(){ const a = 'abcdefghjkmnpqrstuvwxyz23456789'; const r = new Uint32Array(10); crypto.getRandomValues(r); return Array.from(r, n => a[n % a.length]).join(''); }
function linkApp(){ return location.origin + location.pathname.replace(/index\.html$/, ''); }
function cajaEnvio(email, clave, tel, nueva){
  const txt = 'Hola, ' + (nueva ? 'te dejé una contraseña nueva' : 'te creé una cuenta') + ' en Kits SSO (BERLIAM).\n\nEntra aquí: ' + linkApp() + '\nCorreo: ' + email + '\nContraseña: ' + clave + '\n\nAl entrar, cámbiala con el botón "Cambiar contraseña".';
  let num = String(tel || '').replace(/\D/g, ''); if (num.length === 9 && num[0] === '9') num = '56' + num;
  const wa = 'https://wa.me/' + num + '?text=' + encodeURIComponent(txt);
  const ok = el('span', { style: 'font-size:13px;color:#15803d;margin-left:8px' });
  return el('div', { style: 'margin-top:12px;padding:12px 14px;border:1px solid #86efac;background:#f0fdf4;border-radius:10px' },
    el('p', { style: 'margin:0 0 8px;font-weight:600', text: (nueva ? 'Contraseña nueva para ' : 'Cuenta creada: ') + email }),
    el('pre', { style: 'white-space:pre-wrap;font:13px/1.45 system-ui;margin:0 0 10px;color:#1f2937', text: txt }),
    el('button', { class: 'pri', text: 'Enviar por WhatsApp', onclick: () => window.open(wa, '_blank', 'noopener') }), ' ',
    el('button', { text: 'Copiar mensaje', onclick: async () => { try { await navigator.clipboard.writeText(txt); ok.textContent = 'Copiado'; } catch (_) { ok.textContent = 'No se pudo copiar; selecciónalo a mano.'; } } }), ok);
}
function panelClave(){
  document.getElementById('kxAdm')?.remove();
  const a = el('input', { type: 'password', autocomplete: 'current-password', 'aria-label': 'Contraseña actual' });
  const n = el('input', { type: 'password', autocomplete: 'new-password', 'aria-label': 'Contraseña nueva' });
  const er = el('div', { role: 'alert', style: 'font-size:13px;min-height:1em;margin-top:6px' });
  const ov = el('div', { id: 'kxAdm', onclick: (e) => { if (e.target === ov) ov.remove(); } }, el('div', { class: 'bx', style: 'max-width:420px' },
    el('h2', { text: 'Cambiar contraseña' }),
    el('label', {}, 'Contraseña actual', a), el('label', {}, 'Contraseña nueva (mínimo 8)', n), er,
    el('div', { style: 'text-align:right;margin-top:10px' }, el('button', { text: 'Cancelar', onclick: () => ov.remove() }), ' ',
      el('button', { class: 'pri', text: 'Guardar', onclick: async () => { er.style.color = '#b91c1c'; er.textContent = '';
        try { await T('cambiarClave', a.value, n.value); er.style.color = '#15803d'; er.textContent = 'Listo. Desde ahora entra con la nueva.'; setTimeout(() => ov.remove(), 1500); } catch (e) { er.textContent = e.message; } } }))));
  document.body.appendChild(ov); a.focus();
}
(async () => {
  if (token){ try { yo = await call('yo', token); return entrar(); } catch (_) { token = null; } }
  let hay = true; try { hay = (await call('estado')).hay_admin; } catch (e) { return pantallaLogin(false, 'No se pudo conectar con el servidor.'); }
  pantallaLogin(!hay);
})();
})();
