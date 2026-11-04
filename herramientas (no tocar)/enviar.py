import json, sys
from openpyxl import load_workbook

XLSX, OUT = sys.argv[1], sys.argv[2]
SLUGS = json.load(open(sys.argv[3], encoding='utf-8'))
ws = load_workbook(XLSX)['Prospectos']
filas = []
for r in ws.iter_rows(min_row=2, values_only=True):
    prio, nombre, rubro, zona, tel = r[0], r[1], r[2], r[3], r[4]
    mapa = ws.cell(row=len(filas) + 2, column=11).hyperlink.target
    slug = SLUGS.get(nombre) or next((v for k, v in SLUGS.items() if nombre.startswith(k)), None)
    filas.append({'id': len(filas), 'prio': prio, 'nombre': nombre, 'rubro': rubro.lower(), 'zona': zona, 'tel': tel or '', 'mapa': mapa, 'muestra': slug})

# Por ahora solo los negocios con WhatsApp (celular de 8 dígitos que empieza por 6)
import re as _re
def _movil(t):
    return any(len(d) == 8 and d[0] == '6' for d in (x.replace('-', '').replace(' ', '') for x in _re.findall(r'\d[\d\s-]{6,}\d', str(t))))
filas = [f for f in filas if _movil(f['tel'])]
for i, f in enumerate(filas): f['id'] = 'w' + str(i) + '-' + (f['muestra'] or '')

HTML = r'''<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Enviar mensajes a negocios - D.E.K NovaCore</title>
<style>
:root{--bg:#0a0d14;--panel:#121826;--line:#243047;--text:#e8ecf4;--muted:#8a94a8;--acc:#5a9bfc;--ok:#34d399;--warn:#fbbf24}
*{box-sizing:border-box}body{margin:0;font:15px/1.5 system-ui,Segoe UI,Arial;background:var(--bg);color:var(--text)}
header{position:sticky;top:0;z-index:2;background:#0a0d14ee;backdrop-filter:blur(6px);border-bottom:1px solid var(--line);padding:16px 20px}
h1{margin:0 0 4px;font-size:20px}.muted{color:var(--muted)}.sm{font-size:13px}
.cfg{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin-top:12px}
label{display:flex;flex-direction:column;gap:4px;font-size:12px;color:var(--muted)}
input,select,textarea{font:inherit;color:var(--text);background:#0d1320;border:1px solid var(--line);border-radius:10px;padding:9px 11px}
textarea{min-height:150px;resize:vertical;width:100%}
.bar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:10px}
.chip{border:1px solid var(--line);background:none;color:var(--text);border-radius:99px;padding:6px 12px;cursor:pointer}
.chip.on{background:var(--acc);border-color:var(--acc);color:#05080f}
main{padding:16px 20px 60px;display:grid;gap:12px;max-width:1100px;margin:auto}
.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:14px 16px}
.card.hecho{opacity:.55}.top{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap}
.tag{display:inline-block;font-size:12px;border-radius:99px;padding:2px 9px;background:#1b2436;color:var(--muted);margin-right:6px}
.p1{background:#123524;color:var(--ok)}.p2{background:#3a3112;color:var(--warn)}
.acts{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px;align-items:center}
.btn{border:0;border-radius:10px;padding:9px 14px;font-weight:600;cursor:pointer;text-decoration:none;display:inline-flex;gap:6px;align-items:center}
.wa{background:#25d366;color:#05210f}.alt{background:#1b2436;color:var(--text)}.ok{background:none;border:1px solid var(--line);color:var(--muted)}
.wa[aria-disabled=true]{opacity:.4;pointer-events:none}
.aviso{background:#2a1d0b;border:1px solid #6b4f1a;border-radius:12px;padding:10px 14px;font-size:13px}
details summary{cursor:pointer;color:var(--acc);font-size:13px;margin-top:8px}
.num{max-width:170px}
</style></head><body>
<header>
  <h1>Negocios con WhatsApp</h1>
  <div class="muted sm">Pulsa <b>Enviar por WhatsApp</b>: se abre WhatsApp con el mensaje listo y tú solo das a enviar. Lo que marques como enviado se guarda en este navegador.</div>
  <div class="cfg">
    <label>Tu nombre (va en el saludo)<input id="yo" placeholder="Ej: Edwin"></label>
    <label>Enlace de tu portafolio (Vercel)<input id="link" placeholder="https://tu-portafolio.vercel.app"></label>
    <label>Qué ofrecer<select id="oferta">
      <option value="todo">Las 3 opciones (catálogo, web y sistema)</option>
      <option value="catalogo">Solo catálogo digital</option>
      <option value="web">Solo página web</option>
      <option value="sistema">Solo sistema a medida</option></select></label>
  </div>
  <div class="bar" id="filtros"></div>
</header>
<main>
  <div class="aviso">Para no tener problemas con WhatsApp: <b>máximo 15-20 mensajes por día</b>, escribe primero a los que tienen celular (empieza por 6) y espera unos minutos entre uno y otro. Los teléfonos que empiezan por 4 son fijos: ahí no hay WhatsApp; llama o pasa por el local (abajo de cada uno tienes qué decir).</div>
  <div class="aviso" id="falta-link" style="border-color:#7f1d1d;background:#2a0f12">Primero escribe arriba el <b>enlace de tu portafolio</b> (el de Vercel). Hasta entonces los botones de WhatsApp están apagados, para que no salga un mensaje sin enlace.</div>
  <div id="lista"></div>
</main>
<script>
const NEG = __DATA__;
const MUESTRAS = 'https://edwincedexab4356-boop.github.io/clientes-prospectos/muestras/';
const SIS = {
  taller: 'un sistema para llevar cada trabajo, lo que el cliente abonó, lo que le deben y lo que gasta en repuestos',
  ferreteria: 'un sistema de caja e inventario que avisa cuando algo se acaba, con fiado de clientes y reportes de ventas',
  repuestos: 'un sistema de inventario y ventas para saber qué tiene, qué se vende más y cuánto gana',
  panaderia: 'un sistema para llevar pedidos, ventas del día y gastos, y ver cuánto gana de verdad',
  restaurante: 'un sistema para llevar pedidos, ventas del día y gastos, y ver cuánto gana de verdad',
  cafe: 'un sistema para llevar ventas del día y gastos, y ver cuánto gana de verdad',
  belleza: 'un sistema de citas y clientes, con lo que cobra cada día',
  peluqueria: 'un sistema de citas y clientes, con lo que cobra cada día',
  dentista: 'un sistema de citas, pacientes y cobros',
  clinica: 'un sistema de citas, pacientes y cobros',
  lavanderia: 'un sistema para llevar cada orden, cuándo se entrega y quién debe',
  eventos: 'un sistema para llevar cada evento, el adelanto, lo que falta por cobrar y lo que gastó',
  pinturas: 'un sistema de inventario y ventas',
  lavado: 'un sistema para llevar los servicios del día, cobros y gastos',
  ropa: 'un sistema de inventario y ventas',
  otro: 'un sistema para llevar sus ventas, cobros y gastos',
};
const CAT = { taller: 'sus servicios y precios', ferreteria: 'sus productos con foto y precio', repuestos: 'sus repuestos con foto y precio',
  panaderia: 'su menú con fotos y precios', restaurante: 'su menú con fotos y precios', cafe: 'su menú con fotos y precios',
  belleza: 'sus servicios y precios', peluqueria: 'sus cortes y precios', dentista: 'sus servicios', clinica: 'sus servicios',
  lavanderia: 'sus servicios y precios', eventos: 'sus paquetes y fotos de sus eventos', pinturas: 'sus productos con foto y precio',
  lavado: 'sus servicios y precios', ropa: 'su ropa con foto y precio', otro: 'sus productos con foto y precio' };
const $ = (s) => document.querySelector(s);
const st = JSON.parse(localStorage.getItem('dek-envios') || '{}');
const cfg = JSON.parse(localStorage.getItem('dek-envios-cfg') || '{}');
if (!cfg.link) cfg.link = 'https://edwin2030-cmd.github.io/dek-portafolio/';
const guardar = () => { localStorage.setItem('dek-envios', JSON.stringify(st)); localStorage.setItem('dek-envios-cfg', JSON.stringify(cfg)); };
['yo', 'link', 'oferta'].forEach((k) => { if (cfg[k]) $('#' + k).value = cfg[k]; $('#' + k).addEventListener('input', () => { cfg[k] = $('#' + k).value; guardar(); pintar(); }); });
let filtro = 'pendientes';

function numeros(t) {
  return (String(t).match(/\d[\d\s-]{6,}\d/g) || []).map((x) => x.replace(/\D/g, '')).map((d) => d.startsWith('507') ? d.slice(3) : d);
}
const celular = (n) => numeros(n).find((d) => d.length === 8 && d[0] === '6');
const bonito = (n) => n.slice(0, 4) + '-' + n.slice(4);
function limpiaNombre(n) { return n.replace(/\s*\(.*\)$/, ''); }

function mensaje(n) {
  const yo = cfg.yo ? `Soy ${cfg.yo}, de D.E.K NovaCore` : 'Le escribe D.E.K NovaCore';
  const link = (cfg.link || '[ENLACE]').replace(/\/?$/, '/');
  const r = n.rubro, o = cfg.oferta || 'todo';
  const cat = `un catálogo digital con ${CAT[r]}, donde el cliente arma su pedido y le llega a su WhatsApp ya escrito`;
  const web = 'una página web para que lo encuentren en Google, con su ubicación, horario y botón de WhatsApp';
  const sis = SIS[r];
  let cuerpo;
  if (o === 'catalogo') cuerpo = `Le puedo hacer ${cat}. Desde $450.`;
  else if (o === 'web') cuerpo = `Le puedo hacer ${web}. Desde $200.`;
  else if (o === 'sistema') cuerpo = `Le puedo hacer ${sis}. Funciona en el celular y la computadora.`;
  else cuerpo = `Ayudo a negocios de aquí con tres cosas, y puede tomar una o las tres juntas con precio especial:\n\n1) *Catálogo digital*: ${CAT[r]} en línea; el cliente arma su pedido y le llega a su WhatsApp ya escrito.\n2) *Página web*: para que lo encuentren en Google, con su ubicación, horario y botón de WhatsApp.\n3) *Sistema a medida*: ${sis.replace(/^un sistema /, '')}.`;
  return `Hola, buenas. ¿Hablo con *${limpiaNombre(n.nombre)}*? ${yo}, aquí en Colón.\n\n${cuerpo}${n.muestra ? `\n\nYa le preparé una *muestra con el nombre de su negocio* para que vea cómo quedaría: ${MUESTRAS}${n.muestra}/` : ''}\n\nEsta semana tengo *precio de lanzamiento* para los primeros 4 negocios: página $150, catálogo $350 y el primer mes del sistema gratis.\n\nMis trabajos: ${link}\n\n${n.muestra ? '¿Qué le parece? Si le gusta, se la dejo lista con sus fotos y precios.' : '¿Le gustaría que le prepare una muestra para su negocio? Sin compromiso.'}`;
}
function guion(n) {
  return `Buenas, ¿hablo con ${limpiaNombre(n.nombre)}? Mi nombre es ${cfg.yo || '[tu nombre]'}, de D.E.K NovaCore, aquí en Colón. Hago catálogos digitales, páginas web y sistemas para negocios. ¿Con quién puedo hablar sobre eso? ... Le quería ofrecer ${SIS[n.rubro]}, y una página o catálogo para que le pidan por WhatsApp. ¿Me da un WhatsApp para mandarle ejemplos?`;
}

function pintar() {
  $('#falta-link').style.display = cfg.link ? 'none' : '';
  const cuenta = { pendientes: 0, celular: 0, enviados: 0, todos: NEG.length };
  NEG.forEach((n) => { const num = st[n.id]?.num || celular(n.tel); if (!st[n.id]?.ok) cuenta.pendientes++; else cuenta.enviados++; if (num && !st[n.id]?.ok) cuenta.celular++; });
  $('#filtros').innerHTML = [['celular', 'Con WhatsApp'], ['pendientes', 'Pendientes'], ['enviados', 'Enviados'], ['todos', 'Todos']]
    .map(([k, l]) => `<button class="chip ${filtro === k ? 'on' : ''}" data-f="${k}">${l} (${cuenta[k]})</button>`).join('');
  const vis = NEG.filter((n) => {
    const num = st[n.id]?.num || celular(n.tel);
    return filtro === 'todos' || (filtro === 'enviados' ? st[n.id]?.ok : !st[n.id]?.ok && (filtro !== 'celular' || num));
  });
  const sc = (n) => (st[n.id]?.num || celular(n.tel) ? 2 : 0) + (n.muestra ? 1 : 0);
  vis.sort((a, b) => sc(b) - sc(a) || a.prio - b.prio);
  $('#lista').innerHTML = vis.map((n) => {
    const num = st[n.id]?.num || celular(n.tel);
    const fijos = numeros(n.tel).filter((d) => d[0] !== '6');
    const msg = st[n.id]?.msg || mensaje(n);
    const listo = num && cfg.link;
    const wa = listo ? `https://wa.me/507${num}?text=${encodeURIComponent(msg)}` : '#';
    return `<div class="card ${st[n.id]?.ok ? 'hecho' : ''}" data-id="${n.id}">
      <div class="top"><div><b>${n.nombre}</b><div class="muted sm">${n.zona}</div></div>
        <div><span class="tag ${n.prio === 1 ? 'p1' : n.prio === 2 ? 'p2' : ''}">Prioridad ${n.prio}</span><span class="tag">${n.rubro}</span>${n.muestra ? `<a class="tag p1" href="${MUESTRAS}${n.muestra}/" target="_blank" rel="noopener">Ver su muestra</a>` : ''}</div></div>
      <div class="acts">
        <a class="btn wa" href="${wa}" target="_blank" rel="noopener" aria-disabled="${!listo}" data-env="${n.id}" title="${num && !cfg.link ? 'Primero pon el enlace de tu portafolio arriba' : ''}">Enviar por WhatsApp${num ? ' · ' + bonito(num) : ''}</a>
        ${fijos.length ? fijos.map((f) => `<a class="btn alt" href="tel:+507${f}">Llamar ${f.slice(0, 3)}-${f.slice(3)}</a>`).join('') : ''}
        <a class="btn alt" href="${n.mapa}" target="_blank" rel="noopener">Ver en Google Maps</a>
        <input class="num" data-num="${n.id}" placeholder="WhatsApp: 6xxx-xxxx" value="${st[n.id]?.num ? bonito(st[n.id].num) : ''}">
        <button class="btn ok" data-ok="${n.id}">${st[n.id]?.ok ? '✓ Enviado (deshacer)' : 'Marcar enviado'}</button>
      </div>
      <details><summary>Ver o cambiar el mensaje${fijos.length && !num ? ' · qué decir si llamas' : ''}</summary>
        <textarea data-msg="${n.id}">${msg.replace(/</g, '&lt;')}</textarea>
        ${fijos.length ? `<p class="muted sm"><b>Si llamas:</b> ${guion(n)}</p>` : ''}
      </details>
    </div>`;
  }).join('') || '<p class="muted">No hay negocios en este filtro.</p>';
}
document.addEventListener('click', (e) => {
  const f = e.target.closest('[data-f]'); if (f) { filtro = f.dataset.f; pintar(); return; }
  const ok = e.target.closest('[data-ok]'); if (ok) { const id = ok.dataset.ok; st[id] = { ...st[id], ok: !st[id]?.ok, fecha: new Date().toISOString().slice(0, 10) }; guardar(); pintar(); return; }
  const env = e.target.closest('[data-env]'); if (env && env.getAttribute('aria-disabled') !== 'true') { const id = env.dataset.env; setTimeout(() => { st[id] = { ...st[id], ok: true, fecha: new Date().toISOString().slice(0, 10) }; guardar(); pintar(); }, 800); }
});
document.addEventListener('change', (e) => {
  const n = e.target.dataset.num; if (n !== undefined) { const d = e.target.value.replace(/\D/g, '').replace(/^507/, ''); st[n] = { ...st[n], num: d.length === 8 && d[0] === '6' ? d : undefined }; if (d && !st[n].num) alert('Ese número no parece un celular de Panamá (8 dígitos que empiezan por 6).'); guardar(); pintar(); }
  const m = e.target.dataset.msg; if (m !== undefined) { st[m] = { ...st[m], msg: e.target.value }; guardar(); pintar(); }
});
pintar();
</script></body></html>'''
open(OUT, 'w', encoding='utf-8').write(HTML.replace('__DATA__', json.dumps(filas, ensure_ascii=False)))
print(len(filas), 'negocios')
