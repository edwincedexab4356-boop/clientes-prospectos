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
def _region(z):
    z = str(z)
    if _re.search(r'Arraij|Chorrera', z): return 'Panamá Oeste'
    if _re.search(r'San Miguelito|Ciudad de Panamá|(^|, )Panamá$', z): return 'Ciudad de Panamá'
    return 'Colón'
_ids = set()
for i, f in enumerate(filas):
    f['id'] = f['muestra'] or ('w' + str(i))
    if f['id'] in _ids: f['id'] += '-' + f['tel'].replace('-', '')
    _ids.add(f['id'])
    f['region'] = _region(f['zona'])

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
.card.hecho{opacity:.55}.card.toca{border-color:#6b4f1a}.btn.seg{background:#b7791f;color:#fff}.btn.resp{background:#1f3b6b;color:#fff}.tag.ok{background:#0f2a1c;color:var(--ok)}.top{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap}
.tag{display:inline-block;font-size:12px;border-radius:99px;padding:2px 9px;background:#1b2436;color:var(--muted);margin-right:6px}
.p1{background:#123524;color:var(--ok)}.p2{background:#3a3112;color:var(--warn)}
.acts{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px;align-items:center}
.btn{border:0;border-radius:10px;padding:9px 14px;font-weight:600;cursor:pointer;text-decoration:none;display:inline-flex;gap:6px;align-items:center}
.wa{background:#25d366;color:#05210f}.alt{background:#1b2436;color:var(--text)}.ok{background:none;border:1px solid var(--line);color:var(--muted)}
.wa[aria-disabled=true]{opacity:.4;pointer-events:none}
.aviso{background:#2a1d0b;border:1px solid #6b4f1a;border-radius:12px;padding:10px 14px;font-size:13px}
details summary{cursor:pointer;color:var(--acc);font-size:13px;margin-top:8px}
.num{max-width:170px}
#cupo{margin-top:10px;padding:8px 12px;border-radius:10px;font-size:14px;background:#0f2a1c;border:1px solid #1f6b46}
#cupo.espera{background:#2a1d0b;border-color:#6b4f1a}#cupo.lleno{background:#2a0f12;border-color:#7f1d1d}
#cupo .barra{height:6px;border-radius:3px;background:#243047;margin-top:6px;overflow:hidden}#cupo .barra i{display:block;height:100%;background:var(--ok)}
.btn.wa.bloq{opacity:.45;cursor:not-allowed}
.btn.cli{background:#5b3a8a;color:#fff}.tag.cli{background:#2d1f47;color:#c4a5ff}
.rapidas{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px;align-items:center}.rapidas a{font-size:12px;padding:6px 10px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:12px 16px}
.panel summary{font-weight:600;color:var(--text);font-size:15px;margin:0}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin:12px 0}
.kpi{background:#0d1320;border:1px solid var(--line);border-radius:12px;padding:10px}.kpi b{display:block;font-size:22px}
table{width:100%;border-collapse:collapse;font-size:13px}td,th{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line)}
.diag{background:#0f1a2e;border:1px solid #1f3b6b;border-radius:10px;padding:10px 12px;font-size:14px;margin-top:10px}
.copiar{font-size:12px;padding:5px 10px}.rq{border-top:1px solid var(--line);padding:10px 0}.rq p{white-space:pre-wrap;margin:6px 0;font-size:13px}
</style></head><body>
<header>
  <h1>Negocios con WhatsApp <a class="btn alt" style="float:right;font-size:13px" href="ruta-visitas.html" target="_blank">🚶 Ruta de visitas en Colón</a></h1>
  <div class="muted sm">Pulsa <b>Enviar por WhatsApp</b>: se abre WhatsApp con el mensaje listo y tú solo das a enviar. Lo que marques como enviado se guarda en este navegador. A los 3 días, los que no te respondieron aparecen en <b>⏰ Para seguimiento</b>; cuando alguien te conteste, pulsa <b>💬 Me respondió</b>.</div>
  <div class="cfg">
    <label>Tu nombre (va en el saludo)<input id="yo" placeholder="Ej: Edwin"></label>
    <label>Enlace de tu portafolio (Vercel)<input id="link" placeholder="https://tu-portafolio.vercel.app"></label>
    <label>Qué ofrecer<select id="oferta">
      <option value="todo">Las 3 opciones (catálogo, web y sistema)</option>
      <option value="catalogo">Solo catálogo digital</option>
      <option value="web">Solo página web</option>
      <option value="sistema">Solo sistema a medida</option></select></label>
    <label>Tipo de negocio (mejor uno a la vez)<select id="rubro"><option value="">Todos los tipos</option></select></label>
    <label>Tu primer cliente (para mencionarlo)<input id="ref" placeholder="Ej: Taller MC Place"></label>
  </div>
  <div class="bar" id="zonas"></div>
  <div class="bar" id="filtros"></div>
  <div id="cupo"></div>
</header>
<main>
  <div class="aviso">Para no tener problemas con WhatsApp: <b>máximo 40 mensajes por día</b> (la página no te deja pasar de ahí) y al menos 90 segundos entre uno y otro. Escribe primero a los que tienen celular (empieza por 6). Los teléfonos que empiezan por 4 son fijos: ahí no hay WhatsApp; llama o pasa por el local (abajo de cada uno tienes qué decir), o mejor, visítalos con la <b>Ruta de visitas</b>.</div>
  <div class="aviso" id="falta-link" style="border-color:#7f1d1d;background:#2a0f12">Primero escribe arriba el <b>enlace de tu portafolio</b> (el de Vercel). Hasta entonces los botones de WhatsApp están apagados, para que no salga un mensaje sin enlace.</div>
  <details class="panel" id="stats"></details>
  <details class="panel" id="rapidas-panel"><summary>💬 Respuestas listas (para cuando te contesten)</summary><div id="rapidas-lista"></div></details>
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
  veterinaria: 'un sistema de citas, mascotas y cobros',
  gimnasio: 'un sistema para llevar las mensualidades: quién pagó y quién debe',
  hotel: 'un sistema de reservas y cobros',
  tours: 'un sistema de reservas, adelantos y cobros',
  abogados: 'un sistema de clientes, casos, citas y honorarios',
  contable: 'un sistema de clientes, trámites y cobros',
  sala: 'un sistema para llevar cada evento, el adelanto, lo que falta por cobrar y lo que gastó',
};
const CAT = { taller: 'sus servicios y precios', ferreteria: 'sus productos con foto y precio', repuestos: 'sus repuestos con foto y precio',
  panaderia: 'su menú con fotos y precios', restaurante: 'su menú con fotos y precios', cafe: 'su menú con fotos y precios',
  belleza: 'sus servicios y precios', peluqueria: 'sus cortes y precios', dentista: 'sus servicios', clinica: 'sus servicios',
  lavanderia: 'sus servicios y precios', eventos: 'sus paquetes y fotos de sus eventos', pinturas: 'sus productos con foto y precio',
  lavado: 'sus servicios y precios', ropa: 'su ropa con foto y precio', otro: 'sus productos con foto y precio',
  veterinaria: 'sus servicios y precios', gimnasio: 'sus planes y precios', hotel: 'sus habitaciones con fotos y precios', tours: 'sus tours con fotos y precios',
  abogados: 'sus servicios legales', contable: 'sus servicios', sala: 'sus paquetes con fotos y precios' };
const $ = (s) => document.querySelector(s);
const st = JSON.parse(localStorage.getItem('dek-envios') || '{}');
// Migra marcas viejas ("w12-slug") al id nuevo ("slug")
Object.keys(st).forEach((k) => { const m = k.match(/^w\d+-(.+)$/); if (m) { st[m[1]] = { ...st[k], ...st[m[1]] }; delete st[k]; } });
localStorage.setItem('dek-envios', JSON.stringify(st));
const cfg = JSON.parse(localStorage.getItem('dek-envios-cfg') || '{}');
if (!cfg.link) cfg.link = 'https://edwin2030-cmd.github.io/dek-portafolio/';
const guardar = () => { localStorage.setItem('dek-envios', JSON.stringify(st)); localStorage.setItem('dek-envios-cfg', JSON.stringify(cfg)); };
['yo', 'link', 'oferta', 'ref'].forEach((k) => { if (cfg[k]) $('#' + k).value = cfg[k]; $('#' + k).addEventListener('input', () => { cfg[k] = $('#' + k).value; guardar(); pintar(); }); });
let filtro = 'pendientes', limite = 40;
const TOPE = 40, ESPERA = 90; // mensajes por día y segundos mínimos entre mensajes
const hoy = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
const enviadosHoy = () => Object.values(st).reduce((a, s) => a + (s && s.ok && s.fecha === hoy() ? 1 : 0) + (s && s.seg === hoy() ? 1 : 0), 0);
const DIAS_SEG = 3; // días sin respuesta antes del seguimiento
const dias = (f) => f ? Math.round((new Date(hoy()) - new Date(f)) / 864e5) : 0;
const toca = (n) => { const s = st[n.id]; return !!(s && s.ok && !s.resp && !s.seg && dias(s.fecha) >= DIAS_SEG); };
function seguimiento(n) {
  return `Hola, buenas. Le escribí hace unos días ${n.muestra ? `sobre la muestra que le preparé para *${limpiaNombre(n.nombre)}*: ${MUESTRAS}${n.muestra}/` : `sobre una página o catálogo para *${limpiaNombre(n.nombre)}*.`}

¿Pudo verla? Si le interesa, se la dejo lista con sus fotos y precios. Y si no es el momento, no hay problema, ¡gracias!`;
}
const faltaEspera = () => Math.max(0, Math.ceil(((cfg.ult || 0) + ESPERA * 1000 - Date.now()) / 1000));
function cupo() {
  const n = enviadosHoy(), w = faltaEspera(), el = $('#cupo');
  const lleno = n >= TOPE;
  el.className = lleno ? 'lleno' : w ? 'espera' : '';
  el.innerHTML = `<b>Hoy: ${n} de ${TOPE} mensajes.</b> ` + (lleno ? 'Llegaste al límite de hoy. Mañana sigues; así cuidas tu WhatsApp.' : w ? `Espera ${w} s antes del siguiente (para que WhatsApp no te marque como spam).` : `Puedes enviar el siguiente. Te quedan ${TOPE - n} hoy.`) + `<div class="barra"><i style="width:${Math.min(100, n / TOPE * 100)}%"></i></div>`;
  document.querySelectorAll('[data-env]').forEach((b) => b.classList.toggle('bloq', lleno || w > 0));
}
setInterval(cupo, 1000);
const RUBROS = [...new Set(NEG.map((n) => n.rubro))].sort();
$('#rubro').innerHTML += RUBROS.map((r) => `<option value="${r}">${r[0].toUpperCase() + r.slice(1)} (${NEG.filter((n) => n.rubro === r).length})</option>`).join('');
$('#rubro').value = cfg.rubro || ''; $('#rubro').addEventListener('change', () => { cfg.rubro = $('#rubro').value; limite = 40; guardar(); pintar(); });
const ZONAS = ['Colón', 'Panamá Oeste', 'Ciudad de Panamá', 'Todas'];
if (!cfg.zona) cfg.zona = 'Colón';

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
  return `Hola, buenas. ¿Hablo con *${limpiaNombre(n.nombre)}*? ${yo}, ${n.region === 'Colón' ? 'aquí en Colón' : 'desde Colón (trabajo con negocios de todo Panamá)'}.${cfg.ref ? ` Ya trabajo con *${cfg.ref}*.` : ''}\n\n${cuerpo}${n.muestra ? `\n\nYa le preparé una *muestra con el nombre de su negocio* para que vea cómo quedaría: ${MUESTRAS}${n.muestra}/` : ''}\n\nEsta semana tengo *precio de lanzamiento* para los primeros 4 negocios: página $150, catálogo $350 y el primer mes del sistema gratis. Paga la mitad al empezar y la otra mitad al entregar (Yappy o transferencia).\n\nMis trabajos: ${link}\n\n${n.muestra ? '¿Qué le parece? Si le gusta, se la dejo lista con sus fotos y precios.' : '¿Le gustaría que le prepare una muestra para su negocio? Sin compromiso.'}`;
}
const RESP = [
  ['precio', '💲 Precio', (n) => `¡Con gusto! Para *${limpiaNombre(n.nombre)}*:\n\n• Página web: *$150* (precio de lanzamiento; normal $200)\n• Catálogo con pedidos por WhatsApp: *$350* (normal $450)\n• Sistema: ${SIS[n.rubro] || SIS.otro}. *Primer mes gratis*, después $20 al mes.\n\nPaga la mitad para empezar y la otra mitad cuando esté funcionando, por Yappy o transferencia. Tiene 30 días de garantía.\n\n¿Le llamo 5 minutos y le explico cuál le conviene más?`],
  ['llamada', '📞 Pedir llamada', () => `¿Le puedo llamar 5 minutos hoy y se lo explico mejor? ¿A qué hora le queda bien?`],
  ['pensar', '🤔 "Lo voy a pensar"', (n) => `¡Claro, sin presión! Le dejo la muestra para que la vea con calma${n.muestra ? `: ${MUESTRAS}${n.muestra}/` : '.'}\n\nEl precio de lanzamiento se lo guardo hasta el domingo. Cualquier duda me escribe por aquí.`],
  ['caro', '💸 "Está caro"', () => `Le entiendo. Para que no sea un gasto de golpe, lo puede pagar en dos partes: la mitad para empezar y la otra mitad cuando ya esté funcionando.\n\nY piense que con *un solo cliente nuevo* al mes que lo encuentre por ahí, ya se paga. Si prefiere, empezamos solo con la página de $150 y luego le agregamos lo demás.`],
  ['si', '✅ Dijo que sí', () => `¡Excelente, gracias por la confianza! 🙌 Para dejarla lista necesito:\n\n1) Su logo (si tiene)\n2) Fotos de sus productos, servicios o del local\n3) La lista de precios\n4) Horario y dirección\n\nMe lo puede mandar por aquí. Con el 50% (Yappy o transferencia) empiezo hoy y en unos 5 días la tiene funcionando.`],
  ['mensual', '🔁 Mantenimiento $15/mes', () => `Además, si quiere, yo me encargo de mantenerla: cambiar precios, fotos y promociones cuando usted me diga, y que siempre esté en línea. Son *$15 al mes* y usted no se preocupa de nada.`],
  ['referir', '🤝 Pedir recomendados', () => `Gracias otra vez por confiar en mí. Si conoce otro negocio al que le sirva algo así, recomiéndeme: *por cada uno que me contrate, le regalo un mes de mantenimiento*.`],
  ['testimonio', '⭐ Pedir opinión', () => `¿Me haría un favor? Escríbame en una o dos líneas qué le pareció el trabajo, y dígame si me deja mostrar su página como ejemplo a otros negocios. Me ayuda muchísimo. ¡Gracias!`],
];
const EJ = { nombre: 'su negocio', rubro: 'otro', muestra: '' };
function pintarRapidas() {
  $('#rapidas-lista').innerHTML = '<p class="muted sm">En cada negocio que marques como <b>💬 Me respondió</b> salen estos botones con su nombre y su WhatsApp. Aquí los puedes copiar.</p>' +
    RESP.map(([k, l, f]) => `<div class="rq"><b>${l}</b> <button class="btn alt copiar" data-copiar="${k}">Copiar</button><p>${f(EJ).replace(/</g, '&lt;')}</p></div>`).join('');
}
function pintarStats() {
  const env = NEG.filter((n) => st[n.id]?.ok), resp = env.filter((n) => st[n.id]?.resp), cli = NEG.filter((n) => st[n.id]?.cli);
  const pct = (a, b) => b ? Math.round(a / b * 100) + '%' : '—';
  const por = {}; env.forEach((n) => { const r = por[n.rubro] = por[n.rubro] || [0, 0, 0]; r[0]++; if (st[n.id]?.resp) r[1]++; if (st[n.id]?.cli) r[2]++; });
  const filas = Object.entries(por).sort((a, b) => b[1][1] / b[1][0] - a[1][1] / a[1][0] || b[1][0] - a[1][0]);
  let diag;
  if (env.length < 50) diag = `Todavía son pocos (${env.length}) para sacar conclusiones. Cuando pases de 50 enviados, aquí te digo qué cambiar.`;
  else if (resp.length / env.length < 0.05) diag = `Te responde menos del 5%: <b>cambia el mensaje</b>. Prueba en "Qué ofrecer" con una sola opción (más corto) y empieza por los tipos de negocio que más responden en la tabla.`;
  else if (resp.length >= 10 && cli.length / resp.length < 0.1) diag = `Te responden, pero pocos compran: el problema está en <b>cómo cierras</b>. Pide siempre la llamada de 5 minutos (botón 📞) y ofrece el pago en dos partes.`;
  else diag = `Vas bien. Sigue con los tipos de negocio que más responden y no olvides los seguimientos.`;
  const abierto = $('#stats').open;
  $('#stats').innerHTML = `<summary>📊 Tus resultados: ${env.length} enviados · ${resp.length} respondieron · ${cli.length} clientes</summary>
    <div class="kpis"><div class="kpi"><b>${env.length}</b>enviados</div><div class="kpi"><b>${pct(resp.length, env.length)}</b>te responden</div>
    <div class="kpi"><b>${pct(cli.length, resp.length)}</b>de los que responden compran</div><div class="kpi"><b>${cli.length}</b>clientes</div></div>
    ${filas.length ? `<table><tr><th>Tipo de negocio</th><th>Enviados</th><th>Respondieron</th><th>Clientes</th></tr>${filas.map(([r, [a, b, c]]) => `<tr><td>${r}</td><td>${a}</td><td>${b} (${pct(b, a)})</td><td>${c}</td></tr>`).join('')}</table>` : ''}
    <div class="diag">${diag}</div>`;
  $('#stats').open = abierto;
}
function guion(n) {
  return `Buenas, ¿hablo con ${limpiaNombre(n.nombre)}? Mi nombre es ${cfg.yo || '[tu nombre]'}, de D.E.K NovaCore, ${n.region === 'Colón' ? 'aquí en Colón' : 'de Colón'}. Hago catálogos digitales, páginas web y sistemas para negocios. ¿Con quién puedo hablar sobre eso? ... Le quería ofrecer ${SIS[n.rubro]}, y una página o catálogo para que le pidan por WhatsApp. ¿Me da un WhatsApp para mandarle ejemplos?`;
}

function pintar() {
  setTimeout(cupo);
  $('#falta-link').style.display = cfg.link ? 'none' : '';
  $('#zonas').innerHTML = ZONAS.map((z) => `<button class="chip ${cfg.zona === z ? 'on' : ''}" data-z="${z}">${z} (${z === 'Todas' ? NEG.length : NEG.filter((n) => n.region === z).length})</button>`).join('');
  const BASE = NEG.filter((n) => (cfg.zona === 'Todas' || n.region === cfg.zona) && (!cfg.rubro || n.rubro === cfg.rubro));
  pintarStats();
  const cuenta = { seguimiento: BASE.filter(toca).length, respondieron: BASE.filter((n) => st[n.id]?.resp && !st[n.id]?.cli).length, clientes: BASE.filter((n) => st[n.id]?.cli).length, pendientes: 0, celular: 0, enviados: 0, todos: BASE.length };
  BASE.forEach((n) => { const num = st[n.id]?.num || celular(n.tel); if (!st[n.id]?.ok) cuenta.pendientes++; else cuenta.enviados++; if (num && !st[n.id]?.ok) cuenta.celular++; });
  $('#filtros').innerHTML = [['seguimiento', '⏰ Para seguimiento'], ['respondieron', '💬 Respondieron'], ['clientes', '⭐ Clientes'], ['celular', 'Con WhatsApp'], ['pendientes', 'Pendientes'], ['enviados', 'Enviados'], ['todos', 'Todos']]
    .map(([k, l]) => `<button class="chip ${filtro === k ? 'on' : ''}" data-f="${k}">${l} (${cuenta[k]})</button>`).join('');
  const vis = BASE.filter((n) => {
    const num = st[n.id]?.num || celular(n.tel);
    if (filtro === 'seguimiento') return toca(n);
    if (filtro === 'respondieron') return !!st[n.id]?.resp && !st[n.id]?.cli;
    if (filtro === 'clientes') return !!st[n.id]?.cli;
    return filtro === 'todos' || (filtro === 'enviados' ? st[n.id]?.ok : !st[n.id]?.ok && (filtro !== 'celular' || num));
  });
  const sc = (n) => (st[n.id]?.num || celular(n.tel) ? 2 : 0) + (n.muestra ? 1 : 0);
  vis.sort((a, b) => sc(b) - sc(a) || a.prio - b.prio);
  $('#lista').innerHTML = vis.slice(0, limite).map((n) => {
    const num = st[n.id]?.num || celular(n.tel);
    const fijos = numeros(n.tel).filter((d) => d[0] !== '6');
    const s = st[n.id] || {}, sg = toca(n);
    const msg = sg ? seguimiento(n) : s.msg || mensaje(n);
    const listo = num && cfg.link;
    const wa = listo ? `https://wa.me/507${num}?text=${encodeURIComponent(msg)}` : '#';
    const estado = s.cli ? '<span class="tag cli">⭐ Cliente</span>' : s.resp ? '<span class="tag ok">💬 Respondió</span>' : s.seg ? `<span class="tag">Seguimiento enviado hace ${dias(s.seg)} día(s)</span>` : s.ok ? `<span class="tag">Enviado hace ${dias(s.fecha)} día(s)</span>` : '';
    return `<div class="card ${sg ? 'toca' : s.ok ? 'hecho' : ''}" data-id="${n.id}">
      <div class="top"><div><b>${n.nombre}</b><div class="muted sm">${n.zona}</div></div>
        <div><span class="tag ${n.prio === 1 ? 'p1' : n.prio === 2 ? 'p2' : ''}">Prioridad ${n.prio}</span><span class="tag">${n.rubro}</span>${estado}${n.muestra ? `<a class="tag p1" href="${MUESTRAS}${n.muestra}/" target="_blank" rel="noopener">Ver su muestra</a>` : ''}</div></div>
      <div class="acts">
        <a class="btn wa ${sg ? 'seg' : ''}" href="${wa}" target="_blank" rel="noopener" aria-disabled="${!listo}" data-env="${n.id}" ${sg ? 'data-seg' : ''} title="${num && !cfg.link ? 'Primero pon el enlace de tu portafolio arriba' : ''}">${sg ? 'Enviar seguimiento' : 'Enviar por WhatsApp'}${num ? ' · ' + bonito(num) : ''}</a>
        ${fijos.length ? fijos.map((f) => `<a class="btn alt" href="tel:+507${f}">Llamar ${f.slice(0, 3)}-${f.slice(3)}</a>`).join('') : ''}
        <a class="btn alt" href="${n.mapa}" target="_blank" rel="noopener">Ver en Google Maps</a>
        <input class="num" data-num="${n.id}" placeholder="WhatsApp: 6xxx-xxxx" value="${st[n.id]?.num ? bonito(st[n.id].num) : ''}">
        <button class="btn ok" data-ok="${n.id}">${st[n.id]?.ok ? '✓ Enviado (deshacer)' : 'Marcar enviado'}</button>
        ${s.ok ? `<button class="btn resp" data-resp="${n.id}">${s.resp ? '↩ Quitar "respondió"' : '💬 Me respondió'}</button>` : ''}
        ${s.resp ? `<button class="btn cli" data-cli="${n.id}">${s.cli ? '↩ Quitar "cliente"' : '⭐ Ya es cliente'}</button>` : ''}
      </div>
      ${s.resp && num ? `<div class="rapidas"><span class="muted sm">Responder:</span>${RESP.filter(([k]) => s.cli ? ['si', 'mensual', 'referir', 'testimonio'].includes(k) : !['referir', 'testimonio'].includes(k)).map(([k, l, f]) => `<a class="btn alt" href="https://wa.me/507${num}?text=${encodeURIComponent(f(n))}" target="_blank" rel="noopener">${l}</a>`).join('')}</div>` : ''}
      <details><summary>${sg ? 'Ver el mensaje de seguimiento' : 'Ver o cambiar el mensaje'}${fijos.length && !num ? ' · qué decir si llamas' : ''}</summary>
        <textarea ${sg ? 'readonly' : `data-msg="${n.id}"`}>${msg.replace(/</g, '&lt;')}</textarea>
        ${fijos.length ? `<p class="muted sm"><b>Si llamas:</b> ${guion(n)}</p>` : ''}
      </details>
    </div>`;
  }).join('') + (vis.length > limite ? `<button class="btn alt" style="width:100%;justify-content:center" data-mas>Mostrar más (${vis.length - limite} restantes)</button>` : '') || '<p class="muted">No hay negocios en este filtro.</p>';
}
document.addEventListener('click', (e) => {
  const f = e.target.closest('[data-f]'); if (f) { filtro = f.dataset.f; limite = 40; pintar(); return; }
  const z = e.target.closest('[data-z]'); if (z) { cfg.zona = z.dataset.z; limite = 40; guardar(); pintar(); return; }
  if (e.target.closest('[data-mas]')) { limite += 40; pintar(); return; }
  const ok = e.target.closest('[data-ok]'); if (ok) { const id = ok.dataset.ok; st[id] = { ...st[id], ok: !st[id]?.ok, fecha: hoy() }; guardar(); pintar(); return; }
  const cp = e.target.closest('[data-copiar]'); if (cp) { const r = RESP.find(([k]) => k === cp.dataset.copiar); navigator.clipboard.writeText(r[2](EJ)).then(() => { cp.textContent = '✓ Copiado'; setTimeout(() => { cp.textContent = 'Copiar'; }, 1500); }); return; }
  const cl = e.target.closest('[data-cli]'); if (cl) { const id = cl.dataset.cli; st[id] = { ...st[id], cli: !st[id]?.cli }; guardar(); pintar(); return; }
  const rp = e.target.closest('[data-resp]'); if (rp) { const id = rp.dataset.resp; st[id] = { ...st[id], resp: !st[id]?.resp }; guardar(); pintar(); return; }
  const env = e.target.closest('[data-env]'); if (env && env.getAttribute('aria-disabled') !== 'true') {
    if (enviadosHoy() >= TOPE || faltaEspera() > 0) { e.preventDefault(); cupo(); $('#cupo').scrollIntoView({ block: 'nearest' }); return; }
    cfg.ult = Date.now(); guardar(); const id = env.dataset.env, esSeg = env.hasAttribute('data-seg'); setTimeout(() => { st[id] = esSeg ? { ...st[id], seg: hoy() } : { ...st[id], ok: true, fecha: hoy() }; guardar(); pintar(); }, 800); }
});
document.addEventListener('change', (e) => {
  const n = e.target.dataset.num; if (n !== undefined) { const d = e.target.value.replace(/\D/g, '').replace(/^507/, ''); st[n] = { ...st[n], num: d.length === 8 && d[0] === '6' ? d : undefined }; if (d && !st[n].num) alert('Ese número no parece un celular de Panamá (8 dígitos que empiezan por 6).'); guardar(); pintar(); }
  const m = e.target.dataset.msg; if (m !== undefined) { st[m] = { ...st[m], msg: e.target.value }; guardar(); pintar(); }
});
pintarRapidas(); pintar(); cupo();
</script></body></html>'''
open(OUT, 'w', encoding='utf-8').write(HTML.replace('__DATA__', json.dumps(filas, ensure_ascii=False)))
print(len(filas), 'negocios')
