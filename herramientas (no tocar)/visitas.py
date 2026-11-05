# Genera "ruta-visitas.html": negocios de Colón agrupados por sector, para visitarlos en persona.
# Uso: python visitas.py "Prospectos Colón.xlsx" ruta-visitas.html slugs.json
import json, re, sys
from openpyxl import load_workbook

XLSX, OUT = sys.argv[1], sys.argv[2]
SLUGS = json.load(open(sys.argv[3], encoding='utf-8'))
ws = load_workbook(XLSX)['Prospectos']
filas = []
for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    prio, nombre, rubro, zona, tel = r[0], r[1], r[2], str(r[3] or 'Colón'), str(r[4] or '')
    if re.search(r'Arraij|Chorrera|San Miguelito|(^|, )Panamá$', zona): continue
    link = ws.cell(row=i, column=11).hyperlink
    slug = SLUGS.get(nombre) or next((v for k, v in SLUGS.items() if k and nombre.startswith(k)), None)
    partes = [p.strip() for p in zona.split(',')]
    sector = partes[0] if len(partes) > 1 else ('Sabanitas / Cativá' if 'Sabanitas' in zona or 'Cativa' in zona else 'Colón (sin sector)')
    if sector.lower() in ('colón', 'colon'): sector = 'Colón (sin sector)'
    filas.append({'id': slug or f'v{i}', 'prio': prio or 9, 'nombre': nombre, 'rubro': str(rubro or '').lower(), 'zona': zona,
                  'sector': sector, 'tel': tel, 'mapa': link.target if link else '', 'muestra': slug})

HTML = r'''<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ruta de visitas en Colón - D.E.K NovaCore</title>
<style>
:root{--bg:#0a0d14;--panel:#121826;--line:#243047;--text:#e8ecf4;--muted:#8a94a8;--acc:#5a9bfc;--ok:#34d399;--warn:#fbbf24}
*{box-sizing:border-box}body{margin:0;font:15px/1.5 system-ui,Segoe UI,Arial;background:var(--bg);color:var(--text)}
header{padding:16px 20px;border-bottom:1px solid var(--line)}h1{margin:0 0 4px;font-size:20px}h2{font-size:16px;margin:22px 0 8px}
.muted{color:var(--muted)}.sm{font-size:13px}main{padding:16px 20px 60px;max-width:1000px;margin:auto}
.guion{background:#0f1a2e;border:1px solid #1f3b6b;border-radius:12px;padding:12px 16px;font-size:14px}
.guion summary{cursor:pointer;font-weight:600}.guion p{margin:8px 0}
.bar{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}
.chip{border:1px solid var(--line);background:none;color:var(--text);border-radius:99px;padding:6px 12px;cursor:pointer}.chip.on{background:var(--acc);color:#05080f;border-color:var(--acc)}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:10px 14px;margin-bottom:8px;display:flex;gap:10px;justify-content:space-between;flex-wrap:wrap;align-items:center}
.card.visto{opacity:.5}.card.int{border-color:var(--ok)}
.tag{display:inline-block;font-size:12px;border-radius:99px;padding:2px 9px;background:#1b2436;color:var(--muted);margin-right:6px}
.btn{border:0;border-radius:10px;padding:7px 11px;font-weight:600;cursor:pointer;text-decoration:none;font-size:13px;background:#1b2436;color:var(--text);display:inline-block}
.btn.v{background:#123524;color:var(--ok)}.btn.i{background:#3a3112;color:var(--warn)}
.acts{display:flex;gap:6px;flex-wrap:wrap}
@media print{body{background:#fff;color:#000}.card{border-color:#999;background:#fff}.acts,.bar,.guion summary{display:none}.guion{display:block}}
</style></head><body>
<header><h1>🚶 Ruta de visitas en Colón</h1>
<div class="muted sm">Negocios de Colón, Sabanitas y Cativá agrupados por sector para que visites varios en una salida. Lleva el celular cargado y abre <b>Ver su muestra</b> antes de entrar. Lo que marques se guarda en este navegador.</div></header>
<main>
<details class="guion" open><summary>Qué decir (30 segundos)</summary>
<p>"Buenas, soy <b>[tu nombre]</b>, de aquí de Colón. Hago páginas y sistemas para negocios. Le preparé esta muestra con el nombre de su negocio, mírela." <i>(le das el celular)</i></p>
<p>"Sus clientes ven lo que vende y el pedido le llega a su WhatsApp ya escrito. Esta semana la tengo a <b>$150 en vez de $200</b>. Con la mitad empiezo hoy y en 5 días la tiene funcionando con sus fotos y precios. La otra mitad, cuando la vea lista."</p>
<p><b>Si dice "lo voy a pensar":</b> "Perfecto, ¿me da su WhatsApp y le paso la muestra para que la vea con calma?" (Eso ya es un avance: márcalo como <b>Interesado</b>.)</p>
<p><b>Si dice "está caro":</b> "Se puede en dos partes. Y con un solo cliente nuevo que lo encuentre por ahí, ya se paga. Si prefiere, empezamos solo con la página."</p>
<p><b>Si no está el dueño:</b> "¿A qué hora lo encuentro?" y anótalo en su tarjeta.</p>
</details>
<div class="bar" id="filtros"></div>
<div id="lista"></div>
</main>
<script>
const NEG = __DATA__;
const MUESTRAS = 'https://edwincedexab4356-boop.github.io/clientes-prospectos/muestras/';
const $ = (s) => document.querySelector(s);
const st = JSON.parse(localStorage.getItem('dek-visitas') || '{}');
const guardar = () => localStorage.setItem('dek-visitas', JSON.stringify(st));
let filtro = 'pendientes';
const hoy = () => new Date().toLocaleDateString('es-PA');
function pintar() {
  const c = { pendientes: NEG.filter((n) => !st[n.id]?.v).length, interesados: NEG.filter((n) => st[n.id]?.i).length, visitados: NEG.filter((n) => st[n.id]?.v).length, todos: NEG.length };
  $('#filtros').innerHTML = [['pendientes', 'Por visitar'], ['interesados', '⭐ Interesados'], ['visitados', 'Visitados'], ['todos', 'Todos']]
    .map(([k, l]) => `<button class="chip ${filtro === k ? 'on' : ''}" data-f="${k}">${l} (${c[k]})</button>`).join('');
  const vis = NEG.filter((n) => filtro === 'todos' || (filtro === 'interesados' ? st[n.id]?.i : filtro === 'visitados' ? st[n.id]?.v : !st[n.id]?.v));
  const grupos = {}; vis.forEach((n) => (grupos[n.sector] = grupos[n.sector] || []).push(n));
  const orden = Object.keys(grupos).sort((a, b) => grupos[b].length - grupos[a].length);
  $('#lista').innerHTML = orden.map((g) => `<h2>${g} <span class="muted sm">(${grupos[g].length})</span></h2>` + grupos[g].sort((a, b) => a.prio - b.prio).map((n) => {
    const s = st[n.id] || {};
    const tel = n.tel ? n.tel.split(/[\/,]/).map((t) => t.trim()).filter(Boolean).map((t) => `<a class="btn" href="tel:+507${t.replace(/\D/g, '')}">📞 ${t}</a>`).join('') : '';
    return `<div class="card ${s.i ? 'int' : s.v ? 'visto' : ''}"><div><b>${n.nombre}</b> <span class="tag">${n.rubro}</span>${s.v ? `<span class="tag">visitado ${s.f || ''}</span>` : ''}<div class="muted sm">${n.zona}</div>
      <input data-nota="${n.id}" value="${(s.nota || '').replace(/"/g, '&quot;')}" placeholder="Nota (ej: volver a las 3, dueño Juan)" style="margin-top:6px;width:100%;font:inherit;font-size:13px;color:inherit;background:#0d1320;border:1px solid var(--line);border-radius:8px;padding:5px 8px"></div>
      <div class="acts">${n.muestra ? `<a class="btn" href="${MUESTRAS}${n.muestra}/" target="_blank" rel="noopener">👁 Ver su muestra</a>` : ''}${n.mapa ? `<a class="btn" href="${n.mapa}" target="_blank" rel="noopener">📍 Mapa</a>` : ''}${tel}
      <button class="btn v" data-v="${n.id}">${s.v ? '↩ No visitado' : '✓ Visitado'}</button><button class="btn i" data-i="${n.id}">${s.i ? '↩ Quitar interés' : '⭐ Interesado'}</button></div></div>`;
  }).join('')).join('') || '<p class="muted">No hay negocios en este filtro.</p>';
}
document.addEventListener('click', (e) => {
  const f = e.target.closest('[data-f]'); if (f) { filtro = f.dataset.f; pintar(); return; }
  const v = e.target.closest('[data-v]'); if (v) { const id = v.dataset.v; st[id] = { ...st[id], v: !st[id]?.v, f: hoy() }; guardar(); pintar(); return; }
  const i = e.target.closest('[data-i]'); if (i) { const id = i.dataset.i; st[id] = { ...st[id], i: !st[id]?.i, v: true, f: st[id]?.f || hoy() }; guardar(); pintar(); }
});
document.addEventListener('change', (e) => { const id = e.target.dataset.nota; if (id !== undefined) { st[id] = { ...st[id], nota: e.target.value }; guardar(); } });
pintar();
</script></body></html>'''
open(OUT, 'w', encoding='utf-8').write(HTML.replace('__DATA__', json.dumps(filas, ensure_ascii=False)))
from collections import Counter
print(len(filas), 'negocios para visitar;', Counter(f['sector'] for f in filas).most_common(8))
