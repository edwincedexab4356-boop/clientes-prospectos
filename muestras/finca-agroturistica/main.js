(() => {
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const WHATSAPP = ''; // número de la finca, ej. '50760000000'

  /* Fundido lento de la foto principal */
  const fotos = [...document.querySelectorAll('.hero-fotos img')];
  let i = 0, reloj = null;
  function siguiente() {
    fotos[i].classList.remove('activa');
    i = (i + 1) % fotos.length;
    fotos[i].classList.add('activa');
  }
  function arrancar() { if (!reduce && !reloj) reloj = setInterval(siguiente, 6000); }
  function parar() { clearInterval(reloj); reloj = null; }
  document.addEventListener('visibilitychange', () => document.hidden ? parar() : arrancar());
  arrancar();

  /* Arma tu visita */
  const hoja = document.getElementById('hoja');
  const personas = document.getElementById('personas');
  const total = document.getElementById('total');
  const enviar = document.getElementById('enviar');
  const n = () => Math.min(60, Math.max(1, parseInt(personas.value, 10) || 1));
  function calcular() {
    const elegidos = [...hoja.querySelectorAll('input[type=checkbox]:checked')];
    const pp = elegidos.reduce((s, x) => s + Number(x.dataset.p), 0);
    const t = pp * n();
    total.textContent = `$${t}`;
    const lista = elegidos.map(x => `• ${x.value} ($${x.dataset.p})`).join('\n');
    const msg = elegidos.length
      ? `Hola, quiero reservar una visita a la finca para ${n()} persona${n() > 1 ? 's' : ''}:\n${lista}\nTotal aproximado: $${t}.\n¿Qué fecha tienen disponible?`
      : `Hola, quiero reservar una visita a la finca para ${n()} persona${n() > 1 ? 's' : ''}. ¿Qué fecha tienen disponible?`;
    enviar.href = `https://wa.me/${WHATSAPP}?text=${encodeURIComponent(msg)}`;
  }
  hoja.addEventListener('change', calcular);
  personas.addEventListener('input', calcular);
  personas.addEventListener('blur', () => { personas.value = n(); calcular(); });
  document.getElementById('menos').addEventListener('click', () => { personas.value = n() - 1 || 1; calcular(); });
  document.getElementById('mas').addEventListener('click', () => { personas.value = Math.min(60, n() + 1); calcular(); });
  hoja.addEventListener('submit', e => e.preventDefault());
  calcular();

  /* Galería: arrastrar con el ratón */
  const fila = document.querySelector('.fila-gal');
  let x0 = 0, s0 = 0, arrastrando = false;
  fila.addEventListener('pointerdown', e => { if (e.pointerType !== 'mouse') return; arrastrando = true; x0 = e.clientX; s0 = fila.scrollLeft; fila.style.cursor = 'grabbing'; });
  addEventListener('pointerup', () => { arrastrando = false; fila.style.cursor = ''; });
  fila.addEventListener('pointermove', e => { if (arrastrando) fila.scrollLeft = s0 - (e.clientX - x0); });
  fila.addEventListener('dragstart', e => e.preventDefault());

  /* Entradas escalonadas */
  document.documentElement.classList.add('js');
  const grupos = ['.intro > *', '.cult-cabeza h2', '.fila', '.hoja > *', '.panel-texto', '.gal-cabeza > *', '.fila-gal figure', '.preg-fotos img'];
  const piezas = [];
  grupos.forEach(sel => document.querySelectorAll(sel).forEach((el, k) => {
    el.classList.add('sube'); el.style.transitionDelay = `${(k % 6) * 75}ms`; piezas.push(el);
  }));
  if (reduce || !('IntersectionObserver' in window)) { piezas.forEach(el => el.classList.add('visto')); return; }
  const revisar = () => {
    for (let k = piezas.length - 1; k >= 0; k--) {
      if (piezas[k].getBoundingClientRect().top < innerHeight * 0.95) { piezas[k].classList.add('visto'); piezas.splice(k, 1); }
    }
    if (!piezas.length) removeEventListener('scroll', alMover);
  };
  let pend = false;
  const alMover = () => { if (!pend) { pend = true; requestAnimationFrame(() => { pend = false; revisar(); }); } };
  addEventListener('scroll', alMover, { passive: true });
  addEventListener('resize', alMover);
  revisar();
})();
