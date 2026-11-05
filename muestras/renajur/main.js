(() => {
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const CORREO = 'contacto@renajur.org';

  /* Tira de fotos dentro de la frase */
  const hero = document.querySelector('.hero');
  const carrete = document.querySelector('.carrete');
  const fotos = [...carrete.querySelectorAll('img')];
  const lugar = document.getElementById('lugar');
  const lugares = [
    'Arreo de ganado · Cambutal, Los Santos',
    'Vivero de café · Boquete, Chiriquí',
    'Isla de la comarca Guna Yala',
    'La Palma, provincia de Darién',
    'Semana del campesino · 2019'
  ];
  let actual = -1;
  function mover() {
    const r = hero.getBoundingClientRect();
    const total = hero.offsetHeight - innerHeight;
    const p = reduce || total <= 0 || innerWidth <= 640 ? 0 : Math.min(1, Math.max(0, -r.top / total));
    const h = fotos[0].getBoundingClientRect().height + 24;
    const paso = p * (fotos.length - 1);
    if (innerWidth <= 640) { carrete.style.transform = ''; return; }
    carrete.style.transform = `translateY(${-(fotos[0].getBoundingClientRect().height / 2) - paso * h}px)`;
    const i = Math.round(paso);
    if (i !== actual) { actual = i; lugar.textContent = lugares[i]; }
  }
  addEventListener('scroll', () => requestAnimationFrame(mover), { passive: true });
  addEventListener('resize', mover);
  fotos[0].complete ? mover() : fotos[0].addEventListener('load', mover);
  mover();

  /* Territorio elegido + áreas → correo ya escrito */
  let territorio = '';
  const areas = new Set();
  const miT = document.getElementById('mi-territorio');
  const bJoven = document.getElementById('b-joven');
  const bAliado = document.getElementById('b-aliado');
  function enlace(tipo) {
    const desde = territorio ? ` Soy de ${territorio}.` : '';
    const intereses = areas.size ? ` Me interesa: ${[...areas].join(', ').toLowerCase()}.` : '';
    const asunto = tipo === 'joven' ? 'Quiero unirme a RENAJUR' : 'Quiero ser aliado de RENAJUR';
    const cuerpo = tipo === 'joven'
      ? `Hola, RENAJUR.${desde}${intereses} Quiero unirme a la red. Mi nombre es: `
      : `Hola, RENAJUR. Represento a una institución u organización y queremos colaborar con la red.${desde}${intereses} Nombre y organización: `;
    return `mailto:${CORREO}?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
  }
  function actualizar() {
    miT.textContent = territorio || 'mi provincia o comarca';
    bJoven.href = enlace('joven');
    bAliado.href = enlace('aliado');
  }
  const botonesT = document.querySelectorAll('.nombres button');
  botonesT.forEach(b => {
    b.setAttribute('aria-pressed', 'false');
    b.addEventListener('click', () => {
      botonesT.forEach(o => o.setAttribute('aria-pressed', 'false'));
      b.setAttribute('aria-pressed', 'true');
      territorio = b.dataset.t;
      actualizar();
      document.getElementById('unete').scrollIntoView({ behavior: reduce ? 'auto' : 'smooth' });
    });
  });
  document.querySelectorAll('.chip').forEach(c => c.addEventListener('click', () => {
    const on = c.getAttribute('aria-pressed') !== 'true';
    c.setAttribute('aria-pressed', String(on));
    on ? areas.add(c.textContent) : areas.delete(c.textContent);
    actualizar();
  }));
  actualizar();

  /* Con quién: un nombre encendido a la vez */
  const nombres = [...document.querySelectorAll('#cq li')];
  let k = 0;
  if (!reduce) setInterval(() => {
    nombres.forEach(n => n.classList.remove('lit'));
    nombres[k % nombres.length].classList.add('lit');
    k++;
  }, 2000);

  /* Entradas por sección, escalonadas */
  document.documentElement.classList.add('js');
  const grupos = [
    '.quienes .grande, .quienes .datos',
    '.cifras-titulo, .fila',
    '.areas-cabeza h2', '.area',
    '.nombres-cabeza, .nombres li, .nombres-pie',
    '.estr-texto h2, .estr-texto p, .niveles li',
    '.cq-cabeza, .frase-nombres li',
    '.contacto > *'
  ];
  const piezas = [];
  grupos.forEach(sel => document.querySelectorAll(sel).forEach((el, i) => {
    el.classList.add('sube');
    el.style.transitionDelay = `${(i % 6) * 75}ms`;
    piezas.push(el);
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
