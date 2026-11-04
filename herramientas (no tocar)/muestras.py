import json, os, sys, urllib.parse, html

OUT = sys.argv[1]  # carpeta muestras
XLSX = sys.argv[3]
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from muestras_resto import resto, PORTAFOLIO
YO = '50766952340'
OFERTA = 'Oferta de lanzamiento esta semana: página $150 · catálogo $350 · NovaCore primer mes gratis'

TALLER_ITEMS = lambda extra=(): [
    {'cat': 'Mecánica', 'ic': '🧰', 'n': 'Diagnóstico computarizado', 'd': 'Escáner y revisión de luces del tablero.', 'p': 25},
    {'cat': 'Mecánica', 'ic': '🛢️', 'n': 'Cambio de aceite y filtro', 'd': 'Incluye revisión de niveles y presión de llantas.', 'p': 35, 'desde': 1},
    {'cat': 'Mecánica', 'ic': '🛑', 'n': 'Frenos (pastillas delanteras)', 'd': 'Cambio de pastillas y limpieza de discos.', 'p': 45, 'desde': 1},
    {'cat': 'Mecánica', 'ic': '🔧', 'n': 'Suspensión y tren delantero', 'd': 'Revisión de rótulas, terminales y amortiguadores.', 'p': 60, 'desde': 1},
    {'cat': 'Mecánica', 'ic': '❄️', 'n': 'Aire acondicionado', 'd': 'Revisión de fugas y recarga de gas.', 'p': 45, 'desde': 1},
    {'cat': 'Chapistería', 'ic': '🚗', 'n': 'Enderezado de golpes', 'd': 'Puertas, bómperes y guardafangos.', 'p': 80, 'desde': 1},
    {'cat': 'Chapistería', 'ic': '🎨', 'n': 'Pintura por pieza', 'd': 'Igualación de color y acabado brillante.', 'p': 90, 'desde': 1, 'u': 'por pieza'},
    {'cat': 'Chapistería', 'ic': '✨', 'n': 'Pulido y brillado', 'd': 'Quita rayones finos y devuelve el brillo.', 'p': 50, 'desde': 1},
    {'cat': 'Seguros', 'ic': '📄', 'n': 'Reparación por aseguradora', 'd': 'Le ayudamos con el avalúo y el trámite con su seguro.', 'p': 0, 'u': 'según avalúo'},
    *extra,
]
FERRE_ITEMS = [
    {'cat': 'Construcción', 'ic': '🧱', 'n': 'Cemento gris 42.5 kg', 'd': 'Para pegar bloque, repellos y vaciados.', 'p': 9.50},
    {'cat': 'Construcción', 'ic': '🧱', 'n': 'Bloque de 4"', 'd': 'Bloque de cemento estándar.', 'p': 0.85, 'u': 'c/u'},
    {'cat': 'Construcción', 'ic': '🪨', 'n': 'Arena y piedra (carretilla)', 'd': 'Entrega en Colón y Sabanitas.', 'p': 6, 'u': 'por carretilla'},
    {'cat': 'Construcción', 'ic': '🔩', 'n': 'Varilla corrugada 3/8"', 'd': 'Largo de 30 pies.', 'p': 4.50},
    {'cat': 'Plomería', 'ic': '🚰', 'n': 'Tubo PVC 1/2" (20 pies)', 'd': 'Para agua potable.', 'p': 3.25},
    {'cat': 'Plomería', 'ic': '🚿', 'n': 'Llave de chorro', 'd': 'Metálica, de 1/2".', 'p': 8},
    {'cat': 'Electricidad', 'ic': '🔌', 'n': 'Cable THHN #12 (rollo)', 'd': '100 metros.', 'p': 65},
    {'cat': 'Electricidad', 'ic': '💡', 'n': 'Bombillo LED 9W', 'd': 'Luz blanca, rosca normal.', 'p': 1.75},
    {'cat': 'Pintura', 'ic': '🪣', 'n': 'Pintura de agua (galón)', 'd': 'Interior y exterior, varios colores.', 'p': 22},
    {'cat': 'Pintura', 'ic': '🖌️', 'n': 'Brocha de 3"', 'd': 'Cerda natural.', 'p': 2.50},
    {'cat': 'Herramientas', 'ic': '🔨', 'n': 'Martillo de uña', 'd': 'Mango de fibra.', 'p': 9},
    {'cat': 'Herramientas', 'ic': '📏', 'n': 'Cinta métrica 5 m', 'd': 'Con freno.', 'p': 6},
]
EXTRA_TALLER = {
    'titulo': 'Y para llevar su taller: NovaCore',
    'texto': 'Un sistema sencillo, en el celular o la computadora, para que sepa cuánto gana de verdad cada mes.',
    'puntos': ['Cada trabajo con su cliente, el precio y lo que ya le abonaron', 'Lo que le deben, sin cuadernos', 'Gastos en repuestos y ayudantes', 'Reporte del mes: cuánto entró, cuánto salió y cuánto le quedó'],
    'btn': 'Quiero ver NovaCore', 'link': None,
}
EXTRA_FERRE = {
    'titulo': 'Y para su caja e inventario: un sistema a medida',
    'texto': 'Además del catálogo, le hago un sistema para el mostrador. Puede probar una demostración de una ferretería ahora mismo.',
    'puntos': ['Caja rápida con ticket', 'Inventario que avisa cuando algo se acaba', 'Fiado de clientes y abonos', 'Ventas por día y por forma de pago'],
    'btn': 'Probar la demostración', 'link': 'https://edwin2030-cmd.github.io/dek-portafolio/demo/',
}

N = [
    dict(slug='taller-auto-fino', nombre='Taller Auto Fino', ini='AF', color='#c2410c', c2='#ffedd5', rubro='Taller mecánico y chapistería',
         zona='Barrio Sur, Av. Central, Calle 13 y 14, Colón', tel='441-3122', telLink='+5074413122',
         eyebrow='Taller en Colón', h1='Su carro en buenas manos, aquí en Colón', p='Mecánica, enderezado y pintura. Pida su cita por WhatsApp y le decimos el precio antes de empezar.',
         modo='cita', items=TALLER_ITEMS(), extraLbl='Carro (marca, modelo y año)', extraEj='Toyota Hilux 2018', extra=EXTRA_TALLER),
    dict(slug='taller-don-bosco', nombre='Taller Don Bosco', ini='DB', color='#1d4ed8', c2='#dbeafe', rubro='Taller mecánico',
         zona='Colón', tel='447-0070', telLink='+5074470070',
         eyebrow='Taller en Colón', h1='Mecánica honesta y a tiempo', p='Revisamos su carro, le explicamos qué tiene y cuánto cuesta. Reserve su cita en un minuto.',
         modo='cita', items=TALLER_ITEMS(), extraLbl='Carro (marca, modelo y año)', extraEj='Hyundai Accent 2016', extra=EXTRA_TALLER),
    dict(slug='mc-place', nombre='MC Place', ini='MC', color='#0f766e', c2='#ccfbf1', rubro='Taller automotriz',
         zona='Colón', tel='6022-1313', telLink='+50760221313',
         eyebrow='Taller automotriz · Colón', h1='Reparamos su carro y lo dejamos como nuevo', p='Mecánica general, chapistería y reparaciones por seguro. Escríbanos y agende su cita.',
         modo='cita', items=TALLER_ITEMS(), extraLbl='Carro (marca, modelo y año)', extraEj='Nissan Frontier 2020', extra=EXTRA_TALLER),
    dict(slug='quality-paint-shop', nombre='Quality Paint Shop', ini='QP', color='#7c3aed', c2='#ede9fe', rubro='Chapistería y pintura',
         zona='Av. Meléndez, Colón', tel='6614-6792', telLink='+50766146792',
         eyebrow='Chapistería y pintura · Colón', h1='Pintura y enderezado con acabado de agencia', p='Golpes, rayones y pintura completa. Mande fotos por WhatsApp y le damos un precio aproximado.',
         modo='cita', items=[i for i in TALLER_ITEMS() if i['cat'] in ('Chapistería', 'Seguros')] + [
             {'cat': 'Chapistería', 'ic': '🚙', 'n': 'Pintura completa', 'd': 'Carro completo, color original o cambio de color.', 'p': 650, 'desde': 1},
             {'cat': 'Chapistería', 'ic': '🧽', 'n': 'Restauración de focos', 'd': 'Quita lo amarillo y opaco.', 'p': 25, 'u': 'el par'},
             {'cat': 'Chapistería', 'ic': '🪛', 'n': 'Reparación de bómper', 'd': 'Soldadura de plástico y pintura.', 'p': 70, 'desde': 1}],
         extraLbl='Carro y qué le pasó', extraEj='Corolla 2017, golpe en la puerta', extra=EXTRA_TALLER),
    dict(slug='ferreteria-colon', nombre='Ferretería Colón', ini='FC', color='#b45309', c2='#fef3c7', rubro='Ferretería y materiales',
         zona='Colón', tel='445-1338', telLink='+5074451338',
         eyebrow='Ferretería en Colón', h1='Todo para su obra, pedido por WhatsApp', p='Materiales de construcción, plomería, electricidad y pintura. Arme su pedido aquí y se lo tenemos listo.',
         modo='pedido', items=FERRE_ITEMS, extraLbl='¿Lo retira o se lo llevamos? (dirección)', extraEj='Lo retiro en la tienda', extra=EXTRA_FERRE),
    dict(slug='ferreteria-anthony', nombre='Ferretería Anthony', ini='FA', color='#15803d', c2='#dcfce7', rubro='Ferretería',
         zona='Av. Herrera, Colón', tel='441-4643', telLink='+5074414643',
         eyebrow='Ferretería · Av. Herrera', h1='Su ferretería de confianza, ahora en el celular', p='Vea precios, arme su lista y mándela por WhatsApp. Se la preparamos para que solo pase a recogerla.',
         modo='pedido', items=FERRE_ITEMS, extraLbl='¿Lo retira o se lo llevamos? (dirección)', extraEj='Entrega en Cativa', extra=EXTRA_FERRE),
    dict(slug='punto-cuatro', nombre='Punto Cuatro', ini='P4', color='#db2777', c2='#fce7f3', rubro='Artículos de fiesta',
         zona='Calle 13, Colón', tel='433-2582', telLink='+5074332582',
         eyebrow='Todo para su fiesta · Colón', h1='Arme su fiesta desde el celular', p='Globos, piñatas, desechables y decoración. Elija lo que necesita y se lo tenemos listo.',
         modo='pedido', items=[
             {'cat': 'Globos', 'ic': '🎈', 'n': 'Globos de colores (50)', 'd': 'Látex, 12 pulgadas.', 'p': 6},
             {'cat': 'Globos', 'ic': '🔢', 'n': 'Globo número metalizado', 'd': '32 pulgadas, dorado o plateado.', 'p': 4},
             {'cat': 'Globos', 'ic': '🌈', 'n': 'Arco de globos', 'd': 'Armado e instalado en su evento.', 'p': 60, 'desde': 1},
             {'cat': 'Piñatas', 'ic': '🪅', 'n': 'Piñata de personaje', 'd': 'Varios personajes, por encargo.', 'p': 25, 'desde': 1},
             {'cat': 'Piñatas', 'ic': '🍬', 'n': 'Relleno de piñata', 'd': 'Caramelos y juguetes surtidos.', 'p': 12},
             {'cat': 'Desechables', 'ic': '🍽️', 'n': 'Platos de fiesta (25)', 'd': 'Varios diseños.', 'p': 3.50},
             {'cat': 'Desechables', 'ic': '🥤', 'n': 'Vasos de fiesta (25)', 'd': 'A juego con los platos.', 'p': 3},
             {'cat': 'Decoración', 'ic': '🎉', 'n': 'Kit de cumpleaños', 'd': 'Banderín, cortina metálica y centro de mesa.', 'p': 18},
             {'cat': 'Decoración', 'ic': '🕯️', 'n': 'Velas de cumpleaños', 'd': 'Número o paquete de 24.', 'p': 1.50},
         ], extraLbl='Fecha de la fiesta', extraEj='Sábado 15, cumpleaños infantil', extra={
             'titulo': 'Y para sus decoraciones: NovaCore', 'texto': 'Si arma eventos, lleve cada uno en orden desde el celular.',
             'puntos': ['Cada evento con su fecha, cliente y precio', 'El adelanto y lo que falta por cobrar', 'Lo que gastó en materiales', 'Cuánto ganó en cada evento y en el mes'],
             'btn': 'Quiero ver NovaCore', 'link': None}),
    dict(slug='panaderia-colon', nombre='Panadería Colón', ini='PC', color='#a16207', c2='#fef9c3', rubro='Panadería y repostería',
         zona='Av. Herrera, Colón', tel='441-3941', telLink='+5074413941',
         eyebrow='Tradición colonense', h1='Los pastelitos de siempre, ahora por encargo', p='Pastelitos de carne, plátano y bacalao, pan del día y dulces. Haga su pedido por WhatsApp y páselo a buscar.',
         modo='pedido', items=[
             {'cat': 'Pastelitos', 'ic': '🥟', 'n': 'Pastelito de carne', 'd': 'El clásico de la casa.', 'p': 1.00},
             {'cat': 'Pastelitos', 'ic': '🥟', 'n': 'Pastelito de plátano', 'd': 'Dulce y suave.', 'p': 1.00},
             {'cat': 'Pastelitos', 'ic': '🥟', 'n': 'Pastelito de bacalao', 'd': 'Receta tradicional.', 'p': 1.25},
             {'cat': 'Pastelitos', 'ic': '📦', 'n': 'Docena surtida', 'd': 'Para la oficina o la familia.', 'p': 11},
             {'cat': 'Pan', 'ic': '🍞', 'n': 'Pan de molde', 'd': 'Del día.', 'p': 2.75},
             {'cat': 'Pan', 'ic': '🥖', 'n': 'Pan de flauta', 'd': 'Crujiente por fuera.', 'p': 0.50},
             {'cat': 'Dulces', 'ic': '🍰', 'n': 'Dulce del día (porción)', 'd': 'Pregunte por el de hoy.', 'p': 1.50},
             {'cat': 'Por encargo', 'ic': '🎂', 'n': 'Bizcocho de cumpleaños', 'd': 'Decorado, por encargo con 2 días.', 'p': 25, 'desde': 1},
         ], extraLbl='¿A qué hora lo pasa a buscar?', extraEj='Mañana a las 8:00 a. m.', extra={
             'titulo': 'Y para el mostrador: un sistema a medida', 'texto': 'Lleve los encargos, las ventas del día y los gastos sin cuaderno.',
             'puntos': ['Encargos con fecha y hora de entrega', 'Ventas del día por producto', 'Gastos de harina, gas y empleados', 'Cuánto le quedó al final del mes'],
             'btn': 'Quiero saber más', 'link': None}),
]

for n in N: n.setdefault('clave', n['nombre'])
N += resto(XLSX, N, TALLER_ITEMS(), FERRE_ITEMS, EXTRA_TALLER)

TPL = '''<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="robots" content="noindex, nofollow"><meta name="theme-color" content="{color}">
<title>{nombre} · Muestra de D.E.K NovaCore</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="../muestra.css">
<style>:root{{--c:{color};--c2:{c2}}}</style>
</head><body>
<div class="propuesta"><div class="wrap"><span>👋 <b>Muestra hecha por D.E.K NovaCore para {nombre_e}.</b> <span class="oferta">{oferta}</span></span>
<a class="btn-p" href="{wa}" target="_blank" rel="noopener">Me interesa</a></div></div>
<header class="top"><div class="wrap"><a class="marca" href="#"><i>{ini}</i><span>{nombre_e}<small>{rubro}</small></span></a>
{btn_tel}</div></header>
<main>
<section class="hero"><div class="wrap"><span class="eyebrow">{eyebrow}</span><h1>{h1}</h1><p>{p}</p>
<div class="acciones"><a class="btn c" href="#catalogo">{cta}</a><a class="btn l" href="{mapa}" target="_blank" rel="noopener">Cómo llegar</a></div>
<div class="datos"><span>{zona}</span><span>Tel. {tel_txt}</span><span>Lunes a sábado</span></div></div></section>
<section class="sec" id="catalogo"><div class="wrap"><h2>{titcat}</h2><p class="sub">{subcat}</p>
<div class="filtros" id="filtros"></div><div class="grid" id="grid"></div></div></section>
<section class="sec"><div class="wrap"><div class="info">
<div><h3>Dirección</h3><p>{zona}</p></div><div><h3>Teléfono</h3><p>{tel_txt}</p></div><div><h3>Horario</h3><p>Lunes a sábado · 8:00 a. m. a 5:00 p. m.</p></div></div></div></section>
<section class="sec"><div class="wrap"><div class="extra"><h2>{x_tit}</h2><p>{x_txt}</p><ul>{x_li}</ul>
<a class="btn c" href="{x_link}" {x_target}>{x_btn}</a></div></div></section>
</main>
<footer>Muestra de propuesta. Precios, horario y textos son de ejemplo y se cambian por los suyos.<br>Hecha por <a href="https://edwin2030-cmd.github.io/dek-portafolio/">D.E.K NovaCore</a> · WhatsApp 6695-2340</footer>
<div class="carrito" id="carrito"><span id="c-txt"></span><button type="button" class="btn" id="ver-pedido">{btnpedido}</button></div>
<div class="modal" id="modal"><div class="hoja" id="hoja" role="dialog" aria-modal="true"></div></div>
<script>window.MUESTRA = {data};</script>
<script src="../muestra.js"></script>
</body></html>
'''
e = html.escape
for n in N:
    wa = 'https://wa.me/' + YO + '?text=' + urllib.parse.quote(f'Hola, vi la muestra que me hicieron para {n["nombre"]} y me interesa. ¿Cómo seguimos?')
    x = n['extra']
    x_link = x['link'] or 'https://wa.me/' + YO + '?text=' + urllib.parse.quote(f'Hola, soy de {n["nombre"]}. Quiero saber más de: {x["titulo"].replace("Y para ", "").replace("Y para sus ", "")}')
    data = {'nombre': n['nombre'], 'tel': n['tel'], 'modo': n['modo'], 'items': n['items'], 'extraLbl': n['extraLbl'], 'extraEj': n['extraEj'], 'waPropuesta': wa}
    pag = TPL.format(color=n['color'], c2=n['c2'], nombre=e(n['nombre']), nombre_e=e(n['nombre']), oferta=e(OFERTA), wa=e(wa), ini=e(n['ini']), rubro=e(n['rubro']),
                     btn_tel=(f'<a class="btn c" href="tel:{n["telLink"]}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6.6 10.8a15.1 15.1 0 006.6 6.6l2.2-2.2c.3-.3.7-.4 1-.2 1.1.4 2.3.6 3.6.6.6 0 1 .4 1 1V20c0 .6-.4 1-1 1A17 17 0 013 4c0-.6.4-1 1-1h3.5c.6 0 1 .4 1 1 0 1.3.2 2.5.6 3.6.1.3 0 .7-.2 1L6.6 10.8z"/></svg><span>{e(n["tel"])}</span></a>' if n['tel'] else '<a class="btn c" href="#catalogo">' + ('Pedir cita' if n['modo'] == 'cita' else 'Hacer pedido') + '</a>'),
                     tel_txt=e(n['tel']) if n['tel'] else 'su número aquí', eyebrow=e(n['eyebrow']), h1=e(n['h1']), p=e(n['p']),
                     cta='Pedir una cita' if n['modo'] == 'cita' else 'Ver productos y precios',
                     mapa=e('https://www.google.com/maps/search/?api=1&query=' + urllib.parse.quote(n['nombre'] + ' ' + n['zona'] + ' Panamá')),
                     zona=e(n['zona']), titcat='Servicios y precios' if n['modo'] == 'cita' else 'Productos y precios',
                     subcat='Elija lo que necesita y pida su cita por WhatsApp.' if n['modo'] == 'cita' else 'Agregue lo que necesita y mande su pedido por WhatsApp.',
                     x_tit=e(x['titulo']), x_txt=e(x['texto']), x_li=''.join(f'<li>{e(i)}</li>' for i in x['puntos']),
                     x_link=e(x_link), x_target='' if x['link'] else 'target="_blank" rel="noopener"', x_btn=e(x['btn']),
                     btnpedido='Ver solicitud' if n['modo'] == 'cita' else 'Ver pedido',
                     data=json.dumps(data, ensure_ascii=False).replace('</', '<\\/'))
    d = os.path.join(OUT, n['slug']); os.makedirs(d, exist_ok=True)
    open(os.path.join(d, 'index.html'), 'w', encoding='utf-8', newline='\n').write(pag)
json.dump({n['clave']: n['slug'] for n in N}, open(sys.argv[2], 'w', encoding='utf-8'), ensure_ascii=False)
print('muestras:', len(N))
