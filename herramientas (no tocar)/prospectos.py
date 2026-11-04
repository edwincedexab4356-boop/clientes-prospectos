import json, urllib.parse, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

BASE = sys.argv[1]
OUT = sys.argv[2]

OFERTA = {
    'taller':      ('NovaCore (rubro Taller) + página sencilla', 'NovaCore $20/mes · página $200', 'taller'),
    'ferreteria':  ('Sistema de inventario y caja (como la demo de la ferretería)', 'Sistema desde $900 o caja/reportes $350', 'ferreteria'),
    'repuestos':   ('Catálogo de repuestos con pedidos por WhatsApp', 'Catálogo con pedidos $450', 'catalogo'),
    'panaderia':   ('Página con menú y pedidos por WhatsApp', 'Catálogo con pedidos $450 o presentación $200', 'comida'),
    'restaurante': ('Página con menú y pedidos por WhatsApp', 'Presentación $200 o catálogo con pedidos $450', 'comida'),
    'cafe':        ('Página con menú y pedidos por WhatsApp', 'Presentación $200', 'comida'),
    'belleza':     ('Página con servicios, precios y citas por WhatsApp', 'Presentación $200', 'belleza'),
    'peluqueria':  ('Página con servicios, precios y citas por WhatsApp', 'Presentación $200', 'belleza'),
    'dentista':    ('Página con servicios y citas por WhatsApp', 'Web $380', 'salud'),
    'clinica':     ('Página con servicios y citas por WhatsApp', 'Web $380', 'salud'),
    'lavanderia':  ('NovaCore para órdenes y cobros + página sencilla', 'NovaCore $20/mes · página $200', 'servicio'),
    'eventos':     ('NovaCore (rubro Eventos) + página con fotos', 'NovaCore $20/mes · web $380', 'eventos'),
    'pinturas':    ('Catálogo con pedidos por WhatsApp', 'Catálogo con pedidos $450', 'catalogo'),
    'lavado':      ('Página con precios y reservas por WhatsApp', 'Presentación $200', 'servicio'),
    'ropa':        ('Tienda en línea / catálogo por WhatsApp', 'Catálogo $450 o tienda $700', 'catalogo'),
    'otro':        ('Página de presentación', 'Presentación $200', 'servicio'),
}

def tipo_tel(t):
    d = ''.join(c for c in t if c.isdigit())
    if d.startswith('507'): d = d[3:]
    if not d: return ''
    return 'Celular (WhatsApp)' if d.startswith('6') and len(d) == 8 else 'Fijo (llamar o visitar)'

def maps(nombre, zona='Colón'):
    return 'https://www.google.com/maps/search/?api=1&query=' + urllib.parse.quote(f'{nombre} {zona} Panamá')

# Negocios con teléfono encontrado en directorios / OpenStreetMap
CON_TEL = [
    # (nombre, rubro, zona/dirección, teléfono, fuente)
    ('Taller Auto Fino', 'taller', 'Barrio Sur, Av. Central, Calle 13 y 14, Colón', '441-3122 / 441-3118', 'Lista de talleres autorizados ASSA/AFTA'),
    ('Taller Don Bosco', 'taller', 'Colón', '447-0070', 'Lista de talleres autorizados ASSA/AFTA'),
    ('MC Place, S.A. (taller)', 'taller', 'Colón', '6022-1313 / 6077-2498 / 441-3544', 'Lista de talleres autorizados ASSA'),
    ('Quality Paint Shop (pintura de autos)', 'taller', 'Av. Meléndez, Colón', '6614-6792 / 445-3627', 'Lista de talleres autorizados AFTA'),
    ('Ferretería Colón', 'ferreteria', 'Colón', '445-1338', 'Páginas Amarillas Panamá'),
    ('Ferretería Anthony', 'ferreteria', 'Av. Herrera, Colón', '441-4643', 'Directorio near-place / OpenStreetMap'),
    ('Panadería Colón', 'panaderia', 'Av. Herrera, Colón', '441-3941', 'OpenStreetMap'),
    ('Panadería Ketty', 'panaderia', 'Calle 10, Colón', '445-7791', 'OpenStreetMap'),
    ('Punto Cuatro (artículos de fiesta)', 'eventos', 'Calle 13, Colón', '433-2582', 'OpenStreetMap'),
    ('Rosania Clinic Dental', 'dentista', 'Av. del Frente, Colón', '441-7663', 'OpenStreetMap'),
    ('Auto Master (repuestos)', 'repuestos', 'Av. Central, Colón', '433-0860', 'OpenStreetMap'),
    ('Restaurante Grand Café', 'restaurante', 'Calle 1, Paseo Washington, Colón', '433-2092', 'coloncity.com'),
    ('Dayma (ropa)', 'ropa', 'Calle 10, Colón', '442-2232', 'OpenStreetMap'),
    ('Jinxu Shop', 'otro', 'Carretera Transístmica, Colón', '6888-6638', 'OpenStreetMap'),
]
# Negocios con Instagram / redes
CON_RED = [
    ('Platinum Barber Shop', 'peluqueria', 'Colón', 'instagram.com/platinum_barbershopp', 'Instagram'),
    ('IAM Fashion (salón, barbería y boutique)', 'belleza', 'Plaza 4 Altos, Colón', 'Buscar en TikTok/Instagram', 'TikTok'),
    ('Ariana Beauty Lounge', 'belleza', 'Buena Vista, Colón', 'Buscar en TikTok/Instagram', 'TikTok'),
]

PRIORIDAD = {'taller': 1, 'ferreteria': 1, 'eventos': 1, 'lavanderia': 2, 'panaderia': 2, 'restaurante': 2, 'belleza': 2,
             'peluqueria': 2, 'dentista': 2, 'clinica': 2, 'repuestos': 3, 'cafe': 3, 'pinturas': 3, 'lavado': 3, 'ropa': 3, 'otro': 3}

filas = []
vistos = set()
for n, r, z, t, f in CON_TEL:
    filas.append([PRIORIDAD[r], n, r, z, t, tipo_tel(t), 'No encontrada', f]); vistos.add(n.lower().split(' (')[0])
for n, r, z, t, f in CON_RED:
    filas.append([PRIORIDAD[r], n, r, z, t, 'Mensaje por Instagram/TikTok', 'No encontrada', f]); vistos.add(n.lower().split(' (')[0])

# Resto de OpenStreetMap (sin teléfono): solo rubros que encajan y fuera de la Zona Libre (mayoristas)
cand = json.load(open(BASE + '/cand.json', encoding='utf-8'))
def en_zona_libre(lat, lon): return 9.344 <= lat <= 9.3545 and -79.898 <= lon <= -79.884
tomar = {'taller', 'ferreteria', 'panaderia', 'restaurante', 'belleza', 'peluqueria', 'dentista', 'clinica', 'lavanderia', 'lavado', 'cafe', 'pinturas'}
genericos = {'barbería', 'mini super', 'rosticería', 'restaurant & lounge', 'decely', 'decently', 'fantastic casino', 'total', 'cruz', 'edwin', 'unico int.', 'isssa', 'issa', 'lafayette', 'westinghouse tienda', 'gvp international', 'international hardware', 'amor de casa', 'la casa del equipd pesado', 'fraiche', 'jenny liang', 'the buffalo store', "anthony's", 'auto fino', 'ketty', 'grand cafe', 'q.p.s.', 'lanco pinturas', 'sur'}
for rub, n, calle, ciudad, lat, lon, web in cand:
    if rub not in tomar or n.lower() in genericos or en_zona_libre(lat, lon): continue
    k = n.lower()
    if any(k.startswith(v) or v.startswith(k) for v in vistos): continue
    vistos.add(k)
    zona = (calle + ', ' if calle else '') + (ciudad or ('Sabanitas / Cativa' if lon > -79.86 else 'Colón'))
    filas.append([PRIORIDAD[rub] + 1, n, rub, zona, 'Ver en Google Maps', '', 'No encontrada', 'OpenStreetMap'])

filas.sort(key=lambda x: (x[0], x[2], x[1]))

wb = Workbook()
ws = wb.active
ws.title = 'Prospectos'
cab = ['Prioridad', 'Negocio', 'Rubro', 'Zona / dirección', 'Teléfono o red', 'Cómo contactar', 'Página web', 'Qué ofrecerle', 'Precio de referencia', 'Mensaje (hoja Mensajes)', 'Google Maps', 'Estado', 'Fecha de contacto', 'Notas', 'Fuente del dato']
ws.append(cab)
for p, n, r, z, t, como, web, f in filas:
    of, precio, msg = OFERTA[r]
    ws.append([p, n, r.capitalize(), z, t, como, web, of, precio, msg, 'Abrir mapa', 'Pendiente', '', '', f])
    c = ws.cell(row=ws.max_row, column=11); c.hyperlink = maps(n.split(' (')[0], 'Colón'); c.font = Font(color='1D4ED8', underline='single')

hdr = PatternFill('solid', fgColor='0A0D14')
for c in ws[1]:
    c.font = Font(bold=True, color='FFFFFF'); c.fill = hdr; c.alignment = Alignment(vertical='center', wrap_text=True)
colores = {1: 'DCFCE7', 2: 'FEF9C3', 3: 'F1F5F9', 4: 'F8FAFC'}
for row in ws.iter_rows(min_row=2):
    row[0].fill = PatternFill('solid', fgColor=colores.get(row[0].value, 'FFFFFF'))
    for c in row: c.alignment = Alignment(vertical='top', wrap_text=True)
for col, w in zip('ABCDEFGHIJKLMNO', [9, 34, 13, 34, 26, 22, 12, 40, 30, 14, 12, 14, 14, 30, 30]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'C2'
ws.auto_filter.ref = ws.dimensions
dv = DataValidation(type='list', formula1='"Pendiente,Contactado,Respondió,Reunión,Cotización enviada,Cliente,No le interesa"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f'L2:L{ws.max_row}')

# Hoja de mensajes
M = wb.create_sheet('Mensajes')
M.append(['Clave', 'Para quién', 'Mensaje para WhatsApp (cambia [NOMBRE] y [ENLACE])'])
MENSAJES = [
    ('taller', 'Talleres', 'Hola, buenas. Soy [TU NOMBRE], de Colón. Hago sistemas para negocios de la zona. Tengo uno para talleres donde anotas cada trabajo, lo que el cliente abonó, lo que falta por cobrar y lo que gastaste en repuestos, y a fin de mes ves cuánto ganaste de verdad. Funciona en el celular y cuesta $20 al mes. ¿Le puedo mostrar cómo se ve en 5 minutos? Aquí mis trabajos: [ENLACE]'),
    ('ferreteria', 'Ferreterías y materiales', 'Hola, buenas. Soy [TU NOMBRE], de Colón. Hago sistemas a medida para ferreterías: caja, inventario que avisa cuando algo se acaba, fiado de clientes y reportes de ventas por día. Le dejo una demostración que puede probar desde el celular, sin instalar nada: [ENLACE]/demo/ . Si le interesa, paso por el local y se lo muestro con sus productos.'),
    ('catalogo', 'Repuestos, pinturas, ropa', 'Hola, buenas. Soy [TU NOMBRE], de Colón. Hago catálogos en línea para negocios: sus productos con foto y precio, y el cliente le hace el pedido directo a su WhatsApp ya escrito. Así no tiene que mandar fotos una por una. ¿Le muestro un ejemplo? [ENLACE]'),
    ('comida', 'Panaderías, restaurantes, fondas, cafés', 'Hola, buenas. Soy [TU NOMBRE], de Colón. Hago páginas para negocios de comida: el menú con fotos y precios, y el cliente arma su pedido y le llega a su WhatsApp listo para preparar. Sale en Google cuando buscan comida en Colón. Mire un trabajo que hice para una heladería de aquí: [ENLACE] ¿Le interesa que le arme una?'),
    ('belleza', 'Salones y barberías', 'Hola, buenas. Soy [TU NOMBRE], de Colón. Hago páginas para salones y barberías: servicios, precios, fotos de sus trabajos y un botón para que el cliente pida su cita por WhatsApp. Es un enlace que puede poner en su Instagram. Desde $200. Aquí ve mis trabajos: [ENLACE]'),
    ('salud', 'Clínicas y dentistas', 'Buenas, le escribe [TU NOMBRE], de Colón. Hago páginas web para clínicas: servicios, horario, ubicación y citas por WhatsApp, para que los pacientes lo encuentren en Google. ¿Le puedo enviar una propuesta sin compromiso? Mis trabajos: [ENLACE]'),
    ('eventos', 'Eventos, fiestas, decoración', 'Hola, buenas. Soy [TU NOMBRE], de Colón. Tengo un sistema para negocios de eventos donde lleva cada evento, el adelanto que le dieron, lo que falta por cobrar y lo que gastó, y ve cuánto ganó en cada uno. Cuesta $20 al mes. También hago páginas con fotos de sus trabajos. ¿Le muestro? [ENLACE]'),
    ('servicio', 'Lavanderías, car wash y otros servicios', 'Hola, buenas. Soy [TU NOMBRE], de Colón. Ayudo a negocios de servicio a tener su página (precios, ubicación y pedidos por WhatsApp) y un sistema sencillo para llevar sus cobros y gastos. ¿Le puedo mostrar un ejemplo? [ENLACE]'),
    ('seguimiento', 'Si no responden (3 días después)', 'Hola de nuevo, ¿pudo ver lo que le mandé? Si quiere, paso por su local esta semana y se lo muestro en 5 minutos, sin compromiso.'),
]
for r in MENSAJES: M.append(list(r))
for c in M[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = hdr
M.column_dimensions['A'].width = 12; M.column_dimensions['B'].width = 30; M.column_dimensions['C'].width = 110
for row in M.iter_rows(min_row=2):
    for c in row: c.alignment = Alignment(vertical='top', wrap_text=True)

wb.save(OUT)
print('filas', len(filas), {p: sum(1 for f in filas if f[0] == p) for p in (1, 2, 3, 4)})
print('con contacto', sum(1 for f in filas if f[4] != 'Ver en Google Maps'))
