"""Importa al Excel los negocios de Colón de micita.pa (perfiles.json del rastreo del 04/11/2026)."""
import json, re, sys, urllib.parse
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

PERFILES, XLSX = sys.argv[1], sys.argv[2]

CAT = {  # categoría de micita -> (rubro, prioridad, qué ofrecer, precio)
    'Restaurante': ('restaurante', 2, 'Menú digital con pedidos por WhatsApp', 'Catálogo $350 (lanzamiento)'),
    'Restaurante de Carnes': ('restaurante', 2, 'Menú digital con pedidos y reservas', 'Catálogo $350 (lanzamiento)'),
    'Restaurante de cocina': ('restaurante', 2, 'Menú digital con pedidos por WhatsApp', 'Catálogo $350 (lanzamiento)'),
    'Restaurante de comida rápida': ('restaurante', 2, 'Menú digital con pedidos y delivery', 'Catálogo $350 (lanzamiento)'),
    'Restaurante de hamburguesas': ('restaurante', 2, 'Menú digital con pedidos y delivery', 'Catálogo $350 (lanzamiento)'),
    'Restaurante de mariscos': ('restaurante', 2, 'Menú digital con pedidos y reservas', 'Catálogo $350 (lanzamiento)'),
    'Pizzería': ('restaurante', 2, 'Menú digital con pedidos y delivery', 'Catálogo $350 (lanzamiento)'),
    'Café': ('cafe', 2, 'Menú digital con pedidos por WhatsApp', 'Presentación $150 (lanzamiento)'),
    'Bar': ('restaurante', 3, 'Página con menú y reservas', 'Presentación $150 (lanzamiento)'),
    'Panadería': ('panaderia', 2, 'Catálogo con encargos por WhatsApp', 'Catálogo $350 (lanzamiento)'),
    'Barbería': ('peluqueria', 2, 'Página con servicios y citas por WhatsApp', 'Presentación $150 (lanzamiento)'),
    'Peluquería': ('belleza', 2, 'Página con servicios y citas por WhatsApp', 'Presentación $150 (lanzamiento)'),
    'Estética y Belleza': ('belleza', 2, 'Página con servicios y citas por WhatsApp', 'Presentación $150 (lanzamiento)'),
    'Estética y Spa': ('belleza', 2, 'Página con servicios y citas + sistema de citas', 'Presentación $150 / web $380'),
    'Estética integral': ('belleza', 2, 'Página con servicios y citas por WhatsApp', 'Presentación $150 (lanzamiento)'),
    'Auto lavado y detallado': ('lavado', 2, 'Página con precios y turnos por WhatsApp', 'Presentación $150 (lanzamiento)'),
    'Taller mecánico': ('taller', 1, 'NovaCore (taller) + página con citas', 'NovaCore 1er mes gratis + página $150'),
    'Llantería': ('taller', 1, 'NovaCore (taller) + página', 'NovaCore 1er mes gratis + página $150'),
    'Taller de motocicletas': ('taller', 1, 'NovaCore (taller) + página', 'NovaCore 1er mes gratis + página $150'),
    'Chapistería y pintura': ('taller', 1, 'NovaCore (taller) + página', 'NovaCore 1er mes gratis + página $150'),
    'Odontología General': ('dentista', 1, 'Página web con citas + sistema de citas', 'Web $380 / presentación $150'),
    'Clínica dental': ('dentista', 1, 'Página web con citas + sistema de citas', 'Web $380 / presentación $150'),
    'Clínica de especialidades dentales': ('dentista', 1, 'Página web con citas + sistema de citas', 'Web $380 / presentación $150'),
    'Centro Médico Multiespecialidad': ('clinica', 1, 'Página web con citas + sistema de pacientes', 'Web $380 / sistema desde $900'),
    'Laboratorio Clínico': ('clinica', 1, 'Página con precios de exámenes y citas', 'Web $380'),
    'Fisioterapia y Rehabilitación Física': ('clinica', 1, 'Página con servicios y citas', 'Web $380 / presentación $150'),
    'Clínica Veterinaria': ('veterinaria', 1, 'Página con servicios y citas + sistema de pacientes', 'Web $380 / presentación $150'),
    'Gimnasio': ('gimnasio', 2, 'Página con planes + NovaCore para mensualidades', 'Presentación $150 + NovaCore'),
    'Alojamiento': ('hotel', 1, 'Página web con habitaciones y reservas por WhatsApp', 'Web $380'),
    'Operador turístico': ('tours', 1, 'Página web con tours y reservas por WhatsApp', 'Web $380'),
    'Operador Turístic…': ('tours', 1, 'Página web con tours y reservas por WhatsApp', 'Web $380'),
    'Firma de Abogados': ('abogados', 1, 'Página web profesional con citas', 'Web $380'),
    'Firma de Abogados en…': ('abogados', 1, 'Página web profesional con citas', 'Web $380'),
    'Consultoría Empresarial': ('contable', 1, 'Página web profesional + sistema de clientes', 'Web $380'),
    'Entretenimiento': ('sala', 1, 'Página con paquetes y reservas + NovaCore (eventos)', 'Web $380 + NovaCore'),
}
FUERA = re.compile(r'radisson|smart fit|dairy queen|duty free|shelter bay|ancon seguros|energyasia|beauty style internacional|ismael laguna|ecleens|champagne|margarita, corregimiento|multillantas z/l', re.I)

perf = json.load(open(PERFILES, encoding='utf-8'))
wb = load_workbook(XLSX)
ws = wb['Prospectos']
ya_tel = {d.replace('-', '') for c in ws['E'][1:] for d in re.findall(r'\d[\d-]{6,}\d', str(c.value or ''))}
ya_nom = {str(c.value).lower() for c in ws['B'][1:]} | {str(c.value).lower().split(' (')[0] for c in ws['B'][1:]}
vistos, n = set(), 0
for p in sorted(perf, key=lambda x: x['name'].lower()):
    if 'err' in p or p['cat'] not in CAT or FUERA.search(p['name']): continue
    wa = p['wa'][0][3:]
    if wa in vistos or wa in ya_tel or p['name'].lower() in ya_nom: continue
    vistos.add(wa)
    rubro, prio, oferta, precio = CAT[p['cat']]
    tel = f'{wa[:4]}-{wa[4:]}'
    zona = (p.get('area') or '').strip() or 'Colón'
    if zona.lower() != 'colón': zona = f'{zona}, Colón'
    nombre = re.sub(r'\s+', ' ', p['name']).strip()
    ws.append([prio, nombre, rubro.capitalize(), zona, tel, 'Celular (WhatsApp)', 'No tiene', oferta, precio, rubro, 'Abrir mapa', 'Pendiente', '', p['cat'], 'micita.pa (WhatsApp verificado) ' + p['url']])
    r = ws.max_row
    c = ws.cell(row=r, column=11); c.hyperlink = 'https://www.google.com/maps/search/?api=1&query=' + urllib.parse.quote(f'{nombre} {zona} Panamá'); c.font = Font(color='1D4ED8', underline='single')
    ws.cell(row=r, column=1).fill = PatternFill('solid', fgColor={1: 'DCFCE7', 2: 'FEF9C3', 3: 'F1F5F9'}[prio])
    for cell in ws[r]: cell.alignment = Alignment(vertical='top', wrap_text=True)
    n += 1
ws.auto_filter.ref = ws.dimensions
for dv in ws.data_validations.dataValidation: dv.sqref = f'L2:L{ws.max_row}'
wb.save(XLSX)
print('agregados', n, '· filas totales', ws.max_row - 1)
