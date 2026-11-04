"""Agrega al Excel negocios de Colón con WhatsApp encontrados en micita.pa (04/11/2026)."""
import sys, urllib.parse
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

XLSX = sys.argv[1]
NUEVOS = [  # nombre, rubro, zona, whatsapp, qué ofrecer, precio
    ('Joy Barber Studio', 'peluqueria', 'Colón', '6225-6060', 'Página con servicios, precios y citas por WhatsApp', 'Presentación $150 (lanzamiento)'),
    ('Barbería El Patrón', 'peluqueria', 'Medalla Milagrosa, Los Lagos, Colón', '6300-8750', 'Página con servicios, precios y citas por WhatsApp', 'Presentación $150 (lanzamiento)'),
    ('Consultorio Vida Dental', 'dentista', 'Hospital Cuatro Altos, Colón', '6688-3950', 'Página con servicios y citas por WhatsApp + sistema de citas', 'Web $380 / presentación $150'),
    ('Donde La Chomba', 'restaurante', 'Terrazas de Sabanitas, planta alta', '6979-6685', 'Menú digital con pedidos y delivery por WhatsApp', 'Catálogo $350 (lanzamiento)'),
    ('Santizo Afro Kitchen', 'restaurante', 'Calle 4ta, Colón', '6861-4946', 'Menú digital con pedidos por WhatsApp', 'Catálogo $350 (lanzamiento)'),
    ('Taller y Repuestos Cabrini', 'taller', 'Colón', '6980-0454', 'NovaCore (taller) + catálogo de repuestos', 'NovaCore 1er mes gratis + catálogo $350'),
    ('Café Rumè', 'restaurante', 'Plaza El Sol, Colón', '6966-0263', 'Menú digital con pedidos y reservas por WhatsApp', 'Catálogo $350 (lanzamiento)'),
    ('Aspinwall Lounge Restaurant', 'restaurante', 'Río Duque, Buena Vista, Colón', '6980-4265', 'Menú digital con reservas y pedidos por WhatsApp', 'Catálogo $350 (lanzamiento)'),
    ('Pool Bar La Fogata #2', 'restaurante', 'Cativá, Colón', '6781-1527', 'Página con menú, eventos y reservas por WhatsApp', 'Presentación $150 (lanzamiento)'),
]
wb = load_workbook(XLSX)
ws = wb['Prospectos']
ya = {str(c.value).lower() for c in ws['B'][1:]}
n = 0
for nombre, rubro, zona, tel, oferta, precio in NUEVOS:
    if nombre.lower() in ya: continue
    ws.append([1, nombre, rubro.capitalize(), zona, tel, 'Celular (WhatsApp)', 'No encontrada', oferta, precio,
               {'peluqueria': 'belleza', 'dentista': 'salud', 'restaurante': 'comida', 'taller': 'taller'}[rubro], 'Abrir mapa', 'Pendiente', '', '', 'micita.pa (WhatsApp verificado)'])
    r = ws.max_row
    c = ws.cell(row=r, column=11); c.hyperlink = 'https://www.google.com/maps/search/?api=1&query=' + urllib.parse.quote(f'{nombre} {zona} Panamá'); c.font = Font(color='1D4ED8', underline='single')
    ws.cell(row=r, column=1).fill = PatternFill('solid', fgColor='DCFCE7')
    for cell in ws[r]: cell.alignment = Alignment(vertical='top', wrap_text=True)
    n += 1
ws.auto_filter.ref = ws.dimensions
for dv in ws.data_validations.dataValidation: dv.sqref = f'L2:L{ws.max_row}'
wb.save(XLSX)
print('agregados', n, 'filas', ws.max_row - 1)
