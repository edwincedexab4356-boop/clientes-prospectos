"""Muestras genéricas por rubro para los negocios que no tienen una hecha a mano."""
import re, unicodedata
from openpyxl import load_workbook

PORTAFOLIO = 'https://edwin2030-cmd.github.io/dek-portafolio/'

def slugify(s):
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+', '-', s).strip('-')[:48]

def I(cat, ic, n, d, p, **k):
    return {'cat': cat, 'ic': ic, 'n': n, 'd': d, 'p': p, **k}

NOVACORE = lambda tit, txt, puntos: {'titulo': tit, 'texto': txt, 'puntos': puntos, 'btn': 'Quiero ver cómo funciona', 'link': None}
SIS_TIENDA = {'titulo': 'Y para su caja e inventario: un sistema a medida', 'texto': 'Además del catálogo, le hago un sistema para el mostrador. Pruebe una demostración ahora mismo.',
              'puntos': ['Caja rápida con ticket', 'Inventario que avisa cuando algo se acaba', 'Clientes y fiado', 'Ventas por día y por forma de pago'],
              'btn': 'Probar la demostración', 'link': PORTAFOLIO + 'demo/'}
SIS_COMIDA = NOVACORE('Y para el negocio: un sistema a medida', 'Lleve encargos, ventas del día y gastos sin cuaderno.',
                      ['Encargos con fecha y hora de entrega', 'Ventas del día por producto', 'Gastos de compras, gas y empleados', 'Cuánto le quedó al final del mes'])
SIS_CITAS = NOVACORE('Y para la agenda: un sistema de citas y clientes', 'Sepa quién viene hoy, quién le debe y cuánto entró en el mes.',
                     ['Agenda de citas del día y de la semana', 'Historial de cada cliente', 'Cobros, abonos y gastos', 'Reporte del mes'])
SIS_SERV = NOVACORE('Y para llevar su negocio: NovaCore', 'Un sistema sencillo, en el celular o la computadora.',
                    ['Cada trabajo u orden con su cliente y su precio', 'Lo que ya pagaron y lo que deben', 'Gastos del negocio', 'Cuánto ganó de verdad cada mes'])

FONDA = [I('Almuerzos', '🍛', 'Arroz con pollo', 'Con ensalada de papa y tajadas.', 3.50), I('Almuerzos', '🍲', 'Sancocho de gallina', 'Con arroz blanco.', 3.50),
         I('Almuerzos', '🍗', 'Pollo guisado', 'Con arroz, frijoles y ensalada.', 4.00), I('Almuerzos', '🥩', 'Bistec picado', 'Con arroz y lentejas.', 4.50),
         I('Almuerzos', '🍖', 'Ropa vieja', 'Con arroz y patacones.', 4.50), I('Almuerzos', '🐟', 'Pescado frito', 'Entero, con patacones y ensalada.', 6.50),
         I('Extras', '🍌', 'Patacones (porción)', 'Recién hechos.', 1.50), I('Bebidas', '🥤', 'Chicheme o jugo natural', 'Vaso grande.', 1.00)]
CHINO = [I('Platos', '🍜', 'Chow mein de pollo', 'Fideos salteados con vegetales.', 5.50), I('Platos', '🍚', 'Arroz frito especial', 'Con pollo, puerco y camarón.', 6.00),
         I('Platos', '🥟', 'Wantán frito (12)', 'Con salsa agridulce.', 4.00), I('Platos', '🍗', 'Pollo agridulce', 'Con arroz blanco.', 6.50),
         I('Platos', '🥢', 'Chop suey de res', 'Vegetales frescos y res.', 6.50), I('Sopas', '🍲', 'Sopa de wantán', 'Caldo de la casa.', 4.50),
         I('Bebidas', '🥤', 'Soda o té frío', '', 1.00)]
POLLO = [I('Combos', '🍗', 'Combo 2 presas', 'Con papas y soda.', 4.50), I('Combos', '🍗', 'Combo 3 presas', 'Con papas y soda.', 5.75),
         I('Familiar', '🪣', 'Balde de 8 presas', 'Con papas grandes.', 14.00), I('Familiar', '🪣', 'Balde de 12 presas', 'Con 2 papas grandes y soda de 2 L.', 20.00),
         I('Extras', '🍟', 'Papas fritas', 'Porción mediana.', 1.75), I('Extras', '🍌', 'Patacones', 'Porción.', 1.50), I('Bebidas', '🥤', 'Soda 2 litros', '', 2.50)]
PIZZA = [I('Pizzas', '🍕', 'Pizza de queso (mediana)', '8 porciones.', 8.00), I('Pizzas', '🍕', 'Pizza de pepperoni (mediana)', '8 porciones.', 9.50),
         I('Pizzas', '🍕', 'Pizza hawaiana (mediana)', 'Jamón y piña.', 9.50), I('Pizzas', '🍕', 'Pizza suprema (grande)', 'Todo lo de la casa.', 14.00),
         I('Extras', '🥖', 'Pan de ajo', '6 piezas.', 2.50), I('Extras', '🍗', 'Alitas (8)', 'BBQ o picantes.', 6.00), I('Bebidas', '🥤', 'Soda 2 litros', '', 2.50)]
DOMINICANO = [I('Platos', '🍌', 'Mangú con los tres golpes', 'Salami, queso frito y huevo.', 5.00), I('Platos', '🍲', 'Sancocho dominicano', 'Siete carnes, con arroz.', 6.00),
              I('Platos', '🍗', 'Pollo guisado', 'Con arroz y habichuelas.', 5.00), I('Platos', '🐖', 'Pernil asado', 'Con moro de guandú.', 6.50),
              I('Platos', '🐟', 'Pescado frito', 'Con tostones.', 7.00), I('Bebidas', '🥤', 'Morir soñando', 'Jugo de naranja con leche.', 2.00)]
BAR = [I('Para picar', '🍗', 'Alitas (10)', 'BBQ, picantes o al ajillo.', 7.00), I('Para picar', '🍟', 'Picada para compartir', 'Carne, chorizo, patacones y yuca.', 12.00),
       I('Para picar', '🦐', 'Ceviche de corvina', 'Con galletas o patacones.', 6.00), I('Platos', '🍖', 'Costillas BBQ', 'Con papas.', 9.00),
       I('Platos', '🐟', 'Filete de pescado', 'A la plancha o apanado.', 8.50), I('Bebidas', '🥤', 'Sodas y jugos', '', 1.50)]
CAFE = [I('Café', '☕', 'Café negro', 'Recién colado.', 1.00), I('Café', '☕', 'Café con leche', 'Taza grande.', 1.50), I('Café', '🧊', 'Frappé de café', 'Con crema.', 3.00),
        I('Desayunos', '🍳', 'Desayuno típico', 'Huevo, carne guisada, hojaldra o tortilla.', 4.50), I('Desayunos', '🥪', 'Sándwich de jamón y queso', 'En pan tostado.', 2.75),
        I('Dulces', '🍰', 'Porción de dulce', 'Pregunte por el del día.', 2.00), I('Bebidas', '🍹', 'Jugo natural', 'Maracuyá, piña o naranja.', 1.75)]
REFRES = [I('Jugos', '🍹', 'Batido de frutas', 'Fresa, papaya o guineo.', 2.00), I('Jugos', '🍹', 'Jugo natural', 'Naranja, piña o maracuyá.', 1.50),
          I('Raspados', '🍧', 'Raspado con leche condensada', 'Varios sabores.', 1.50), I('Comida', '🥪', 'Emparedado de la casa', '', 3.00),
          I('Comida', '🌭', 'Hot dog sencillo', '', 1.75), I('Comida', '🥟', 'Empanada', 'Carne o pollo.', 0.75)]
PANADERIA = [I('Pan', '🍞', 'Pan de molde', 'Del día.', 2.75), I('Pan', '🥖', 'Pan de flauta', 'Crujiente.', 0.50), I('Pan', '🥯', 'Bollitos de mantequilla (6)', '', 1.50),
             I('Dulces', '🥐', 'Cachitos', 'De queso o jamón.', 0.75), I('Dulces', '🍰', 'Dulce del día (porción)', '', 1.50), I('Dulces', '🍪', 'Galletas de la casa (docena)', '', 3.00),
             I('Por encargo', '🎂', 'Bizcocho de cumpleaños', 'Decorado, por encargo con 2 días.', 25, desde=1), I('Por encargo', '🧁', 'Cupcakes (docena)', 'Para fiestas.', 12)]
SALON = [I('Cabello', '✂️', 'Corte de dama', 'Lavado y secado.', 10, desde=1), I('Cabello', '🎨', 'Tinte', 'Color completo.', 35, desde=1), I('Cabello', '✨', 'Keratina', 'Alisado y brillo.', 60, desde=1),
         I('Cabello', '💇', 'Cepillado', '', 8), I('Uñas', '💅', 'Manicure', '', 10), I('Uñas', '🦶', 'Pedicure', '', 15), I('Uñas', '💎', 'Uñas acrílicas', 'Set completo.', 25),
         I('Rostro', '👁️', 'Diseño de cejas', '', 5), I('Rostro', '💄', 'Maquillaje', 'Para eventos.', 30, desde=1)]
BARBERIA = [I('Cortes', '💈', 'Corte de caballero', 'Con lavado.', 6), I('Cortes', '🧔', 'Corte y barba', '', 9), I('Cortes', '👦', 'Corte de niño', 'Hasta 12 años.', 5),
            I('Extras', '✏️', 'Diseño o rayas', '', 2), I('Extras', '🪒', 'Arreglo de barba', 'Con toalla caliente.', 4), I('Extras', '👁️', 'Cejas', '', 3), I('Extras', '🎨', 'Tinte de barba', '', 8)]
DENTISTA = [I('Consultas', '🦷', 'Consulta y diagnóstico', 'Revisión completa.', 20), I('Prevención', '🪥', 'Limpieza dental', 'Con ultrasonido.', 35),
            I('Tratamientos', '🦷', 'Resina (calza)', 'Del color del diente.', 40, desde=1), I('Tratamientos', '🦷', 'Extracción simple', '', 35, desde=1),
            I('Estética', '✨', 'Blanqueamiento', 'En consultorio.', 150), I('Ortodoncia', '😁', 'Consulta de ortodoncia', 'Brackets y retenedores.', 30)]
CLINICA = [I('Consultas', '🩺', 'Medicina general', 'Consulta y receta.', 20), I('Consultas', '👶', 'Pediatría', '', 25), I('Consultas', '🤰', 'Ginecología', '', 30),
           I('Servicios', '💉', 'Inyecciones y curaciones', '', 5, desde=1), I('Servicios', '📄', 'Certificado de salud', '', 15), I('Laboratorio', '🧪', 'Exámenes de laboratorio', 'Resultados el mismo día.', 8, desde=1)]
LAB = [I('Exámenes', '🧪', 'Hemograma completo', '', 8), I('Exámenes', '🩸', 'Glicemia (azúcar)', '', 4), I('Exámenes', '❤️', 'Perfil lipídico', 'Colesterol y triglicéridos.', 15),
       I('Exámenes', '💧', 'Urinálisis', '', 4), I('Exámenes', '🦠', 'Prueba de embarazo', 'En sangre.', 10), I('Paquetes', '📋', 'Chequeo básico', 'Hemograma, glicemia, orina y heces.', 22),
       I('Servicios', '🏠', 'Toma de muestra a domicilio', 'Colón y Sabanitas.', 5)]
LAVANDERIA = [I('Lavado', '🧺', 'Lavado y secado (carga)', 'Usted lava, nosotros ayudamos.', 4), I('Lavado', '👕', 'Lavado y doblado por libra', 'Lo dejamos listo.', 1, u='por libra'),
              I('Planchado', '👔', 'Planchado de camisa', '', 1.50), I('Tintorería', '🧥', 'Saco o chaqueta', 'Lavado en seco.', 6), I('Especiales', '🛏️', 'Edredón', 'Tamaño queen o king.', 8),
              I('Servicio', '🛵', 'Recogida y entrega', 'En Colón.', 2)]
LAVADO = [I('Lavado', '🚿', 'Lavado exterior', '', 5), I('Lavado', '🚗', 'Lavado completo', 'Por dentro y por fuera.', 10), I('Lavado', '⚙️', 'Lavado de motor', '', 8),
          I('Detallado', '✨', 'Encerado', '', 15), I('Detallado', '🧽', 'Lavado de tapicería', '', 40, desde=1), I('Detallado', '🧴', 'Aspirado', '', 4)]
PINTURAS = [I('Pinturas', '🪣', 'Pintura de agua (galón)', 'Interior y exterior.', 22), I('Pinturas', '🪣', 'Pintura de aceite (galón)', 'Para metal y madera.', 28),
            I('Pinturas', '🪣', 'Cubeta de pintura (5 galones)', 'Rinde unos 175 m².', 95), I('Accesorios', '🖌️', 'Rodillo con bandeja', '', 6), I('Accesorios', '🖌️', 'Brocha de 3"', '', 2.50),
            I('Accesorios', '🧴', 'Sellador', 'Galón.', 18), I('Servicio', '🎨', 'Igualación de color', 'Traiga su muestra.', 0, u='sin costo')]
REPUESTOS = [I('Mantenimiento', '🛢️', 'Aceite de motor (galón)', '5W-30 o 10W-40.', 24), I('Mantenimiento', '🧰', 'Filtro de aceite', 'Para la mayoría de modelos.', 6),
             I('Frenos', '🛑', 'Pastillas de freno', 'Delanteras.', 25, desde=1), I('Eléctrico', '🔋', 'Batería', 'Con garantía.', 95, desde=1), I('Eléctrico', '⚡', 'Bujías (juego de 4)', '', 16),
             I('Eléctrico', '💡', 'Bombillos', 'Halógenos y LED.', 4, desde=1), I('Suspensión', '🔧', 'Amortiguador', '', 45, desde=1)]
ROPA = [I('Dama', '👚', 'Blusas', 'Varios estilos y tallas.', 15, desde=1), I('Dama', '👗', 'Vestidos', '', 25, desde=1), I('Caballero', '👕', 'Camisas', '', 18, desde=1),
        I('Jeans', '👖', 'Jeans', 'Dama y caballero.', 25, desde=1), I('Niños', '🧒', 'Ropa de niños', '', 10, desde=1), I('Accesorios', '👜', 'Carteras', '', 20, desde=1)]
VARIEDADES = [I('Hogar', '🏠', 'Artículos para el hogar', '', 3, desde=1), I('Hogar', '🍳', 'Cocina', 'Ollas, sartenes y utensilios.', 5, desde=1),
              I('Escolar', '✏️', 'Útiles escolares', '', 1, desde=1), I('Regalos', '🎁', 'Regalos y detalles', '', 5, desde=1),
              I('Limpieza', '🧹', 'Limpieza', '', 2, desde=1), I('Tecnología', '🔌', 'Cables y cargadores', '', 4, desde=1)]

BURGER = [I('Hamburguesas', '🍔', 'Hamburguesa clásica', 'Carne de la casa, queso, vegetales y papas.', 7.50), I('Hamburguesas', '🍔', 'Hamburguesa doble con tocino', 'Doble carne, cheddar y tocino.', 10.50),
          I('Bowls', '🥗', 'Bowl de pollo', 'Arroz, pollo a la plancha y vegetales.', 8.50), I('Alitas', '🍗', 'Alitas BBQ (10)', 'Con papas y salsa ranch.', 9.00),
          I('Alitas', '🌶️', 'Alitas picantes (10)', 'Con papas.', 9.00), I('Extras', '🍟', 'Papas con queso y tocino', '', 4.50), I('Bebidas', '🥤', 'Batido o soda', '', 2.50)]
AFRO = [I('Del mar', '🐟', 'Pescado en salsa de coco', 'Con arroz con coco y patacones.', 9.50), I('Del mar', '🍲', 'Rondón', 'Sopa caribeña de mariscos y coco.', 10.00),
        I('Parrilla', '🍖', 'Costillas ahumadas', 'Con ensalada y yuca frita.', 11.00), I('Parrilla', '🍗', 'Pollo a la brasa', 'Medio pollo con arroz con coco.', 8.50),
        I('Criollo', '🍛', 'Arroz con coco y frijoles', 'Acompañante de la casa.', 3.00), I('Criollo', '🍌', 'Patacones (porción)', '', 2.00), I('Bebidas', '🥤', 'Agua de saril o jugo natural', '', 2.00)]
PARRILLA = [I('Cortes', '🥩', 'Churrasco 10 oz', 'Con papas y chimichurri.', 16.00), I('Cortes', '🥩', 'Picaña', 'Al punto que prefiera.', 18.00), I('Cortes', '🍖', 'Costillas a la parrilla', '', 14.00),
            I('Pollo', '🍗', 'Pechuga a la parrilla', 'Con ensalada.', 10.00), I('Entradas', '🧀', 'Provoleta', '', 6.00), I('Café', '☕', 'Café de especialidad', '', 2.50), I('Postres', '🍰', 'Postre del día', '', 4.00)]
MARISCOS = [I('Ceviches', '🦐', 'Ceviche de corvina', 'Con galletas o patacones.', 7.00), I('Ceviches', '🦐', 'Ceviche mixto', 'Corvina, camarón y pulpo.', 9.00),
            I('Del mar', '🐟', 'Pescado entero frito', 'Con patacones y ensalada.', 12.00), I('Del mar', '🦞', 'Camarones al ajillo', 'Con arroz con coco.', 13.00),
            I('Picadas', '🍖', 'Tabla de carnes para compartir', 'Res, cerdo, chorizo y yuca.', 18.00), I('Almuerzos', '🍛', 'Almuerzo criollo del día', 'Pregunte por el de hoy.', 6.00), I('Bebidas', '🍹', 'Jugos naturales', '', 2.50)]
POOLBAR = [I('Para picar', '🍗', 'Alitas (10)', 'BBQ o picantes.', 7.00), I('Para picar', '🍟', 'Picada de la casa', 'Para compartir.', 12.00), I('Para picar', '🌮', 'Nachos con todo', '', 6.50),
           I('Reservas', '🎱', 'Mesa de billar (1 hora)', '', 5.00), I('Reservas', '🎤', 'Karaoke / salón VIP', 'Para cumpleaños y eventos.', 0, u='a cotizar'), I('Bebidas', '🥤', 'Sodas y jugos', '', 1.50)]

def tipo(rubro, nombre):
    n = nombre.lower()
    if rubro == 'restaurante':
        if 'chomba' in n: return 'burger'
        if 'afro' in n: return 'afro'
        if 'rum' in n: return 'parrilla'
        if 'aspinwall' in n: return 'mariscos'
        if 'pool bar' in n: return 'poolbar'
        if any(k in n for k in ('hong kong', 'fuig', 'jenny liang', 'china')): return 'chino'
        if 'pollo' in n or 'chicken' in n: return 'pollo'
        if 'pizza' in n: return 'pizza'
        if 'dominicano' in n: return 'dominicano'
        if n.startswith('bar') or 'texas' in n: return 'bar'
        if 'refresquer' in n: return 'refres'
        if 'cafeter' in n: return 'cafe'
        return 'fonda'
    if rubro in ('belleza', 'peluqueria'): return 'barberia' if ('barber' in n or 'barbershop' in n) else 'salon'
    if rubro == 'clinica': return 'lab' if 'laborator' in n else 'clinica'
    return rubro

PLANT = {  # tipo: (rubro visible, modo, color, c2, eyebrow, h1, p, items, extraLbl, extraEj, extra)
    'burger':     ('Hamburguesas y comida rápida', 'pedido', '#dc2626', '#fee2e2', 'Hamburguesas y alitas', 'Hamburguesas artesanales con delivery', 'Hamburguesas, bowls y alitas. Pida por WhatsApp y se lo llevamos.', BURGER, '¿Delivery o para retirar? (dirección)', 'Delivery a Sabanitas', SIS_COMIDA),
    'afro':       ('Cocina afro-caribeña', 'pedido', '#15803d', '#dcfce7', 'Cocina afro-caribeña', 'El sabor del Caribe colonense', 'Pescado en coco, rondón y carnes ahumadas. Haga su pedido o reserve por WhatsApp.', AFRO, '¿Para llevar o reserva? (hora)', 'Reserva para 2, jueves 1 p. m.', SIS_COMIDA),
    'parrilla':   ('Restaurante y parrilla', 'pedido', '#7c2d12', '#ffedd5', 'Parrilla y café', 'Cortes a la parrilla en el corazón de Colón', 'Churrasco, picaña y café de especialidad. Reserve su mesa o pida por WhatsApp.', PARRILLA, '¿Reserva o para llevar? (hora)', 'Mesa para 4, sábado 7 p. m.', SIS_COMIDA),
    'mariscos':   ('Restaurante y lounge', 'pedido', '#0e7490', '#cffafe', 'Cocina de autor · Buena Vista', 'Mariscos, ceviches y picadas frente al río', 'Cocina con sabores locales. Reserve su mesa o pida para llevar por WhatsApp.', MARISCOS, '¿Reserva o para llevar? (hora)', 'Mesa para 6, domingo 1 p. m.', SIS_COMIDA),
    'poolbar':    ('Pool bar', 'pedido', '#4338ca', '#e0e7ff', 'Billar, karaoke y eventos', 'Billar, karaoke y buen ambiente en Cativá', 'Reserve su mesa de billar o el salón VIP para su cumpleaños por WhatsApp.', POOLBAR, '¿Para cuántas personas y qué día?', '8 personas, viernes 9 p. m.', SIS_COMIDA),
    'fonda':      ('Fonda y comida casera', 'pedido', '#b91c1c', '#fee2e2', 'Comida casera', 'Comida casera como en casa, lista para llevar', 'Almuerzos del día, sopas y frituras. Haga su pedido por WhatsApp y páselo a buscar caliente.', FONDA, '¿Para llevar o para comer aquí?', 'Para llevar a las 12:30', SIS_COMIDA),
    'chino':      ('Restaurante de comida china', 'pedido', '#b91c1c', '#fee2e2', 'Comida china', 'Su comida china favorita, pedida en un minuto', 'Chow mein, arroz frito y wantanes. Arme su pedido y se lo tenemos listo.', CHINO, '¿Para llevar o a domicilio? (dirección)', 'Para llevar', SIS_COMIDA),
    'pollo':      ('Pollo frito', 'pedido', '#ea580c', '#ffedd5', 'Pollo frito', 'El pollo más crujiente, en combo o en balde', 'Combos personales y baldes para la familia. Pídalo por WhatsApp y páselo a buscar.', POLLO, '¿Para llevar o a domicilio? (dirección)', 'Para llevar', SIS_COMIDA),
    'pizza':      ('Pizzería', 'pedido', '#c2410c', '#ffedd5', 'Pizzería', 'Pizza recién horneada, pedida por WhatsApp', 'Elija su pizza, agregue los extras y mande su pedido. Se la tenemos lista.', PIZZA, '¿Para llevar o a domicilio? (dirección)', 'A domicilio en Sabanitas', SIS_COMIDA),
    'dominicano': ('Restaurante dominicano', 'pedido', '#1d4ed8', '#dbeafe', 'Sabor dominicano', 'El sabor de Quisqueya aquí en Colón', 'Mangú, sancocho y pernil como en casa. Haga su pedido por WhatsApp.', DOMINICANO, '¿Para llevar o para comer aquí?', 'Para llevar', SIS_COMIDA),
    'bar':        ('Bar y restaurante', 'pedido', '#7c2d12', '#ffedd5', 'Bar y restaurante', 'Buena comida, buen ambiente', 'Picadas, alitas y platos fuertes. Reserve su mesa o pida para llevar por WhatsApp.', BAR, '¿Mesa para cuántos o para llevar?', 'Mesa para 4, viernes 8 p. m.', SIS_COMIDA),
    'refres':     ('Refresquería', 'pedido', '#db2777', '#fce7f3', 'Refresquería', 'Batidos, raspados y algo rico para comer', 'Refréscate con lo mejor de la casa. Pide por WhatsApp y pásalo a buscar.', REFRES, '¿A qué hora lo pasa a buscar?', 'En 20 minutos', SIS_COMIDA),
    'cafe':       ('Café y cafetería', 'pedido', '#78350f', '#fef3c7', 'Café', 'Su café de la mañana, listo cuando llegue', 'Café, desayunos y dulces. Pida por WhatsApp y no haga fila.', CAFE, '¿A qué hora lo pasa a buscar?', 'Mañana a las 7:30 a. m.', SIS_COMIDA),
    'panaderia':  ('Panadería y repostería', 'pedido', '#a16207', '#fef9c3', 'Panadería', 'Pan fresco todos los días y dulces por encargo', 'Pan del día, dulces y bizcochos para sus fiestas. Haga su pedido por WhatsApp.', PANADERIA, '¿A qué hora lo pasa a buscar?', 'Mañana a las 8:00 a. m.', SIS_COMIDA),
    'salon':      ('Salón de belleza', 'cita', '#be185d', '#fce7f3', 'Salón de belleza', 'Luzca increíble, reserve su cita en un minuto', 'Cabello, uñas, cejas y maquillaje. Elija su servicio y pida su cita por WhatsApp.', SALON, 'Día y hora que prefiere', 'Sábado a las 10:00 a. m.', SIS_CITAS),
    'barberia':   ('Barbería', 'cita', '#111827', '#e5e7eb', 'Barbería', 'Su corte, a su hora, sin esperar', 'Cortes, barba y diseños. Separe su turno por WhatsApp.', BARBERIA, 'Día y hora que prefiere', 'Hoy a las 4:00 p. m.', SIS_CITAS),
    'dentista':   ('Clínica dental', 'cita', '#0e7490', '#cffafe', 'Clínica dental', 'Una sonrisa sana, con cita en un minuto', 'Limpiezas, resinas, extracciones y blanqueamiento. Pida su cita por WhatsApp.', DENTISTA, 'Día y hora que prefiere', 'Martes en la mañana', SIS_CITAS),
    'clinica':    ('Clínica médica', 'cita', '#0f766e', '#ccfbf1', 'Clínica', 'Atención médica cerca de usted', 'Medicina general, especialidades y laboratorio. Pida su cita por WhatsApp.', CLINICA, 'Día y hora que prefiere', 'Jueves a las 9:00 a. m.', SIS_CITAS),
    'lab':        ('Laboratorio clínico', 'cita', '#0369a1', '#e0f2fe', 'Laboratorio clínico', 'Sus exámenes, rápidos y con resultados el mismo día', 'Vea los precios y separe su turno por WhatsApp. También vamos a su casa.', LAB, 'Día y hora que prefiere', 'Mañana a las 7:00 a. m. (en ayunas)', SIS_CITAS),
    'lavanderia': ('Lavandería', 'cita', '#2563eb', '#dbeafe', 'Lavandería', 'Su ropa limpia sin perder el día', 'Lavado, planchado y tintorería. Pida que pasemos a buscarla por WhatsApp.', LAVANDERIA, 'Dirección para recoger', 'Calle 9, casa 12', SIS_SERV),
    'lavado':     ('Lavado de autos', 'cita', '#0284c7', '#e0f2fe', 'Car wash', 'Su carro brillando como nuevo', 'Lavado, encerado y tapicería. Separe su turno por WhatsApp y no espere.', LAVADO, 'Carro y hora que prefiere', 'Corolla, sábado 9 a. m.', SIS_SERV),
    'taller':     ('Taller mecánico', 'cita', '#c2410c', '#ffedd5', 'Taller en Colón', 'Su carro en buenas manos', 'Revisamos su carro, le decimos qué tiene y cuánto cuesta. Pida su cita por WhatsApp.', None, 'Carro (marca, modelo y año)', 'Toyota Corolla 2015', None),
    'ferreteria': ('Ferretería y materiales', 'pedido', '#b45309', '#fef3c7', 'Ferretería', 'Todo para su obra, pedido por WhatsApp', 'Materiales, plomería, electricidad y pintura. Arme su pedido aquí y se lo tenemos listo.', None, '¿Lo retira o se lo llevamos? (dirección)', 'Lo retiro en la tienda', SIS_TIENDA),
    'pinturas':   ('Pinturas', 'pedido', '#7c3aed', '#ede9fe', 'Pinturas', 'El color que busca, listo para llevar', 'Pinturas, selladores y accesorios. Arme su pedido por WhatsApp.', PINTURAS, '¿Lo retira o se lo llevamos? (dirección)', 'Lo retiro', SIS_TIENDA),
    'repuestos':  ('Repuestos para autos', 'pedido', '#1e40af', '#dbeafe', 'Repuestos', 'El repuesto que necesita, sin dar vueltas', 'Pregunte por su repuesto, vea precios y apártelo por WhatsApp.', REPUESTOS, 'Carro (marca, modelo y año)', 'Hyundai Tucson 2014', SIS_TIENDA),
    'ropa':       ('Tienda de ropa', 'pedido', '#9d174d', '#fce7f3', 'Moda', 'Lo nuevo de la tienda, en su celular', 'Vea lo que llegó, elija su talla y apártelo por WhatsApp.', ROPA, 'Talla y color', 'Talla M, color negro', SIS_TIENDA),
    'otro':       ('Tienda', 'pedido', '#334155', '#e2e8f0', 'Tienda', 'Todo lo que necesita, pedido por WhatsApp', 'Vea nuestros productos y arme su pedido. Se lo tenemos listo.', VARIEDADES, '¿Lo retira o se lo llevamos? (dirección)', 'Lo retiro', SIS_TIENDA),
}

def iniciales(n):
    p = [w for w in re.split(r'\W+', n) if w and w.lower() not in ('de', 'del', 'la', 'el', 'y', 's', 'a')]
    return (''.join(w[0] for w in p[:2]) or n[:2]).upper()

def resto(xlsx, hechos, taller_items, ferre_items, extra_taller):
    ws = load_workbook(xlsx)['Prospectos']
    ya = {h['nombre'].lower() for h in hechos}
    claves = [h['clave'].lower() for h in hechos]
    out, usados = [], {h['slug'] for h in hechos}
    for r in ws.iter_rows(min_row=2, values_only=True):
        nombre_x, rubro, zona, tel = r[1], r[2].lower(), r[3], r[4] or ''
        if any(nombre_x.lower().startswith(c) for c in claves): continue
        nombre = re.sub(r'\s*\(.*\)$', '', nombre_x).strip()
        t = tipo(rubro, nombre)
        rub, modo, color, c2, eyebrow, h1, p, items, xl, xe, extra = PLANT[t]
        if t == 'taller': items, extra = taller_items, extra_taller
        if t == 'ferreteria': items = ferre_items
        if rubro == 'panaderia' and not nombre.lower().startswith(('panader', 'dulcer')): nombre_vis = 'Panadería ' + nombre
        else: nombre_vis = nombre
        digits = [d for d in re.findall(r'\d[\d-]{6,}\d', tel)]
        tel1 = digits[0] if digits else ''
        s = slugify(nombre_vis); base = s; k = 2
        while s in usados: s = f'{base}-{k}'; k += 1
        usados.add(s)
        z = zona if zona else 'Colón'
        out.append(dict(slug=s, clave=nombre_x, nombre=nombre_vis, ini=iniciales(nombre_vis), color=color, c2=c2, rubro=rub, zona=z,
                        tel=tel1, telLink=('+507' + tel1.replace('-', '')) if tel1 else '',
                        eyebrow=f'{eyebrow} · {z.split(",")[-1].strip()}', h1=h1, p=p, modo=modo, items=items, extraLbl=xl, extraEj=xe, extra=extra))
    return out
