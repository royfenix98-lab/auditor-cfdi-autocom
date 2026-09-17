import re
import os
import bz2
import csv
import json
import sqlite3
import unicodedata
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
import requests
import bd_supabase as _bd

# ---------------------------------------------------------
# 1. FUNCIÓN DE NORMALIZACIÓN (ELIMINA ACENTOS)
# ---------------------------------------------------------
def quitar_acentos(texto):
    """Convierte 'CÁMARA' en 'CAMARA' para evitar fallos de acentuación en Regex."""
    if not texto:
        return ""
    texto_norm = unicodedata.normalize('NFD', texto)
    return "".join(c for c in texto_norm if unicodedata.category(c) != 'Mn')

# ---------------------------------------------------------
# 2. BASE DE DATOS DE NOMBRES/DESCRIPCIONES DE CLAVES SAT
# ---------------------------------------------------------
NOMBRES_CLAVES_SAT = {
    # Vehículos y Equipos de Transporte
    "2510": "Vehículos de motor (Automóviles/Camionetas)",
    "2410": "Maquinaria y equipo de manejo de materiales",
    "2517": "Componentes y accesorios de vehículos",
    "2518": "Sistemas de cuerpo y carrocería",
    "2519": "Sistemas de propulsión y transmisión",
    "2520": "Sistemas de frenos y suspensión",

    # Maquinaria, Herramientas y Taller Automotriz
    "2010": "Maquinaria para montaje/soldadura",
    "2315": "Maquinaria para procesos industriales/taller",
    "2711": "Herramientas de mano",
    "4010": "Calefacción, ventilación y aire acondicionado",
    "4014": "Bombas y compresores",

    # Combustibles y lubricantes
    "15": "Combustibles, lubricantes, aditivos, ceras y materiales de limpieza",
    "1512": "Lubricantes, aceites y grasas",
    "1521": "Combustibles (gasolina, diésel, gas)",

    # Componentes eléctricos / manufactura
    "26": "Componentes y suministros eléctricos, de iluminación y generación de energía",
    "31": "Componentes y suministros de manufactura",
    "39": "Distribución y generación eléctrica, y suministros de control",
    "41": "Equipo de laboratorio, medición, observación y comprobación",

    # Limpieza y suministros generales (incluye limpieza/detallado automotriz)
    "44": "Suministros de oficina, accesorios y productos",
    "46": "Equipo y suministros de seguridad y control de vigilancia",
    "47": "Equipo de limpieza y productos",
    "4713": "Equipo y productos para limpieza y detallado automotriz",

    # Tecnología, Redes y Cómputo
    "4321": "Equipos informáticos y accesorios",
    "4322": "Equipos y componentes de comunicación/redes",
    "4323": "Software",
    "4512": "Equipos de imagen y fotografía",
    "5216": "Equipos de audio y video",

    # Mobiliario y Oficina
    "5610": "Muebles de oficina y comercio",
    "5611": "Mobiliario comercial e industrial",

    # Servicios, Honorarios y Seguros
    "72": "Servicios de construcción, mantenimiento, reparación e instalación",
    "78": "Servicios de transporte, almacenaje y correo",
    "7810": "Servicios de transporte de carga y fletes",
    "80": "Servicios de gestión de empresas, negocios y administración",
    "8010": "Servicios de consultoría de negocios y administración",
    "8011": "Servicios de recursos humanos",
    "8012": "Servicios legales",
    "8013": "Servicios inmobiliarios",
    "8016": "Servicios de gestión de empresas",
    "81": "Servicios basados en ingeniería, investigación y tecnología",
    "8110": "Servicios de ingeniería y arquitectura",
    "8111": "Servicios de tecnología de la información (TI)",
    "82": "Servicios editoriales, de diseño, artes gráficas y publicidad",
    "8210": "Servicios de publicidad y mercadotecnia",
    "84": "Servicios financieros y de seguros",
    "8411": "Servicios contables, de auditoría y fiscales",
    "8412": "Servicios de seguros (pólizas, primas)",
    "85": "Servicios médicos y de salud",
    "8512": "Servicios médicos y de salud",
    "86": "Servicios educativos y de capacitación",
    "90": "Servicios de viajes, alimentación, alojamiento y entretenimiento",

    # Publicidad/mercadotecnia (categorías nuevas, complementa "8210" ya
    # definida arriba en la sección de Servicios)
    "8211": "Servicios de gestión de medios de publicidad",
    "8212": "Servicios de relaciones públicas",
}

# ---------------------------------------------------------
# 2.5 CATÁLOGO OFICIAL SAT DESDE BASE DE DATOS LOCAL (sat_catalogos.db)
# ---------------------------------------------------------
# El catálogo interno NOMBRES_CLAVES_SAT de arriba se conserva como respaldo
# curado, pero la fuente principal de nombres de Clave ProdServ es la base
# SQLite oficial generada por el proyecto phpcfdi/resources-sat-catalogs
# (el mismo que ya usabas): ~52,000 claves con su descripción oficial.
#
# Estructura esperada (tabla del proyecto phpcfdi, CFDI 4.0):
#     tabla:  cfdi_40_productos_servicios
#     id                -> clave de 8 dígitos (ej. '47131800')
#     texto             -> descripción oficial del SAT
#     similares         -> sinónimos separados por coma (búsqueda)
#     vigencia_desde / vigencia_hasta -> vigencia de la clave
#
# Se declara aquí como constante para que sea fácil de cambiar si algún día
# el proyecto renombra la tabla o si usas tu propio esquema.
TABLA_CATALOGO_PRODSERV = "cfdi_40_productos_servicios"
TABLA_CATALOGO_PRODSERV_FALLBACK = "cfdi_productos_servicios"  # CFDI 3.3
COLUMNA_CATALOGO_CLAVE = "id"
COLUMNA_CATALOGO_TEXTO = "texto"

# Nombre del archivo de base de datos que la app busca junto a app.py.
NOMBRE_ARCHIVO_CATALOGO = "sat_catalogos.db"

# URL de descarga de la versión más reciente publicada por phpcfdi. El
# archivo viene comprimido con bzip2 y se descomprime al guardarlo.
URL_CATALOGO_SAT = "https://github.com/phpcfdi/resources-sat-catalogs/releases/latest/download/catalogs.db.bz2"

# Cada cuántos días se intenta buscar una actualización del catálogo.
DIAS_ENTRE_ACTUALIZACIONES_CATALOGO = 30


# ---------------------------------------------------------
# 3. PATRONES REGEX MULTILINGÜES Y DE MODELOS/MARCAS
# ---------------------------------------------------------
PALABRAS_ACTIVOS = [
    # Taller / Equipos / Audio / Foto / Industrial
    "CAMARA", "CAMERA", "ELEVADOR", "LIFT", "RAMPA", "RAMP", "DESMONTADORA", "BALANCEADORA",
    "ALINEADORA", "ALIGNER", "COMPRESOR", "COMPRESSOR", "PRENSA", "PRESS", "SCANNER", "ESCANER",
    "DIAGNOSTIC", "SOLDADORA", "WELDER", "GENERADOR", "GENERATOR", "LAVADORA", "WASHER", "TORNO",
    "LATHE", "RECTIFICADORA", "GRINDER", "TORRE", "TOWER", "PLANT", "PLANTAS", "CHILLER", "BOILER",
    r"AIR\s*CONDITIONING", "MINISPLIT",
    # Vehículos y Maquinaria Pesada / Marcas / Modelos
    "MONTACARGAS", "FORKLIFT", "APILADOR", "STACKER", "REMOLQUE", "TRAILER", "PLATAFORMA",
    "PLATFORM", "BACKHOE", "EXCAVATOR", "BULLDOZER", "CATERPILLAR", "CAT", "KOMATSU", "BOBCAT",
    r"JOHN\s*DEERE", "SNAP-ON", "SNAPON", "TOOLBOX", "CARAVAN", "SKID", "MOTOCICLETA", "MOTORCYCLE",
    # Cómputo / TI / Oficina / Marcas
    "COMPUTADORA", "COMPUTER", "LAPTOP", "NOTEBOOK", "SERVER", "SERVIDOR", "PROYECTOR", "PROJECTOR",
    "MONITOR", "DISPLAY", "SCREEN", "PANTALLA", "DRONE", "DRON", "IMPRESORA", "PRINTER",
    "MULTIFUNCTIONAL", "MULTIFUNCIONAL", "SWITCH", "ROUTER", "FIREWALL", "PBX", "CONMUTADOR",
    "TABLET", "WORKSTATION", "PLOTTER", "DESK", "ESCRITORIO", "CHAIR", "SILLA", "CABINET",
    "ARCHIVERO", "CREDENZA", "BOOKCASE", "LIBRERO", "THINKPAD", "DELL", "HP", "LENOVO", "CISCO",
    "FORTINET", "APPLE", "MACBOOK", "IPAD", "ZEBRA", "HONEYWELL", "SONY",
    # Equipo específico de agencia/taller automotriz (activo fijo del negocio,
    # no vehículo de reventa ni refacción)
    "ELEVADOR DE 2 POSTES", "ELEVADOR DE 4 POSTES", "GATO HIDRAULICO", "RAMPA HIDRAULICA",
    "ALINEADORA DE DIRECCION", "ESCANER AUTOMOTRIZ", "SCANNER AUTOMOTRIZ", "CARGADOR DE BATERIAS",
    "EQUIPO DE AIRE ACONDICIONADO AUTOMOTRIZ", "RECUPERADORA DE REFRIGERANTE", "CABINA DE PINTURA",
    "HORNO DE SECADO", "BANCO DE PRUEBAS", "GRUA TALLER", "GRUA DE PATIO", "MONTALLANTAS",
    "DESMONTALLANTAS", "BALANCEADORA DE LLANTAS", "EQUIPO DE DIAGNOSTICO OBD", "OBD2",
    "SOPORTE DE MOTOR HIDRAULICO", "PISTOLA DE IMPACTO NEUMATICA", "COMPRESOR DE TALLER",
]
PATRON_ACTIVOS = re.compile(r'\b(' + '|'.join(PALABRAS_ACTIVOS) + r')\b', re.IGNORECASE)

PALABRAS_REFACCIONES = [
    "REFACCION", "SPARE", "PART", "PARTS", "REPUESTO", "REPLACEMENT", "BALATA", r"BRAKE\s*PAD",
    "FILTRO", "FILTER", "LLANTA", "TIRE", "TYRE", "BUJIA", r"SPARK\s*PLUG", "ACEITE", "OIL",
    "ANTICONGELANTE", "COOLANT", "LUBRICANTE", "LUBRICANT", "GRASA", "GREASE", "FLUID", "BALERO",
    "BEARING", "RETEN", "SEAL", "EMPAQUE", "GASKET", "JUNTA", "BANDA", "BELT", "CORREA", "DISCO",
    "DISC", "ROTOR", "TAMBOR", "DRUM", "CALIPER", "AMORTIGUADOR", "SHOCK", "ABSORBER", "RESORTE",
    "SPRING", "HORQUILLA", "ROTULA", r"BALL\s*JOINT", "TERMINAL", "BIELETA", "BUJE", "BUSHING",
    "CREMALLERA", "PUMP", "BOMBA", "PISTON", "VALVULA", "VALVE", "INYECTOR", "INJECTOR",
    "RADIADOR", "RADIATOR", "TERMOSTATO", "THERMOSTAT", "CLUTCH", "EMBRAGUE", "TURBO", "CARTER",
    "POLEA", "PULLEY", "SENSOR", "BATERIA", "BATTERY", "ACUMULADOR", "ALTERNADOR", "ALTERNATOR",
    "MARCHA", "STARTER", "FARO", "HEADLIGHT", "CALAVERA", "TAILLIGHT", "FOCO", "BULB", "DEFENSA",
    "BUMPER", "FACIA", "PARRILLA", "GRILL", "COFRE", "HOOD", "SALPICADERA", "FENDER", "PUERTA",
    "DOOR", "ESPEJO", "MIRROR", "CRISTAL", "GLASS", "PARABRISAS", "WINDSHIELD", "TAPETE", "MAT",
    "CUBIERTA", "COVER", "TORNILLO", "SCREW", "BOLT", "TUERCA", "NUT", "ABRAZADERA", "CLAMP",
    # Refacciones/accesorios adicionales de agencia/taller automotriz
    "EMBLEMA", "MOLDURA", "TAPICERIA", "VOLANTE", "TABLERO", "ARNES", "MODULO", "ECU",
    "COMPUTADORA AUTOMOTRIZ", "SENSOR DE OXIGENO", "CATALIZADOR", "MOFLE", "ESCAPE", "SILENCIADOR",
    "CREMALLERA DE DIRECCION", "BOMBA DE GASOLINA", "BOMBA DE AGUA", "TERMOCONTACTO", "MANGUERA",
    "HOSE", "CHICOTE", "CABLE DE FRENO", "ZAPATA", "MAZA", "HOMOCINETICA", "CRUCETA", "FLECHA",
    "DIFERENCIAL", "TRANSMISION", "CAJA DE VELOCIDADES", "CLUTCH DE AIRE ACONDICIONADO",
]
PATRON_REFACCIONES = re.compile(r'\b(' + '|'.join(PALABRAS_REFACCIONES) + r')\b', re.IGNORECASE)

PALABRAS_HONORARIOS = [
    "HONORARIO", "HONORARIOS", "FEE", "FEES", "CONSULTORIA", "CONSULTING", "ASESORIA", "ADVISORY",
    "DICTAMEN", "AUDITORIA", "AUDIT", "LEGAL", "CONTABLE", "ACCOUNTING", "TAX", "FINANCIERO",
    "FINANCIAL", r"PROFESSIONAL\s*SERVICES", r"MANAGEMENT\s*FEE",
]
PATRON_HONORARIOS = re.compile(r'\b(' + '|'.join(PALABRAS_HONORARIOS) + r')\b', re.IGNORECASE)

# Palabras que sí describen con certeza un gasto de mantenimiento/reparación
# ya realizado (mano de obra de taller, servicio periódico), lo cual es un
# GASTO GENERAL legítimo y NO debe mandarse a revisión. Este es el ÚNICO
# camino "directo" a GENERAL que queda en clasificar_concepto(): cualquier
# otra cosa que no matchee ninguna categoría se manda a REVISION.
PALABRAS_MANTENIMIENTO = [
    "MANTENIMIENTO", "MAINTENANCE", "SERVICIO DE MANTENIMIENTO", "AFINACION", r"TUNE[\s-]*UP",
    r"CAMBIO\s*DE\s*ACEITE", r"OIL\s*CHANGE", "REPARACION", "REPAIR", "HOJALATERIA",
    r"PINTURA\s*AUTOMOTRIZ", "ENDEREZADO", "DESABOLLADO", r"BODY\s*SHOP", r"SERVICIO\s*DE\s*\d+\s*KM",
    "DIAGNOSTICO MECANICO", r"REVISION\s*MECANICA", r"SERVICIO\s*DE\s*GARANTIA", "GARANTIA",
    "WARRANTY", "SINIESTRO", "CALIBRACION", "CALIBRATION", "LAVADO Y ENCERADO", "DETALLADO",
]
PATRON_MANTENIMIENTO = re.compile(r'\b(' + '|'.join(PALABRAS_MANTENIMIENTO) + r')\b', re.IGNORECASE)

# NOTA: 8111 (TI), 8210/8211/8212 (publicidad/mercadotecnia) se sacaron de
# esta tupla a propósito porque ahora tienen su propia categoría más
# específica (SOFTWARE_TI / PUBLICIDAD_MERCADOTECNIA, ver más abajo) — no
# deben quedarse "atrapadas" aquí como Honorario Profesional genérico.
PREFIJOS_HONORARIOS = ('8010', '8011', '8012', '8013', '8015', '8016', '8110', '8112', '8113', '8114', '8310', '8311', '8312', '8411', '8512')
PREFIJOS_ACTIVOS = ('20', '21', '22', '23', '24', '25', '26', '27', '30', '31', '43', '44', '45', '4512', '56', '7315', '8111')
PREFIJOS_REFACCIONES = ('2517', '2518', '2519', '2520', '3912', '4010', '4014', '3120', '3026', '1512')
PREFIJOS_FLETES = ('7810',)

# Combustibles y lubricantes para la flotilla/inventario (muy común y con
# tratamiento fiscal propio en automotrices: IEPS, monederos electrónicos,
# complemento "Estado de Cuenta de Combustibles").
PALABRAS_COMBUSTIBLE = [
    "GASOLINA", "GASOLINE", "DIESEL", r"DI[EÉ]SEL", "COMBUSTIBLE", "FUEL", "MAGNA", "PREMIUM",
    r"MONEDERO\s*ELECTRONICO", r"VALE\s*DE\s*GASOLINA", r"ESTACION\s*DE\s*SERVICIO", "GAS LP",
    r"GAS\s*NATURAL", "OCTANAJE",
]
PATRON_COMBUSTIBLE = re.compile(r'\b(' + '|'.join(PALABRAS_COMBUSTIBLE) + r')\b', re.IGNORECASE)
PREFIJOS_COMBUSTIBLE = ('1521',)

# Seguros (pólizas de flotilla, responsabilidad civil, seguro de agencia,
# etc.) — importante separarlo de Honorarios: fiscalmente no se retiene
# igual y suele amortizarse como gasto pagado por anticipado.
PALABRAS_SEGUROS = [
    "SEGURO", "INSURANCE", r"POLIZA\s*DE\s*SEGURO", "POLICY", r"PRIMA\s*DE\s*SEGURO", "ASEGURADORA",
    "COBERTURA", "DEDUCIBLE", r"RESPONSABILIDAD\s*CIVIL", "FIANZA",
]
PATRON_SEGUROS = re.compile(r'\b(' + '|'.join(PALABRAS_SEGUROS) + r')\b', re.IGNORECASE)
PREFIJOS_SEGUROS = ('8412', '8413', '8414')

# Software, licencias y TI (muy común: DMS/sistema de agencia tipo Quiter,
# licencias de Office/antivirus, hosting, soporte técnico). Se separa de
# ACTIVO_FIJO porque casi siempre es un gasto/suscripción recurrente, no la
# compra de un bien de uso duradero.
PALABRAS_SOFTWARE_TI = [
    "SOFTWARE", "LICENCIA", "LICENSE", "LICENCIAMIENTO", "SUSCRIPCION", "SUBSCRIPTION",
    "SAAS", r"NUBE\b", "CLOUD", "HOSTING", "DOMINIO", "DOMAIN", r"SOPORTE\s*TECNICO",
    r"SOPORTE\s*T\.?I\.?", "ANTIVIRUS", "FIREWALL", "BACKUP", "RESPALDO", "SERVIDOR VIRTUAL",
    "OFFICE 365", "MICROSOFT 365", "ERP", "DMS", "QUITER", "CRM", "HELP DESK", "MESA DE AYUDA",
]
PATRON_SOFTWARE_TI = re.compile(r'\b(' + '|'.join(PALABRAS_SOFTWARE_TI) + r')\b', re.IGNORECASE)
PREFIJOS_SOFTWARE_TI = ('4323', '8111')

# Publicidad y mercadotecnia (muy común en agencias: campañas, redes
# sociales, promocionales de piso de venta, rotulación).
PALABRAS_PUBLICIDAD_MERCADOTECNIA = [
    "PUBLICIDAD", "ADVERTISING", r"MERCADOTECNIA", "MARKETING", "CAMPA[ÑN]A", "CAMPAIGN",
    "PROMOCIONAL", "PROMOTIONAL", "ROTULACION", "LONA", "ESPECTACULAR", "BANNER", "VOLANTES",
    "REDES SOCIALES", "SOCIAL MEDIA", "AGENCIA DE PUBLICIDAD", "PAUTA", "GOOGLE ADS",
    "FACEBOOK ADS", "META ADS", "DISEÑO GRAFICO", "IMPRENTA", "STAND", "EXHIBICION",
]
PATRON_PUBLICIDAD_MERCADOTECNIA = re.compile(r'\b(' + '|'.join(PALABRAS_PUBLICIDAD_MERCADOTECNIA) + r')\b', re.IGNORECASE)
PREFIJOS_PUBLICIDAD_MERCADOTECNIA = ('8210', '8211', '8212')

# Arrendamiento (renta de local/showroom, renta de vehículos de flotilla o
# de cortesía, renta de equipo de taller/cómputo). Se separa de Honorarios
# porque fiscalmente no se retiene igual (10% ISR/10.6667% IVA solo aplica a
# arrendamiento de PERSONA FÍSICA, régimen 606 — ver validación en app.py).
PALABRAS_ARRENDAMIENTO = [
    "ARRENDAMIENTO", "RENTA", "RENTAL", "LEASE", "LEASING", "ALQUILER", r"RENTA\s*DE\s*AUTO",
    r"RENTA\s*DE\s*LOCAL", r"RENTA\s*DE\s*EQUIPO", "COMODATO",
]
PATRON_ARRENDAMIENTO = re.compile(r'\b(' + '|'.join(PALABRAS_ARRENDAMIENTO) + r')\b', re.IGNORECASE)
PREFIJOS_ARRENDAMIENTO = ()

# Capacitación (cursos y certificaciones de marca/fabricante para
# vendedores y técnicos, muy frecuente en agencias automotrices).
PALABRAS_CAPACITACION = [
    "CAPACITACION", "CAPACITACIÓN", "ENTRENAMIENTO", "TRAINING", "CURSO", "COURSE",
    "CERTIFICACION", "CERTIFICACIÓN", "CERTIFICATION", "DIPLOMADO", "TALLER DE CAPACITACION",
    "WORKSHOP", "SEMINARIO", "SEMINAR",
]
PATRON_CAPACITACION = re.compile(r'\b(' + '|'.join(PALABRAS_CAPACITACION) + r')\b', re.IGNORECASE)
PREFIJOS_CAPACITACION = ('86',)

UMBRAL_ACTIVO_FIJO = 12000.0

# Umbral a partir del cual un concepto "no clasificado" se marca como
# REVISION (en vez de despacharse directo a GENERAL). Se deja deliberadamente
# bajo: la política del negocio es que casi NADA no reconocido debe colarse
# como GENERAL sin más — solo partidas verdaderamente insignificantes (unos
# cuantos pesos, ej. redondeos o cargos menores) se quedan como GENERAL sin
# pasar por revisión manual / IA.
UMBRAL_AMBIGUEDAD_IA = 1.0

# Porcentaje mínimo de conceptos/líneas de UNA factura que deben cumplir
# ("Cumple" == "SI") para que, aunque existan una o más líneas en REVISION o
# incoherentes, la factura completa se considere aceptable por "regla de
# mayoría". Ej. con el valor por defecto (0.80): si 8 de 10 conceptos están
# correctos, la factura pasa como ACEPTADA (con nota) en vez de quedar en
# REVISIÓN MANUAL solo por 1-2 líneas dudosas.
#
# IMPORTANTE: esta regla de mayoría NUNCA aplica sobre bloqueos críticos de
# cumplimiento fiscal a nivel factura completa — lista negra 69-B, CFDI
# cancelado/no encontrado en el SAT, retenciones de ISR/IVA fuera de norma, o
# corte de PUE del último jueves del mes. Esos siguen bloqueando la factura
# sin importar qué tan bien esté clasificado el resto de los conceptos.
UMBRAL_PORCENTAJE_CONCEPTOS_CORRECTOS = 0.80

# ---------------------------------------------------------
# 3.5 CATÁLOGO DE "REGLAS ENSEÑABLES" (lista blanca para la IA)
# ---------------------------------------------------------
# Esta es la ÚNICA superficie que la función de "enseñar reglas nuevas" de
# app.py puede tocar. La IA nunca genera código Python libre: solo propone
# "agregar un elemento a esta lista/tupla/diccionario" o "cambiar este
# número", y app.py valida que el nombre exista aquí antes de escribir nada
# en disco. Si agregas una lista/umbral nuevo al motor de reglas y quieres
# que también sea enseñable por la IA, regístralo aquí.
REGLAS_EDITABLES = {
    "PALABRAS_ACTIVOS": {
        "tipo": "lista", "categoria": "ACTIVO_FIJO",
        "descripcion": "Palabras/frases en la descripción que indican Activo Fijo.",
    },
    "PALABRAS_REFACCIONES": {
        "tipo": "lista", "categoria": "REFACCION",
        "descripcion": "Palabras/frases en la descripción que indican Refacción.",
    },
    "PALABRAS_HONORARIOS": {
        "tipo": "lista", "categoria": "HONORARIO_PROFESIONAL",
        "descripcion": "Palabras/frases en la descripción que indican Honorario/Servicio Profesional.",
    },
    "PALABRAS_MANTENIMIENTO": {
        "tipo": "lista", "categoria": "GENERAL",
        "descripcion": "Palabras/frases que indican un gasto de mantenimiento/reparación ya realizado (único camino directo a GENERAL; todo lo demás no clasificado se manda a REVISION).",
    },
    "PREFIJOS_HONORARIOS": {
        "tipo": "tupla", "categoria": "HONORARIO_PROFESIONAL",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Honorario/Servicio Profesional.",
    },
    "PREFIJOS_ACTIVOS": {
        "tipo": "tupla", "categoria": "ACTIVO_FIJO",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Activo Fijo.",
    },
    "PREFIJOS_REFACCIONES": {
        "tipo": "tupla", "categoria": "REFACCION",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Refacción.",
    },
    "PREFIJOS_FLETES": {
        "tipo": "tupla", "categoria": "FLETE",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Flete/Autotransporte.",
    },
    "PALABRAS_COMBUSTIBLE": {
        "tipo": "lista", "categoria": "COMBUSTIBLE",
        "descripcion": "Palabras/frases en la descripción que indican Combustible/Lubricante.",
    },
    "PREFIJOS_COMBUSTIBLE": {
        "tipo": "tupla", "categoria": "COMBUSTIBLE",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Combustible.",
    },
    "PALABRAS_SEGUROS": {
        "tipo": "lista", "categoria": "SEGUROS",
        "descripcion": "Palabras/frases en la descripción que indican Seguros/Pólizas.",
    },
    "PREFIJOS_SEGUROS": {
        "tipo": "tupla", "categoria": "SEGUROS",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Seguros.",
    },
    "PALABRAS_SOFTWARE_TI": {
        "tipo": "lista", "categoria": "SOFTWARE_TI",
        "descripcion": "Palabras/frases en la descripción que indican Software/TI/licenciamiento.",
    },
    "PREFIJOS_SOFTWARE_TI": {
        "tipo": "tupla", "categoria": "SOFTWARE_TI",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Software/TI.",
    },
    "PALABRAS_PUBLICIDAD_MERCADOTECNIA": {
        "tipo": "lista", "categoria": "PUBLICIDAD_MERCADOTECNIA",
        "descripcion": "Palabras/frases en la descripción que indican Publicidad/Mercadotecnia.",
    },
    "PREFIJOS_PUBLICIDAD_MERCADOTECNIA": {
        "tipo": "tupla", "categoria": "PUBLICIDAD_MERCADOTECNIA",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Publicidad/Mercadotecnia.",
    },
    "PALABRAS_ARRENDAMIENTO": {
        "tipo": "lista", "categoria": "ARRENDAMIENTO",
        "descripcion": "Palabras/frases en la descripción que indican Arrendamiento/Renta.",
    },
    "PREFIJOS_ARRENDAMIENTO": {
        "tipo": "tupla", "categoria": "ARRENDAMIENTO",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Arrendamiento/Renta.",
    },
    "PALABRAS_CAPACITACION": {
        "tipo": "lista", "categoria": "CAPACITACION",
        "descripcion": "Palabras/frases en la descripción que indican Capacitación/Cursos.",
    },
    "PREFIJOS_CAPACITACION": {
        "tipo": "tupla", "categoria": "CAPACITACION",
        "descripcion": "Prefijos de Clave ProdServ SAT que indican Capacitación.",
    },
    "TASA_RET_ISR_RESICO_PF": {
        "tipo": "numero", "categoria": "RETENCIONES",
        "descripcion": (
            "Tasa de retención de ISR para emisor RESICO Persona Física (régimen 626), art. "
            "113-J LISR. Actual: 1.25%. NO aplica a RESICO Persona Moral: ese régimen (mismo "
            "código 626) calcula su propio ISR con tasa corporativa y no genera retención."
        ),
    },
    "TASA_RET_ISR_SERVICIOS_PF": {
        "tipo": "numero", "categoria": "RETENCIONES",
        "descripcion": "Tasa de retención de ISR por servicios profesionales de Persona Física. Actual: 10%.",
    },
    "TASA_RET_ISR_ARRENDAMIENTO_PF": {
        "tipo": "numero", "categoria": "RETENCIONES",
        "descripcion": "Tasa de retención de ISR por arrendamiento de Persona Física (régimen 606). Actual: 10%.",
    },
    "TASA_RET_IVA_SERVICIOS_PF": {
        "tipo": "numero", "categoria": "RETENCIONES",
        "descripcion": "Tasa de retención de IVA (2/3 partes) por servicios profesionales de PF. Actual: 10.6667%.",
    },
    "TASA_RET_IVA_ARRENDAMIENTO_PF": {
        "tipo": "numero", "categoria": "RETENCIONES",
        "descripcion": "Tasa de retención de IVA (2/3 partes) por arrendamiento de PF. Actual: 10.6667%.",
    },
    "TASA_RET_IVA_FLETES": {
        "tipo": "numero", "categoria": "RETENCIONES",
        "descripcion": "Tasa de retención de IVA por autotransporte terrestre de carga. Actual: 4%.",
    },
    "NOMBRES_CLAVES_SAT": {
        "tipo": "diccionario", "categoria": None,
        "descripcion": "Nombre legible de una clave o familia de clave SAT (solo cosmético, no clasifica).",
    },
    "UMBRAL_ACTIVO_FIJO": {
        "tipo": "numero", "categoria": "ACTIVO_FIJO",
        "descripcion": "Importe mínimo (MXN) por concepto para considerarse Activo Fijo.",
    },
    "UMBRAL_AMBIGUEDAD_IA": {
        "tipo": "numero", "categoria": None,
        "descripcion": "Importe mínimo para mandar un concepto no clasificado a REVISION en vez de GENERAL. Se mantiene muy bajo a propósito: casi todo lo que no matchea con certeza debe ir a REVISION, salvo partidas de centavos.",
    },
    "UMBRAL_PORCENTAJE_CONCEPTOS_CORRECTOS": {
        "tipo": "numero", "categoria": None,
        "descripcion": "Porcentaje (0-1) de conceptos que deben cumplir para aceptar la factura completa por regla de mayoría, aunque queden líneas en REVISION. No aplica a bloqueos críticos (69-B, CFDI cancelado, retenciones, corte PUE).",
    },
}


def obtener_valor_actual(nombre_regla):
    """Valor actual (en memoria, del módulo ya cargado) de una regla enseñable."""
    return globals().get(nombre_regla)


def proponer_texto_nuevo_bloque(nombre_regla, valor_nuevo):
    """
    Construye el fragmento de texto Python que se insertará antes del
    cierre del bloque (lista/tupla/diccionario), según el tipo registrado
    en REGLAS_EDITABLES. NO escribe nada en disco — eso lo hace app.py
    después de que el usuario confirma, usando insertar_antes_del_cierre().
    """
    info = REGLAS_EDITABLES.get(nombre_regla)
    if info is None:
        raise ValueError(f"'{nombre_regla}' no es una regla enseñable reconocida.")

    tipo = info["tipo"]
    if tipo == "numero":
        try:
            numero = float(str(valor_nuevo).replace(",", ""))
        except (TypeError, ValueError):
            raise ValueError("El nuevo valor del umbral debe ser un número.")
        if numero <= 0:
            raise ValueError("El umbral debe ser un número positivo.")
        return repr(numero)
    if tipo == "lista":
        palabra = str(valor_nuevo).strip().upper()
        if not palabra or not re.match(r'^[A-Z0-9ÁÉÍÓÚÑ\\\s\*]+$', palabra):
            raise ValueError("La palabra/frase a agregar tiene caracteres no permitidos.")
        return f'    "{palabra}",\n'
    if tipo == "tupla":
        prefijo = str(valor_nuevo).strip()
        if not prefijo.isdigit():
            raise ValueError("El prefijo de Clave ProdServ SAT debe ser numérico.")
        return f'"{prefijo}", '
    if tipo == "diccionario":
        if not isinstance(valor_nuevo, dict) or "clave" not in valor_nuevo or "nombre" not in valor_nuevo:
            raise ValueError("Para NOMBRES_CLAVES_SAT, valor_nuevo debe ser {'clave': ..., 'nombre': ...}.")
        clave = str(valor_nuevo["clave"]).strip()
        nombre = str(valor_nuevo["nombre"]).strip().replace('"', "'")
        if not clave.isdigit():
            raise ValueError("La clave SAT debe ser numérica.")
        return f'    "{clave}": "{nombre}",\n'
    raise ValueError(f"Tipo de regla '{tipo}' no soportado para inserción automática.")


def insertar_antes_del_cierre(codigo_fuente, nombre_regla, texto_nuevo):
    """
    Localiza el bloque `NOMBRE_REGLA = [...]` / `(...)` / `{...}` dentro del
    texto fuente de reglas_sat.py (contando paréntesis/corchetes/llaves para
    no confundirse con símbolos dentro de los propios valores) e inserta
    `texto_nuevo` justo antes del carácter de cierre. Regresa el código
    fuente completo ya modificado. Lanza ValueError si no encuentra el
    bloque — nunca escribe una modificación a medias.
    """
    info = REGLAS_EDITABLES.get(nombre_regla)
    if info is None:
        raise ValueError(f"'{nombre_regla}' no es una regla enseñable reconocida.")

    if info["tipo"] == "numero":
        patron = re.compile(rf'^{re.escape(nombre_regla)}\s*=\s*[0-9.]+', re.MULTILINE)
        if not patron.search(codigo_fuente):
            raise ValueError(f"No se encontró la asignación de '{nombre_regla}' en el archivo.")
        return patron.sub(f'{nombre_regla} = {texto_nuevo}', codigo_fuente, count=1)

    apertura_por_tipo = {"lista": "[", "tupla": "(", "diccionario": "{"}
    cierre_por_tipo = {"lista": "]", "tupla": ")", "diccionario": "}"}
    apertura = apertura_por_tipo[info["tipo"]]
    cierre = cierre_por_tipo[info["tipo"]]

    patron_inicio = re.compile(rf'{re.escape(nombre_regla)}\s*=\s*{re.escape(apertura)}')
    coincidencia = patron_inicio.search(codigo_fuente)
    if not coincidencia:
        raise ValueError(f"No se encontró la asignación de '{nombre_regla}' en el archivo.")

    idx_apertura = coincidencia.end() - 1
    profundidad = 0
    idx = idx_apertura
    while idx < len(codigo_fuente):
        if codigo_fuente[idx] == apertura:
            profundidad += 1
        elif codigo_fuente[idx] == cierre:
            profundidad -= 1
            if profundidad == 0:
                break
        idx += 1
    else:
        raise ValueError(f"No se encontró el cierre del bloque de '{nombre_regla}'.")

    idx_cierre = idx

    # Si el último elemento del bloque no termina en coma (por ejemplo el
    # último renglón de un diccionario escrito sin coma final), hay que
    # agregarla antes de insertar el nuevo elemento o el archivo queda con
    # un error de sintaxis.
    j = idx_cierre - 1
    while j > idx_apertura and codigo_fuente[j] in ' \t\r\n':
        j -= 1
    necesita_coma = j > idx_apertura and codigo_fuente[j] not in (apertura, ',')
    if necesita_coma:
        return (
            codigo_fuente[:j + 1] + ','
            + codigo_fuente[j + 1:idx_cierre] + texto_nuevo + codigo_fuente[idx_cierre:]
        )
    return codigo_fuente[:idx_cierre] + texto_nuevo + codigo_fuente[idx_cierre:]

# ---------------------------------------------------------
# 4. FUNCIONES DE EVALUACIÓN INTELIGENTE
# ---------------------------------------------------------
def obtener_nombre_clave_sat(clave, catalogo_externo=None):
    """
    Devuelve "12345678 - Nombre oficial del SAT" para una Clave ProdServ.

    Orden de resolución:
      1) `catalogo_externo` — dict {clave: texto} cargado desde la base
         oficial `sat_catalogos.db` (tabla cfdi_40_productos_servicios del
         proyecto phpcfdi/resources-sat-catalogs). Es el catálogo COMPLETO
         (~52,000 claves) y por eso va primero: con él casi ninguna clave
         debería quedar sin nombre oficial.
      2) `NOMBRES_CLAVES_SAT` — catálogo interno curado de AUTOCOM, que se
         conserva como respaldo para cuando la base no esté disponible
         (primer arranque, archivo movido, etc.) y porque sus nombres están
         redactados en términos de agencia automotriz.
      3) Familia (4 dígitos) y división (2 dígitos) del catálogo interno.
    """
    if not clave or len(clave) < 2:
        return "CLAVE NO ESPECIFICADA"

    clave = clave.strip()

    if catalogo_externo:
        nombre_oficial = catalogo_externo.get(clave)
        if nombre_oficial:
            return f"{clave} - {nombre_oficial}"

    if clave in NOMBRES_CLAVES_SAT:
        return f"{clave} - {NOMBRES_CLAVES_SAT[clave]}"
    
    familia = clave[:4]
    if familia in NOMBRES_CLAVES_SAT:
        return f"{clave} - {NOMBRES_CLAVES_SAT[familia]} [por familia {familia}]"
    
    division = clave[:2]
    if division in NOMBRES_CLAVES_SAT:
        return f"{clave} - {NOMBRES_CLAVES_SAT[division]} [por división {division}]"
    
    # IMPORTANTE: esto NO significa que el SAT clasifique la clave como
    # "General" — es una etiqueta interna de AUTOCOM que indica que esta
    # clave específica, su familia (4 dígitos) y su división (2 dígitos) no
    # están ni en el catálogo oficial cargado ni en el catálogo interno
    # `NOMBRES_CLAVES_SAT`. Por eso `clasificar_concepto` manda estos casos
    # a REVISION en vez de dejarlos pasar como GENERAL. Si aparece seguido,
    # lo más probable es que `sat_catalogos.db` no se esté cargando: revise
    # el estatus del catálogo en la barra lateral.
    return f"{clave} - [NO CATALOGADA: no está en el catálogo SAT cargado; verifique manualmente]"

def es_activo_fijo(clave_prod, descripcion, importe):
    """Conservada por compatibilidad/uso puntual: valida si, YA SABIENDO que
    la Clave ProdServ es de familia Activo Fijo, el importe y la
    descripción no lo descartan (mantenimiento, refacción, etc.). La
    clasificación real ya no pasa por aquí — ver `clasificar_concepto`."""
    if importe < UMBRAL_ACTIVO_FIJO:
        return False

    desc_limpia = quitar_acentos(descripcion).upper()

    if PATRON_REFACCIONES.search(desc_limpia) or clave_prod.startswith(PREFIJOS_REFACCIONES):
        return False

    if clave_prod.startswith('80') or PATRON_MANTENIMIENTO.search(desc_limpia):
        return False

    tiene_clave_activo = clave_prod.startswith(PREFIJOS_ACTIVOS)
    tiene_patron_activo = bool(PATRON_ACTIVOS.search(desc_limpia))

    return tiene_clave_activo or tiene_patron_activo


# ---------------------------------------------------------
# 4.05 FAMILIA DE CATEGORÍA SEGÚN LA CLAVE PRODSERV DEL SAT (FUENTE DE VERDAD)
# ---------------------------------------------------------
# Mapeo de prefijos de Clave de Producto/Servicio (ClaveProdServ) del SAT ->
# la categoría que esa clave oficialmente "promete". Esta es la fuente de
# verdad que decide la categoría de un concepto — YA NO se adivina por
# palabras sueltas en la descripción. La descripción libre solo sirve,
# aparte, para VALIDAR que sea coherente con la clave (ver
# `validar_clave_vs_descripcion` más abajo): correcto / incorrecto.
FAMILIAS_CLAVE_PRODSERV = (
    (PREFIJOS_REFACCIONES, "REFACCION"),
    (PREFIJOS_FLETES, "FLETE"),
    (PREFIJOS_HONORARIOS, "HONORARIO_PROFESIONAL"),
    (PREFIJOS_COMBUSTIBLE, "COMBUSTIBLE"),
    (PREFIJOS_SEGUROS, "SEGUROS"),
    (PREFIJOS_SOFTWARE_TI, "SOFTWARE_TI"),
    (PREFIJOS_PUBLICIDAD_MERCADOTECNIA, "PUBLICIDAD_MERCADOTECNIA"),
    (PREFIJOS_ARRENDAMIENTO, "ARRENDAMIENTO"),
    (PREFIJOS_CAPACITACION, "CAPACITACION"),
    (PREFIJOS_ACTIVOS, "ACTIVO_FIJO"),
)

# Categorías cuya única regla es "si la Clave ProdServ cae en esta familia,
# la respuesta es directa" (a diferencia de ACTIVO_FIJO, que además necesita
# revisar importe/descripción — ver `clasificar_concepto`).
CATEGORIAS_CLAVE_DIRECTA = (
    "REFACCION", "FLETE", "HONORARIO_PROFESIONAL", "COMBUSTIBLE", "SEGUROS",
    "SOFTWARE_TI", "PUBLICIDAD_MERCADOTECNIA", "ARRENDAMIENTO", "CAPACITACION",
)

# Todas las categorías fiscales que puede devolver `clasificar_concepto`,
# usada por app.py para validar la respuesta de la IA y para el
# desplegable de reclasificación manual.
CATEGORIAS_VALIDAS = CATEGORIAS_CLAVE_DIRECTA + ("ACTIVO_FIJO", "GENERAL")


# ---------------------------------------------------------
# 4.06 CATEGORÍAS DINÁMICAS (aprendidas en caliente, aprobadas por Admin)
# ---------------------------------------------------------
# A diferencia de las categorías de arriba (código Python fijo, requieren
# editar este archivo), las categorías dinámicas se guardan en un JSON y se
# pueden crear/ampliar en caliente desde la app — sin tocar código — cuando
# un Admin aprueba una propuesta. Sirven para dos cosas:
#   1) Crear una categoría de negocio COMPLETAMENTE NUEVA (ej. "AGUA") que
#      no existía en ninguna parte de este archivo.
#   2) REFORZAR una categoría ya existente (incluso una de las de arriba,
#      como REFACCION) agregándole palabras o prefijos nuevos sin editar
#      Python — las palabras dinámicas se buscan igual que las fijas.
#
# Antes vivían en categorias_dinamicas.json, junto a este archivo. Ese
# archivo se perdía cada vez que Streamlit Community Cloud redesplegaba o
# reciclaba el contenedor, así que ahora se guardan en Supabase (ver
# bd_supabase.py) — el formato del diccionario en memoria es idéntico:
# {
#   "AGUA": {
#     "palabras": ["AGUA PURIFICADA", "GARRAFON", ...],
#     "prefijos": ["5011"],
#     "uso_cfdi": ["G03"],
#     "descripcion": "Agua purificada y garrafones para oficina/taller",
#     "creado_por": "Administrador AUTOCOM",
#     "fecha_creacion": "2026-09-15 10:00"
#   }
# }
def cargar_categorias_dinamicas():
    """Lee las categorías dinámicas desde Supabase. Si la BD no está
    configurada o falla, se comporta como si no hubiera ninguna (nunca
    rompe la clasificación por reglas fijas)."""
    try:
        return _bd.cargar_categorias_dinamicas()
    except Exception:
        return {}


def guardar_categorias_dinamicas(categorias):
    """Guarda el diccionario completo de categorías dinámicas en Supabase
    (upsert por categoría, no se pierde nada de lo ya aprendido)."""
    _bd.guardar_categorias_dinamicas(categorias)


def agregar_o_reforzar_categoria_dinamica(nombre_categoria, palabras_nuevas=None, prefijos_nuevos=None,
                                          uso_cfdi=None, descripcion=None, creado_por=""):
    """
    Da de alta una categoría dinámica nueva, o si ya existe (dinámica o
    incluso con el mismo nombre que una fija de arriba), le agrega palabras
    y prefijos sin duplicar. Devuelve el diccionario actualizado ya
    guardado en disco.
    """
    nombre_categoria = (nombre_categoria or '').strip().upper().replace(' ', '_')
    categorias = cargar_categorias_dinamicas()
    entrada = categorias.get(nombre_categoria, {
        "palabras": [], "prefijos": [], "uso_cfdi": list(uso_cfdi or ["G03"]),
        "descripcion": descripcion or "", "creado_por": creado_por,
        "fecha_creacion": datetime.now().strftime("%Y-%m-%d %H:%M"),
    })

    palabras_existentes = {p.strip().upper() for p in entrada.get("palabras", [])}
    for palabra in (palabras_nuevas or []):
        palabra = str(palabra).strip().upper()
        if palabra and palabra not in palabras_existentes:
            entrada.setdefault("palabras", []).append(palabra)
            palabras_existentes.add(palabra)

    prefijos_existentes = set(entrada.get("prefijos", []))
    for prefijo in (prefijos_nuevos or []):
        prefijo = str(prefijo).strip()
        if prefijo and prefijo not in prefijos_existentes:
            entrada.setdefault("prefijos", []).append(prefijo)
            prefijos_existentes.add(prefijo)

    if descripcion and not entrada.get("descripcion"):
        entrada["descripcion"] = descripcion
    if uso_cfdi:
        entrada["uso_cfdi"] = sorted(set(entrada.get("uso_cfdi", [])) | set(uso_cfdi))

    categorias[nombre_categoria] = entrada
    guardar_categorias_dinamicas(categorias)
    return categorias


def _patron_desde_palabras(palabras):
    """Compila un regex \\b(palabra1|palabra2|...)\\b a partir de una lista de
    palabras/frases dinámicas. None si la lista viene vacía."""
    palabras_validas = [p for p in (palabras or []) if p and p.strip()]
    if not palabras_validas:
        return None
    return re.compile(r'\b(' + '|'.join(re.escape(p) for p in palabras_validas) + r')\b', re.IGNORECASE)


def obtener_categorias_validas(categorias_dinamicas=None):
    """CATEGORIAS_VALIDAS (fijas) + los nombres de todas las categorías
    dinámicas dadas de alta. Es la lista blanca real que debe usar la IA y
    el validador — se recalcula cada vez porque las dinámicas pueden crecer
    en cualquier momento sin reiniciar la app."""
    dinamicas = tuple((categorias_dinamicas or {}).keys())
    return tuple(dict.fromkeys(CATEGORIAS_VALIDAS + dinamicas))  # sin duplicados, conserva orden



def categoria_por_clave_prodserv(clave_prod, categorias_dinamicas=None):
    """
    Determina la categoría fiscal de un concepto ÚNICAMENTE a partir de su
    Clave de Producto/Servicio (ClaveProdServ) del SAT — el "uso de producto
    y servicio" que el propio comprobante declara. No mira la descripción en
    absoluto.

    Revisa primero las familias FIJAS (código Python) y, si no hay match,
    revisa los prefijos de las categorías DINÁMICAS (aprendidas en caliente,
    ver sección 4.06) en el orden en que se dieron de alta.

    Devuelve el nombre de categoría, o None si la clave no cae en ninguna
    familia registrada (código genérico).
    """
    clave_prod = (clave_prod or '').strip()
    for prefijos, categoria in FAMILIAS_CLAVE_PRODSERV:
        if clave_prod.startswith(prefijos):
            return categoria
    for nombre, datos in (categorias_dinamicas or {}).items():
        prefijos_dinamicos = tuple(datos.get("prefijos", []))
        if prefijos_dinamicos and clave_prod.startswith(prefijos_dinamicos):
            return nombre
    return None


def _clave_tiene_nombre_conocido(clave_prod, catalogo_externo=None):
    """True si la Clave ProdServ tiene un nombre legible reconocido, ya sea
    en el catálogo oficial completo (`catalogo_externo`, cargado desde
    sat_catalogos.db) o en el interno NOMBRES_CLAVES_SAT (completa, por
    familia de 4 dígitos o por división de 2). Se usa para decidir si una
    clave 'genérica' (sin familia de negocio) es al menos algo reconocido, o
    si es totalmente desconocida y debe ir a REVISION."""
    clave_prod = (clave_prod or '').strip()
    if not clave_prod:
        return False
    if catalogo_externo and clave_prod in catalogo_externo:
        return True
    return (
        clave_prod in NOMBRES_CLAVES_SAT
        or clave_prod[:4] in NOMBRES_CLAVES_SAT
        or clave_prod[:2] in NOMBRES_CLAVES_SAT
    )


def categoria_por_descripcion(desc_limpia, importe, categorias_dinamicas=None):
    """
    Segundo respaldo (después de la Clave ProdServ) para no mandar todo lo
    no catalogado a GENERAL/REVISION a ciegas: busca en el texto LIBRE de la
    descripción patrones fuertes de alguna categoría de negocio conocida —
    tanto FIJAS (regex de este archivo) como DINÁMICAS (aprendidas en
    caliente, sección 4.06).

    Solo se usa dentro de `clasificar_concepto` cuando la Clave ProdServ del
    concepto es GENÉRICA (no cae en ninguna familia registrada en
    `FAMILIAS_CLAVE_PRODSERV`) — nunca se usa para "ganarle" a una Clave
    ProdServ que sí cae en una familia reconocida.

    Esto es MOTOR DE REGLAS, no IA: es determinista, auditable y siempre
    corre. La IA nunca hace esta primera asignación — solo entra cuando una
    analista la pide expresamente con el botón "Consultar IA".

    `desc_limpia` debe venir ya sin acentos y en MAYÚSCULAS (ver
    `quitar_acentos`). Devuelve la categoría detectada o None si no hay
    ningún patrón que coincida con suficiente certeza.
    """
    if PATRON_REFACCIONES.search(desc_limpia):
        return "REFACCION"
    if PATRON_HONORARIOS.search(desc_limpia):
        return "HONORARIO_PROFESIONAL"
    if PATRON_COMBUSTIBLE.search(desc_limpia):
        return "COMBUSTIBLE"
    if PATRON_SEGUROS.search(desc_limpia):
        return "SEGUROS"
    if PATRON_SOFTWARE_TI.search(desc_limpia):
        return "SOFTWARE_TI"
    if PATRON_PUBLICIDAD_MERCADOTECNIA.search(desc_limpia):
        return "PUBLICIDAD_MERCADOTECNIA"
    if PATRON_ARRENDAMIENTO.search(desc_limpia):
        return "ARRENDAMIENTO"
    if PATRON_CAPACITACION.search(desc_limpia):
        return "CAPACITACION"
    if PATRON_ACTIVOS.search(desc_limpia):
        # Un "activo" detectado solo por descripción (clave genérica) es más
        # delicado: solo se acepta si el importe ya alcanza el umbral fiscal
        # de activo fijo (UMBRAL_ACTIVO_FIJO = $12,000); si no, se deja para
        # que lo decida la ruta normal de REVISION/GENERAL más abajo, en vez
        # de inventar un ACTIVO_FIJO de bajo importe.
        if importe >= UMBRAL_ACTIVO_FIJO:
            return "ACTIVO_FIJO"
        return None

    for nombre, datos in (categorias_dinamicas or {}).items():
        patron_dinamico = _patron_desde_palabras(datos.get("palabras"))
        if patron_dinamico and patron_dinamico.search(desc_limpia):
            return nombre

    return None


def clasificar_concepto(clave_prod, descripcion, importe, catalogo_externo=None, categorias_dinamicas=None):
    """
    Clasifica UN concepto tomando como FUENTE DE VERDAD la Clave de
    Producto/Servicio (ClaveProdServ) del SAT — el "uso de producto y
    servicio" que declara el propio comprobante — en este orden de
    prioridad:

        REFACCION > FLETE > HONORARIO_PROFESIONAL > COMBUSTIBLE > SEGUROS
            > SOFTWARE_TI > PUBLICIDAD_MERCADOTECNIA > ARRENDAMIENTO
            > CAPACITACION > ACTIVO_FIJO > GENERAL > REVISION

    Todo lo que hace esta función es MOTOR DE REGLAS (determinista y
    auditable). La IA NUNCA participa en esta primera asignación: solo se
    consulta cuando una analista lo pide expresamente desde la interfaz.

    La descripción en texto libre no decide la categoría cuando la Clave
    ProdServ SÍ cae en una familia de negocio conocida. Se usa para:
      1) Confirmar el importe mínimo ($12,000) y descartar mano de obra de
         mantenimiento/reparación ya prestada, cuando la clave SÍ es de
         Activo Fijo.
      2) Servir de respaldo — vía `categoria_por_descripcion` — cuando la
         Clave ProdServ es GENÉRICA (no cae en ninguna familia conocida),
         para NO mandar de entrada a GENERAL/REVISION todo lo que el
         proveedor capturó con una clave genérica pero cuya descripción sí
         deja clara la naturaleza del gasto (nunca gana sobre una clave que
         SÍ está catalogada: eso lo detecta `validar_clave_vs_descripcion`
         como incoherencia a verificar, no como reclasificación silenciosa).

    `catalogo_externo`, si se pasa (dict {clave: nombre} cargado desde la
    base de catálogos oficial del SAT, ver `cargar_catalogo_prodserv` en
    app.py), se usa aquí ÚNICAMENTE para decidir si una clave "genérica" al
    menos tiene un nombre oficial reconocido (ver
    `_clave_tiene_nombre_conocido`) — nunca para inferir la categoría.

    `categorias_dinamicas`, si se pasa (dict cargado con
    `cargar_categorias_dinamicas()`), agrega categorías dadas de alta en
    caliente — nuevas o reforzando una fija — con la MISMA prioridad que
    las de código: primero por Clave ProdServ, y solo si la clave es
    genérica, por descripción (ver sección 4.06).

    Si la Clave ProdServ es genérica Y ADEMÁS la descripción no matchea
    ninguna categoría conocida Y ni siquiera tiene un nombre reconocido, el
    concepto se manda SIEMPRE a REVISION — nunca se deja pasar como GENERAL
    solo porque el importe sea bajo, porque en ese caso ni siquiera sabemos
    con certeza qué es.

    Si quieres saber si la DESCRIPCIÓN es o no coherente con la Clave
    ProdServ ya clasificada (correcto/incorrecto), usa
    `validar_clave_vs_descripcion()` — esa es la comparación explícita
    "clave del SAT vs. concepto".
    """
    clave_prod = (clave_prod or '').strip()
    desc_limpia = quitar_acentos(descripcion or '').upper()
    categoria_clave = categoria_por_clave_prodserv(clave_prod, categorias_dinamicas)

    if categoria_clave in CATEGORIAS_CLAVE_DIRECTA:
        return categoria_clave

    if categoria_clave == "ACTIVO_FIJO":
        if PATRON_MANTENIMIENTO.search(desc_limpia):
            # La clave es de Activo Fijo, pero la descripción es claramente
            # mano de obra de mantenimiento/reparación ya prestada: es un
            # gasto general, no la compra/alta de un activo.
            return "GENERAL"
        if importe >= UMBRAL_ACTIVO_FIJO:
            return "ACTIVO_FIJO"
        # Clave de Activo Fijo pero importe por debajo del umbral fiscal de
        # $12,000: no amerita el tratamiento de Activo Fijo. Si el importe
        # no es trivial, se manda a revisión para que un analista confirme.
        return "REVISION" if importe >= UMBRAL_AMBIGUEDAD_IA else "GENERAL"

    # Si la clave cayó en una categoría DINÁMICA (por prefijo), esa ya es la
    # respuesta directa — igual que las fijas de CATEGORIAS_CLAVE_DIRECTA.
    if categoria_clave and categorias_dinamicas and categoria_clave in categorias_dinamicas:
        return categoria_clave

    # Clave ProdServ genérica: no cae en ninguna familia de negocio conocida.
    if PATRON_MANTENIMIENTO.search(desc_limpia):
        return "GENERAL"

    categoria_texto = categoria_por_descripcion(desc_limpia, importe, categorias_dinamicas)
    if categoria_texto:
        return categoria_texto

    # Una clave que ni siquiera tiene un nombre reconocido (ni en el
    # catálogo oficial externo ni en el interno NOMBRES_CLAVES_SAT) es, casi
    # por definición, un caso para revisar manualmente — no debe colarse
    # como GENERAL solo porque el importe sea bajo: no sabemos qué es.
    if not _clave_tiene_nombre_conocido(clave_prod, catalogo_externo):
        return "REVISION"

    if importe >= UMBRAL_AMBIGUEDAD_IA:
        return "REVISION"

    return "GENERAL"


# ---------------------------------------------------------
# 4.5 COHERENCIA CONCEPTO <-> OBJETO DE IMPUESTO (ObjetoImp)
# ---------------------------------------------------------
# Catálogo oficial SAT c_ObjetoImp (CFDI 4.0), a nivel Concepto.
CLAVES_OBJETO_IMP = {
    "01": "No objeto de impuesto",
    "02": "Sí objeto de impuesto",
    "03": "Sí objeto de impuesto y no obligado al desglose",
    "04": "Sí objeto de impuesto y no causa impuesto (ej. exportación definitiva)",
}

# Categorías que, por su naturaleza (venta de refacciones, activos, fletes,
# honorarios), casi siempre causan IVA en operaciones nacionales normales.
# Que vengan marcadas como "no objeto de impuesto" es una señal atípica que
# vale la pena mandar a revisión manual, no un rechazo automático (puede
# haber casos legítimos: exportación, donativos, etc.).
CATEGORIAS_NORMALMENTE_GRAVADAS = {"REFACCION", "ACTIVO_FIJO", "HONORARIO_PROFESIONAL", "FLETE"}


def validar_objeto_impuesto(categoria, objeto_imp, tiene_traslados):
    """
    Revisa que el ObjetoImp declarado en el concepto sea coherente con:
      (a) si el propio concepto trae o no un nodo Impuestos/Traslados, y
      (b) si el tipo de bien/servicio detectado por el motor de reglas es
          de los que normalmente causan IVA.

    Devuelve (es_coherente: bool, motivo: str | None).
    NO decide si la factura se rechaza — solo describe la inconsistencia;
    quien llama (app.py) decide si eso implica "NO cumple" o "REVISAR".
    """
    objeto_imp = (objeto_imp or "").strip()

    if objeto_imp not in CLAVES_OBJETO_IMP:
        return False, f"ObjetoImp del concepto ('{objeto_imp or 'vacío'}') no es un valor válido del catálogo SAT."

    if objeto_imp == "02" and not tiene_traslados:
        return False, (
            "ObjetoImp=02 (sí objeto de impuesto) pero el concepto no trae "
            "impuestos trasladados desglosados."
        )

    if objeto_imp == "01" and tiene_traslados:
        return False, (
            "ObjetoImp=01 (no objeto de impuesto) pero el concepto sí trae "
            "impuestos trasladados — contradicción dentro del propio XML."
        )

    if objeto_imp == "01" and categoria in CATEGORIAS_NORMALMENTE_GRAVADAS:
        return False, (
            f"Concepto clasificado como {categoria.replace('_', ' ').title()}, que normalmente "
            "causa IVA, pero el CFDI lo marca como 'No objeto de impuesto' (ObjetoImp=01). "
            "Confirmar con el proveedor si es correcto."
        )

    return True, None


# ---------------------------------------------------------
# 4.6 COHERENCIA: CLAVE PRODSERV DEL SAT  <->  CONCEPTO (DESCRIPCIÓN)
# ---------------------------------------------------------
# Esta es la comparación explícita "Clave ProdServ vs. Concepto: correcto o
# incorrecto" — usa `categoria_por_clave_prodserv()` (definida arriba, la
# MISMA fuente de verdad que usa `clasificar_concepto`) y solo contrasta esa
# categoría oficial contra lo que la descripción libre sugiere.
def validar_clave_vs_descripcion(clave_prod, descripcion):
    """
    Compara la categoría que la Clave de Producto/Servicio (ProdServ) del
    SAT "promete" (según `categoria_por_clave_prodserv`) contra lo que la
    propia descripción en texto libre del concepto sugiere, para detectar
    capturas donde la Clave ProdServ no corresponde a lo que realmente se
    está facturando (ej. una Clave de Refacción para lo que la descripción
    dice que es un Activo Fijo, o una Clave de Activo para lo que en
    realidad es un flete).

    Devuelve (es_coherente: bool, motivo: str | None). NO decide el "Cumple"
    de la línea — solo dice si la Clave ProdServ y el Concepto son
    coherentes o no; quien llama decide si una incoherencia implica mandar
    el concepto a REVISION.
    """
    categoria_clave = categoria_por_clave_prodserv(clave_prod)
    if categoria_clave is None:
        return True, None  # clave genérica/no registrada: nada oficial que contrastar

    desc_limpia = quitar_acentos(descripcion or '').upper()

    categorias_por_texto = set()
    if PATRON_REFACCIONES.search(desc_limpia):
        categorias_por_texto.add("REFACCION")
    if PATRON_ACTIVOS.search(desc_limpia):
        categorias_por_texto.add("ACTIVO_FIJO")
    if PATRON_HONORARIOS.search(desc_limpia):
        categorias_por_texto.add("HONORARIO_PROFESIONAL")
    if 'FLETE' in desc_limpia or 'AUTOTRANSPORTE' in desc_limpia:
        categorias_por_texto.add("FLETE")
    if PATRON_COMBUSTIBLE.search(desc_limpia):
        categorias_por_texto.add("COMBUSTIBLE")
    if PATRON_SEGUROS.search(desc_limpia):
        categorias_por_texto.add("SEGUROS")
    if PATRON_SOFTWARE_TI.search(desc_limpia):
        categorias_por_texto.add("SOFTWARE_TI")
    if PATRON_PUBLICIDAD_MERCADOTECNIA.search(desc_limpia):
        categorias_por_texto.add("PUBLICIDAD_MERCADOTECNIA")
    if PATRON_ARRENDAMIENTO.search(desc_limpia):
        categorias_por_texto.add("ARRENDAMIENTO")
    if PATRON_CAPACITACION.search(desc_limpia):
        categorias_por_texto.add("CAPACITACION")

    if not categorias_por_texto or categoria_clave in categorias_por_texto:
        # Sin palabras clave fuertes en el texto (nada que contrastar), o la
        # descripción sí coincide con lo que promete la clave: todo bien.
        return True, None

    return False, (
        f"La Clave ProdServ SAT declarada corresponde a {categoria_clave.replace('_', ' ').title()}, pero la "
        f"descripción del concepto corresponde a {'/'.join(c.replace('_', ' ').title() for c in sorted(categorias_por_texto))}. "
        "Verificar con el proveedor si la Clave ProdServ está correctamente capturada."
    )


# ---------------------------------------------------------
# 4.7 CATÁLOGO SAT c_UsoCFDI Y COHERENCIA CONCEPTO <-> USO CFDI
# ---------------------------------------------------------
CATALOGO_USO_CFDI = {
    "G01": "Adquisición de mercancías",
    "G02": "Devoluciones, descuentos o bonificaciones",
    "G03": "Gastos en general",
    "I01": "Construcciones",
    "I02": "Mobiliario y equipo de oficina por inversiones",
    "I03": "Equipo de transporte",
    "I04": "Equipo de computo y accesorios",
    "I05": "Dados, troqueles, moldes, matrices y otros activos",
    "I06": "Comunicaciones telefónicas",
    "I07": "Comunicaciones satelitales",
    "I08": "Otra maquinaria y equipo",
    "D01": "Honorarios médicos, dentales y gastos hospitalarios",
    "D10": "Pagos por servicios educativos (colegiaturas)",
    "P01": "Por definir",
    "S01": "Sin efectos fiscales",
    "CP01": "Pagos",
    "CN01": "Nómina",
}

# Para cada categoría detectada por el motor de reglas, qué Usos CFDI del
# catálogo c_UsoCFDI son coherentes con ella. No pretende cubrir TODO el
# catálogo (el Uso CFDI válido también depende del régimen fiscal del
# receptor) — es una señal de foco rojo automotriz: p.ej. un ACTIVO_FIJO
# (equipo de taller/cómputo) casi nunca debe ir declarado con G01
# (mercancías para reventa).
USOS_CFDI_COHERENTES_POR_CATEGORIA = {
    "ACTIVO_FIJO": ("I01", "I02", "I03", "I04", "I05", "I06", "I07", "I08"),
    "REFACCION": ("G01", "G03"),
    "FLETE": ("G03",),
    "HONORARIO_PROFESIONAL": ("G03", "D01", "D10"),
    "COMBUSTIBLE": ("G03",),
    "SEGUROS": ("G03",),
    "SOFTWARE_TI": ("G03", "I02", "I04"),
    "PUBLICIDAD_MERCADOTECNIA": ("G03",),
    "ARRENDAMIENTO": ("G03",),
    "CAPACITACION": ("G03",),
}


# ---------------------------------------------------------
# 4.75 CATÁLOGO SAT c_FormaPago Y COHERENCIA CON MetodoPago
# ---------------------------------------------------------
CATALOGO_FORMA_PAGO = {
    "01": "Efectivo",
    "02": "Cheque nominativo",
    "03": "Transferencia electrónica de fondos",
    "04": "Tarjeta de crédito",
    "05": "Monedero electrónico",
    "06": "Dinero electrónico",
    "08": "Vales de despensa",
    "12": "Dación en pago",
    "13": "Pago por subrogación",
    "14": "Pago por consignación",
    "15": "Condonación",
    "17": "Compensación",
    "23": "Novación",
    "24": "Confusión",
    "25": "Remisión de deuda",
    "26": "Prescripción o caducidad",
    "27": "A satisfacción del acreedor",
    "28": "Tarjeta de débito",
    "29": "Tarjeta de servicios",
    "30": "Aplicación de anticipos",
    "31": "Intermediario pagos",
    "99": "Por definir",
}


def obtener_nombre_forma_pago(forma_pago):
    """'03 - Transferencia electrónica de fondos', o la clave cruda si no
    está en el catálogo oficial (para no ocultar datos capturados raros)."""
    forma_pago = (forma_pago or '').strip()
    if not forma_pago:
        return "N/A"
    nombre = CATALOGO_FORMA_PAGO.get(forma_pago)
    if nombre:
        return f"{forma_pago} - {nombre}"
    return f"{forma_pago} - [NO ES UN VALOR VÁLIDO DEL CATÁLOGO SAT c_FormaPago]"


def validar_forma_pago_vs_metodo_pago(metodo_pago, forma_pago):
    """
    Valida la coherencia MetodoPago <-> FormaPago según la "Guía de llenado
    de comprobantes fiscales digitales por Internet" del SAT (CFDI 4.0):

      - PPD (Pago en parcialidades o diferido) SIEMPRE debe declarar
        FormaPago = "99" (Por definir), porque a la fecha de emisión aún no
        se conoce cómo se pagará; la forma real se reporta después con el
        Complemento de Recepción de Pagos (REP, UsoCFDI "CP01").
      - PUE (Pago en una sola exhibición) NUNCA debe declarar "99": a esa
        fecha ya se conoce y debe declararse la forma de pago real
        utilizada (efectivo, transferencia, tarjeta, etc.).
      - Cualquier FormaPago que no exista en el catálogo oficial
        c_FormaPago es, por definición, un error de captura del emisor.

    Devuelve (es_coherente: bool, motivo: str | None). NO decide si esto
    bloquea la factura — quien llama (app.py) decide.
    """
    metodo_pago = (metodo_pago or '').strip().upper()
    forma_pago = (forma_pago or '').strip()

    if forma_pago and forma_pago not in CATALOGO_FORMA_PAGO:
        return False, (
            f"FormaPago '{forma_pago}' declarado en el CFDI no es un valor válido "
            "del catálogo oficial SAT c_FormaPago."
        )

    if metodo_pago == 'PPD' and forma_pago != '99':
        return False, (
            f"MetodoPago=PPD requiere FormaPago='99 - Por definir', pero el CFDI declara "
            f"'{obtener_nombre_forma_pago(forma_pago)}'. La forma de pago real se reporta "
            "después vía Complemento de Pago (REP)."
        )

    if metodo_pago == 'PUE' and forma_pago == '99':
        return False, (
            "MetodoPago=PUE no puede declarar FormaPago='99 - Por definir'; a esta fecha "
            "ya se conoce y debe declararse la forma de pago real utilizada."
        )

    return True, None


def validar_concepto_vs_uso_cfdi(categoria, uso_cfdi, categorias_dinamicas=None):
    """
    Compara la categoría que el motor de reglas detectó para el concepto
    (por su Clave ProdServ y su descripción) contra el Uso CFDI que el
    RECEPTOR declaró para la factura completa, usando como referencia el
    catálogo oficial c_UsoCFDI del SAT. Es la validación "Activo Fijo
    requiere I01-I08" llevada un paso más allá: también cubre Refacción,
    Flete y Honorario Profesional contra los usos que normalmente les
    corresponden, además de cualquier categoría dinámica dada de alta con
    su propio "uso_cfdi" esperado.

    Devuelve (es_coherente: bool, motivo: str | None). NO decide "Cumple";
    quien llama decide si la incoherencia implica REVISION o rechazo.
    """
    uso_cfdi = (uso_cfdi or '').strip().upper()
    usos_esperados = USOS_CFDI_COHERENTES_POR_CATEGORIA.get(categoria)
    if not usos_esperados and categorias_dinamicas and categoria in categorias_dinamicas:
        usos_esperados = tuple(categorias_dinamicas[categoria].get("uso_cfdi", []))
    if not usos_esperados or not uso_cfdi or uso_cfdi == "N/A":
        return True, None

    if uso_cfdi in usos_esperados:
        return True, None

    nombre_uso = CATALOGO_USO_CFDI.get(uso_cfdi, uso_cfdi)
    return False, (
        f"Concepto clasificado como {categoria.replace('_', ' ').title()}, pero el Uso CFDI declarado "
        f"es '{uso_cfdi} - {nombre_uso}', que no corresponde al catálogo SAT esperado para esa categoría "
        f"({'/'.join(usos_esperados)}). Confirmar con el receptor cuál es el Uso CFDI correcto."
    )


# ---------------------------------------------------------
# 4.78 PARÁMETROS FISCALES DE RETENCIONES ISR / IVA (FILTRO PRINCIPAL)
# ---------------------------------------------------------
# Todas las tasas de retención que AUTOCOM valida, centralizadas aquí para
# que sean auditables y editables desde la pantalla "Enseñar reglas" en vez
# de estar escritas a mano dentro de app.py.
#
# Claves de impuesto del catálogo SAT c_Impuesto:
#   001 = ISR, 002 = IVA, 003 = IEPS
CLAVE_IMPUESTO_ISR = '001'
CLAVE_IMPUESTO_IVA = '002'
CLAVE_IMPUESTO_IEPS = '003'

NOMBRES_IMPUESTO = {'001': 'ISR', '002': 'IVA', '003': 'IEPS'}

# Tasas de retención vigentes (LISR / LIVA / RLIVA). Se expresan como
# fracción decimal, igual que el atributo TasaOCuota del XML.
TASA_RET_ISR_RESICO_PF = 0.0125       # RESICO Persona Física (art. 113-J LISR)
TASA_RET_ISR_SERVICIOS_PF = 0.10      # Servicios profesionales PF (art. 106 LISR)
TASA_RET_ISR_ARRENDAMIENTO_PF = 0.10  # Arrendamiento PF (art. 116 LISR)
TASA_RET_IVA_SERVICIOS_PF = 0.106667  # 2/3 partes del IVA (art. 1-A LIVA / 3 RLIVA)
TASA_RET_IVA_ARRENDAMIENTO_PF = 0.106667
TASA_RET_IVA_FLETES = 0.04            # Autotransporte terrestre de carga (art. 1-A LIVA)

# Tolerancia al comparar tasas: el SAT permite que el emisor exprese
# 10.6667% como 0.106667, 0.1067 o 0.106666, y los redondeos de centavos
# mueven ligeramente el cálculo cuando la tasa se deriva del importe.
TOLERANCIA_TASA_RETENCION = 0.0015

# Regímenes fiscales del catálogo SAT c_RegimenFiscal relevantes para
# retenciones. Solo se listan los que disparan alguna regla.
REGIMEN_ARRENDAMIENTO_PF = '606'
REGIMEN_ACTIVIDADES_PROFESIONALES_PF = '612'

# El código 626 (RESICO) es COMPARTIDO por Persona Física y Persona Moral:
# no son dos regímenes distintos, es la misma clave con dos mecánicas de
# cálculo de ISR completamente distintas.
#   - RESICO Persona Física: paga ISR con tasas de 1%-2.5% sobre lo cobrado,
#     y por eso existe la retención de 1.25% del art. 113-J LISR cuando
#     quien le paga es una Persona Moral (es la regla de abajo).
#   - RESICO Persona Moral: calcula su ISR con la tasa corporativa (30%)
#     sobre utilidad de flujo de efectivo — es SU cálculo interno, no algo
#     que AUTOCOM le retenga al pagarle. Para lo que a AUTOCOM le toca
#     retener, un proveedor RESICO Moral se trata IGUAL que cualquier otra
#     Persona Moral: sin retención especial (salvo que aplique otra regla
#     independiente del régimen, como fletes).
# Por eso la constante de abajo solo importa para la rama de Persona
# Física: no hay una constante "REGIMEN_RESICO_PM" porque no dispara
# ninguna retención propia.
REGIMEN_RESICO = '626'
REGIMEN_RESICO_PF = REGIMEN_RESICO  # alias por compatibilidad con código existente

NOMBRES_REGIMEN_FISCAL = {
    '601': 'General de Ley Personas Morales',
    '603': 'Personas Morales con Fines no Lucrativos',
    '605': 'Sueldos y Salarios e Ingresos Asimilados a Salarios',
    '606': 'Arrendamiento',
    '607': 'Régimen de Enajenación o Adquisición de Bienes',
    '608': 'Demás ingresos',
    '610': 'Residentes en el Extranjero sin Establecimiento Permanente en México',
    '611': 'Ingresos por Dividendos (socios y accionistas)',
    '612': 'Personas Físicas con Actividades Empresariales y Profesionales',
    '614': 'Ingresos por intereses',
    '615': 'Régimen de los ingresos por obtención de premios',
    '616': 'Sin obligaciones fiscales',
    '620': 'Sociedades Cooperativas de Producción que optan por Diferir sus Ingresos',
    '621': 'Incorporación Fiscal',
    '622': 'Actividades Agrícolas, Ganaderas, Silvícolas y Pesqueras',
    '623': 'Opcional para Grupos de Sociedades',
    '624': 'Coordinados',
    '625': 'Régimen de las Actividades Empresariales con ingresos a través de Plataformas Tecnológicas',
    '626': 'Régimen Simplificado de Confianza (RESICO)',
}

# Categorías de concepto que, por su naturaleza, obligan a retención de IVA
# del 4% por tratarse de autotransporte terrestre de carga.
CATEGORIAS_RETENCION_FLETE = ("FLETE",)

# Categorías que se consideran "servicio profesional" para efectos de la
# retención 10% ISR + 10.6667% IVA cuando el emisor es Persona Física en
# régimen 612. Las categorías nuevas (SOFTWARE_TI, PUBLICIDAD, CAPACITACION)
# entran aquí porque cuando las factura una PF del 612 son igualmente
# servicios profesionales sujetos a retención.
CATEGORIAS_SERVICIO_PROFESIONAL_PF = (
    "HONORARIO_PROFESIONAL", "SOFTWARE_TI", "PUBLICIDAD_MERCADOTECNIA", "CAPACITACION",
)


def obtener_nombre_regimen(regimen):
    """'612 - Personas Físicas con Actividades Empresariales y Profesionales'."""
    regimen = (regimen or '').strip()
    if not regimen:
        return "N/A"
    nombre = NOMBRES_REGIMEN_FISCAL.get(regimen)
    return f"{regimen} - {nombre}" if nombre else f"{regimen} - [régimen no reconocido]"


def _tasa_coincide(tasa_declarada, tasa_esperada):
    """Compara dos tasas con la tolerancia fiscal de TOLERANCIA_TASA_RETENCION."""
    return abs(float(tasa_declarada or 0.0) - float(tasa_esperada)) <= TOLERANCIA_TASA_RETENCION


def validar_retenciones(retenciones_xml, regimen_emisor, es_persona_fisica,
                        categorias_factura=(), es_flete=False):
    """
    FILTRO PRINCIPAL de aceptación/rechazo por retenciones ISR e IVA.

    Recibe:
      - `retenciones_xml`: dict {clave_impuesto: tasa} tal como viene del
        nodo Impuestos/Retenciones del CFDI (ej. {'001': 0.10, '002': 0.1067}).
      - `regimen_emisor`: RegimenFiscal del nodo Emisor (c_RegimenFiscal).
      - `es_persona_fisica`: True si el RFC del emisor tiene 13 posiciones.
      - `categorias_factura`: categorías de los conceptos de la factura, para
        saber si hay servicios profesionales, arrendamiento, fletes, etc.
      - `es_flete`: True si algún concepto quedó clasificado como FLETE.

    Devuelve (todas_correctas: bool, observaciones: list[str]).

    Reglas aplicadas (todas sobre emisor PERSONA FÍSICA; a una Persona Moral
    no se le retiene ISR/IVA en estos supuestos):
      1) RESICO Persona Física (régimen 626): retención de 1.25% de ISR.
         El régimen 626 también existe para Persona Moral (RESICO PM), pero
         esa es una mecánica de cálculo de ISR distinta (tasa corporativa
         sobre flujo de efectivo) que NO genera ninguna retención a cargo
         de AUTOCOM: un proveedor RESICO Moral se trata igual que cualquier
         otra Persona Moral en las reglas de abajo.
      2) Arrendamiento PF (606): 10% ISR + 10.6667% IVA.
      3) Servicios profesionales PF (612) cuando la factura trae conceptos
         de servicio: 10% ISR + 10.6667% IVA.
      4) Fletes / autotransporte de carga (cualquier emisor, PF o PM): 4% de IVA.
      5) Retención declarada donde NO debía haberla (ej. Persona Moral con
         retención de ISR sin ser fletes): se marca para revisar.

    NOTA: esta función devuelve observaciones; quien decide el estatus final
    de la factura es app.py, que las trata como bloqueo de "RECHAZADO DE
    ENTRADA" por ser retenciones fuera de norma.
    """
    observaciones = []
    retenciones_xml = retenciones_xml or {}
    categorias_factura = tuple(categorias_factura or ())

    tasa_isr = retenciones_xml.get(CLAVE_IMPUESTO_ISR, 0.0)
    tasa_iva = retenciones_xml.get(CLAVE_IMPUESTO_IVA, 0.0)
    regimen_emisor = (regimen_emisor or '').strip()

    hay_servicio_profesional = any(c in CATEGORIAS_SERVICIO_PROFESIONAL_PF for c in categorias_factura)
    hay_arrendamiento = "ARRENDAMIENTO" in categorias_factura

    # --- Regla 1: RESICO Persona Física (626) -> 1.25% ISR ---
    # Si el emisor es Persona Moral y también trae régimen 626 (RESICO
    # Moral), esta condición no entra por el "es_persona_fisica": cae de
    # frente a la Regla 5 de abajo, que es exactamente el trato correcto
    # (igual que cualquier otra Persona Moral, sin retención especial).
    if es_persona_fisica and regimen_emisor == REGIMEN_RESICO:
        if not _tasa_coincide(tasa_isr, TASA_RET_ISR_RESICO_PF):
            observaciones.append(
                f"Emisor RESICO Persona Física ({obtener_nombre_regimen(regimen_emisor)}) requiere "
                f"retención de {TASA_RET_ISR_RESICO_PF:.4%} de ISR; el CFDI declara {tasa_isr:.4%}."
            )

    # --- Regla 2: Arrendamiento Persona Física (606) -> 10% ISR + 10.6667% IVA ---
    elif es_persona_fisica and (regimen_emisor == REGIMEN_ARRENDAMIENTO_PF or hay_arrendamiento):
        if not _tasa_coincide(tasa_isr, TASA_RET_ISR_ARRENDAMIENTO_PF):
            observaciones.append(
                f"Arrendamiento de Persona Física requiere retención de "
                f"{TASA_RET_ISR_ARRENDAMIENTO_PF:.2%} de ISR; el CFDI declara {tasa_isr:.4%}."
            )
        if not _tasa_coincide(tasa_iva, TASA_RET_IVA_ARRENDAMIENTO_PF):
            observaciones.append(
                f"Arrendamiento de Persona Física requiere retención de "
                f"{TASA_RET_IVA_ARRENDAMIENTO_PF:.4%} de IVA (2/3 partes); el CFDI declara {tasa_iva:.4%}."
            )

    # --- Regla 3: Servicios profesionales Persona Física (612) -> 10% ISR + 10.6667% IVA ---
    elif es_persona_fisica and regimen_emisor == REGIMEN_ACTIVIDADES_PROFESIONALES_PF and hay_servicio_profesional:
        if not _tasa_coincide(tasa_isr, TASA_RET_ISR_SERVICIOS_PF):
            observaciones.append(
                f"Servicios profesionales de Persona Física requieren retención de "
                f"{TASA_RET_ISR_SERVICIOS_PF:.2%} de ISR; el CFDI declara {tasa_isr:.4%}."
            )
        if not _tasa_coincide(tasa_iva, TASA_RET_IVA_SERVICIOS_PF):
            observaciones.append(
                f"Servicios profesionales de Persona Física requieren retención de "
                f"{TASA_RET_IVA_SERVICIOS_PF:.4%} de IVA (2/3 partes); el CFDI declara {tasa_iva:.4%}."
            )

    # --- Regla 4: Fletes / autotransporte terrestre de carga -> 4% IVA ---
    if es_flete:
        if not _tasa_coincide(tasa_iva, TASA_RET_IVA_FLETES):
            observaciones.append(
                f"Fletes y autotransporte terrestre de carga requieren retención de "
                f"{TASA_RET_IVA_FLETES:.2%} de IVA; el CFDI declara {tasa_iva:.4%}."
            )

    # --- Regla 5: retención declarada donde no correspondía ---
    # Una Persona Moral que no factura fletes no debería traer retenciones de
    # ISR/IVA en estos supuestos; si las trae, hay que verificar antes de
    # pagar. Esto incluye a un proveedor RESICO Moral (régimen 626): su
    # mecánica de cálculo de ISR es interna y no genera retención a cargo de
    # AUTOCOM, así que se le exige lo mismo que a cualquier otra Persona
    # Moral.
    if not es_persona_fisica and not es_flete:
        if tasa_isr > 0:
            observaciones.append(
                f"El emisor es Persona Moral y declara retención de ISR ({tasa_isr:.4%}) "
                "sin un supuesto que la justifique; verificar antes de programar el pago."
            )
        if tasa_iva > 0:
            observaciones.append(
                f"El emisor es Persona Moral y declara retención de IVA ({tasa_iva:.4%}) "
                "sin un supuesto que la justifique; verificar antes de programar el pago."
            )

    return (len(observaciones) == 0), observaciones


# ---------------------------------------------------------
# 4.8 REGLA DE MAYORÍA: % DE CONCEPTOS CORRECTOS ACEPTA TODA LA FACTURA
# ---------------------------------------------------------
def evaluar_regla_mayoria(total_conceptos, conceptos_correctos, umbral=UMBRAL_PORCENTAJE_CONCEPTOS_CORRECTOS):
    """
    Decide si una factura se acepta por "regla de mayoría": si el
    porcentaje de conceptos con Cumple == 'SI' alcanza el umbral (80% por
    defecto), la factura completa se considera correcta aunque queden una o
    más líneas en REVISION / incoherentes.

    Devuelve (cumple_mayoria: bool, porcentaje: float). Con 0 conceptos
    (factura sin líneas de detalle) se considera 100% por definición, para
    no bloquear por división entre cero — ese caso ya se maneja aparte como
    anomalía estructural del CFDI en quien llama.

    Esta regla es exclusivamente sobre la CORRECCIÓN DE LOS CONCEPTOS
    (clasificación, coherencia Clave ProdServ/Uso CFDI). NUNCA debe usarse
    para pasar por alto un bloqueo crítico a nivel factura (69-B, CFDI
    cancelado, retenciones fuera de norma, corte de PUE) — eso lo controla
    quien llama, evaluando esta regla solo cuando no hay ningún bloqueo
    crítico de por medio.
    """
    if total_conceptos <= 0:
        return True, 1.0
    porcentaje = conceptos_correctos / total_conceptos
    return porcentaje >= umbral, porcentaje


def obtener_ultimo_jueves(anio, mes):
    if mes == 12:
        siguiente_mes = datetime(anio + 1, 1, 1)
    else:
        siguiente_mes = datetime(anio, mes + 1, 1)
        
    ultimo_dia = siguiente_mes - pd.Timedelta(days=1)
    dias_a_restar = (ultimo_dia.weekday() - 3) % 7
    return (ultimo_dia - pd.Timedelta(days=dias_a_restar)).day


# ---------------------------------------------------------
# 5. LISTA NEGRA DEL SAT — EFOS (ARTÍCULO 69-B CFF)
# ---------------------------------------------------------
# El SAT publica, como dato abierto, el listado completo de contribuyentes
# con operaciones presuntamente inexistentes (EFOS) en formato CSV. Este
# archivo se actualiza de forma irregular (varias veces al mes), por lo que
# app.py lo descarga con caché (recomendado: refrescar cada 12-24 h, nunca
# en cada factura) y le pasa aquí el diccionario ya cargado en memoria.
#
# Fuente oficial:
#   http://omawww.sat.gob.mx/cifras_sat/Paginas/datos/vinculo.html?page=ListCompleta69B.html
URL_LISTADO_69B = "http://omawww.sat.gob.mx/cifras_sat/Documents/Listado_Completo_69-B.csv"

# El listado del SAT clasifica a cada contribuyente en una de estas
# situaciones. Solo "PRESUNTO" y "DEFINITIVO" deben bloquear la factura de
# entrada: "DESVIRTUADO" (el contribuyente demostró ante el SAT que sus
# operaciones sí existieron) y "SENTENCIA FAVORABLE" (ganó un medio de
# defensa) significan que ya limpió su situación y NO deben rechazarse.
SITUACIONES_69B_BLOQUEAN = {"DEFINITIVO", "PRESUNTO"}


def descargar_listado_69b(timeout=30):
    """
    Descarga el listado público y oficial de EFOS (Art. 69-B CFF) directo
    del portal del SAT y regresa un dict {RFC: SITUACION}.

    Esta función NUNCA debe llamarse una vez por factura: el archivo pesa
    varios MB y el SAT no lo actualiza en tiempo real. app.py la envuelve
    en un cache con vigencia de varias horas.

    Si la descarga o el parseo fallan, se propaga la excepción tal cual —
    quien llama decide qué hacer (reusar la última copia en caché, avisar
    al equipo, etc.). Esta función nunca regresa un listado vacío de forma
    silenciosa ante un error, para no dar una falsa sensación de "sin
    coincidencias en la lista negra".
    """
    resp = requests.get(
        URL_LISTADO_69B,
        timeout=timeout,
        headers={"User-Agent": "Mozilla/5.0 (AuditorCFDI-AUTOCOM)"},
    )
    resp.raise_for_status()

    # El SAT publica este CSV en codificación Windows-1252 y con 3 líneas de
    # título/encabezado antes de la fila con los nombres de columna real.
    texto = resp.content.decode('windows-1252', errors='ignore')
    lineas = texto.splitlines()
    lector = csv.reader(lineas[3:], delimiter=',', quotechar='"')

    listado = {}
    for fila in lector:
        # Columnas del CSV oficial: No. | RFC | Nombre del Contribuyente |
        # Situación del contribuyente | ... (oficios, fechas de publicación, etc.)
        if len(fila) < 4:
            continue
        rfc = (fila[1] or '').strip().upper()
        situacion = (fila[3] or '').strip().upper()
        if rfc:
            listado[rfc] = situacion
    return listado


def verificar_efos(rfc_emisor, listado_69b):
    """
    Consulta el RFC del emisor contra el listado 69-B ya descargado (el
    dict {RFC: SITUACION} que regresa `descargar_listado_69b`).

    Devuelve una tupla (en_lista_negra: bool, situacion: str).
    """
    if not rfc_emisor or not listado_69b:
        return False, ""
    situacion = listado_69b.get(rfc_emisor.strip().upper(), "")
    return situacion in SITUACIONES_69B_BLOQUEAN, situacion


# ---------------------------------------------------------
# 6. VALIDACIÓN EN TIEMPO REAL DE ESTATUS CFDI (VIGENTE / CANCELADO)
# ---------------------------------------------------------
# Consulta el Web Service PÚBLICO y OFICIAL del SAT — el mismo que usa el
# lector de código QR de una representación impresa. No requiere CSD ni
# ningún tipo de autenticación/API key.
#
# Documentación oficial: "Documentación del Servicio de Consulta de CFDI"
#   https://www.sat.gob.mx/minisitio/Factura/documentos/cancelacion/consulta_cfdi.pdf
URL_CONSULTA_CFDI_SAT = "https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc"
SOAP_ACTION_CONSULTA_CFDI = "http://tempuri.org/IConsultaCFDIService/Consulta"


def _formatear_total_para_qr(total):
    """
    El parámetro `tt` de la expresión impresa exige el total con 10 dígitos
    enteros (rellenos con ceros a la izquierda) y 6 decimales, tal como
    aparece en el código QR de una factura impresa
    (ej. 345.30 -> '0000000345.300000').
    """
    entero, decimal = f"{float(total):.6f}".split(".")
    return f"{entero.zfill(10)}.{decimal}"


def construir_expresion_impresa(rfc_emisor, rfc_receptor, total, uuid, sello=""):
    """Construye el querystring '?re=...&rr=...&tt=...&id=...&fe=...' del QR del CFDI."""
    expresion = (
        f"?re={rfc_emisor}&rr={rfc_receptor}"
        f"&tt={_formatear_total_para_qr(total)}&id={uuid}"
    )
    if sello:
        # 'fe' son los últimos 8 caracteres del Sello Digital del comprobante.
        expresion += f"&fe={sello[-8:]}"
    return expresion


def consultar_estatus_cfdi_sat(rfc_emisor, rfc_receptor, total, uuid, sello="", timeout=15):
    """
    Consulta en tiempo real, contra el Web Service oficial del SAT, el
    estatus vigente de UN comprobante.

    Devuelve un dict:
        {
            "estado": "Vigente" | "Cancelado" | "No Encontrado" | "Error de Consulta",
            "es_cancelable": str,
            "estatus_cancelacion": str,
            "validacion_efos_sat": "100" | "200" | "",  # el propio WS valida 69-B
            "detalle": str,
        }

    Esta función NUNCA lanza una excepción ni asume "Vigente" por default:
    ante cualquier falla de red, timeout o respuesta inesperada regresa
    estado "Error de Consulta" para que el analista decida cómo proceder
    (una factura nunca debe darse por buena solo porque no se pudo verificar).
    """
    expresion = construir_expresion_impresa(rfc_emisor, rfc_receptor, total, uuid, sello)
    sobre_soap = (
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:tem="http://tempuri.org/">'
        '<soapenv:Header/><soapenv:Body><tem:Consulta><tem:expresionImpresa>'
        f'<![CDATA[{expresion}]]>'
        '</tem:expresionImpresa></tem:Consulta></soapenv:Body></soapenv:Envelope>'
    )
    headers = {
        "Content-Type": 'text/xml;charset="utf-8"',
        "Accept": "text/xml",
        "SOAPAction": SOAP_ACTION_CONSULTA_CFDI,
    }

    try:
        resp = requests.post(
            URL_CONSULTA_CFDI_SAT,
            data=sobre_soap.encode('utf-8'),
            headers=headers,
            timeout=timeout,
        )
        resp.raise_for_status()

        root = ET.fromstring(resp.content)
        valores = {}
        for elem in root.iter():
            etiqueta = elem.tag.split('}')[-1]
            if etiqueta in ("CodigoEstatus", "EsCancelable", "Estado", "EstatusCancelacion", "ValidacionEFOS"):
                valores[etiqueta] = (elem.text or "").strip()

        codigo_estatus = valores.get("CodigoEstatus", "")
        estado = valores.get("Estado", "")

        if codigo_estatus.startswith("S"):
            estado_normalizado = estado or "Vigente"
        elif "602" in codigo_estatus:
            estado_normalizado = "No Encontrado"
        elif "601" in codigo_estatus:
            estado_normalizado = "Error de Consulta"
        else:
            estado_normalizado = estado or "No Encontrado"

        return {
            "estado": estado_normalizado,
            "es_cancelable": valores.get("EsCancelable", ""),
            "estatus_cancelacion": valores.get("EstatusCancelacion", ""),
            "validacion_efos_sat": valores.get("ValidacionEFOS", ""),
            "detalle": codigo_estatus or "El servicio no regresó código de estatus.",
        }
    except Exception as e:
        return {
            "estado": "Error de Consulta",
            "es_cancelable": "",
            "estatus_cancelacion": "",
            "validacion_efos_sat": "",
            "detalle": f"No fue posible consultar el Web Service del SAT: {e}",
        }

# ---------------------------------------------------------
# 7. CATÁLOGO OFICIAL SAT: CARGA LOCAL Y ACTUALIZACIÓN AUTOMÁTICA
# ---------------------------------------------------------
# Estrategia (ver diagrama de flujo del proyecto):
#
#   Inicia app -> ¿hay internet y actualización disponible en GitHub?
#                  SI  -> descarga y actualiza la DB local
#                  NO  -> continúa con la DB local que ya se tenga
#              -> ejecuta validaciones a ALTA VELOCIDAD desde la DB local
#
# El punto clave de velocidad: la base SQLite NO se consulta una vez por
# concepto (eso sería miles de queries por lote). Se lee UNA sola vez a un
# diccionario en memoria {clave: texto} y a partir de ahí todas las búsquedas
# son instantáneas. app.py cachea ese diccionario con st.cache_resource
# usando la fecha de modificación del archivo como parte de la llave, así que
# cuando el archivo se actualiza, el diccionario se recarga solo.

def ruta_catalogo_por_defecto(directorio_base=None):
    """Ruta esperada de sat_catalogos.db: junto a este archivo .py, salvo que
    se indique otro directorio."""
    base = directorio_base or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, NOMBRE_ARCHIVO_CATALOGO)


def catalogo_necesita_actualizacion(ruta_db, dias=DIAS_ENTRE_ACTUALIZACIONES_CATALOGO):
    """
    True si conviene intentar actualizar el catálogo: porque no existe, o
    porque ya pasaron `dias` desde su última modificación (por defecto 30,
    es decir una vez al mes como se definió).
    """
    if not os.path.exists(ruta_db):
        return True
    try:
        edad_dias = (datetime.now() - datetime.fromtimestamp(os.path.getmtime(ruta_db))).days
        return edad_dias >= dias
    except OSError:
        return True


def descargar_catalogo_sat(ruta_db, timeout=120):
    """
    Descarga la última versión del catálogo desde el repositorio público
    phpcfdi/resources-sat-catalogs y la deja descomprimida en `ruta_db`.

    El archivo publicado viene comprimido en bzip2 (catalogs.db.bz2, ~15 MB)
    y se descomprime al vuelo. La descarga se escribe primero a un archivo
    temporal y solo al final se reemplaza el definitivo, para que un corte de
    red a medias nunca deje la base corrupta.

    Devuelve (exito: bool, mensaje: str). Nunca lanza excepción: si no hay
    internet simplemente devuelve False y la app sigue con la DB que tenga.
    """
    ruta_temporal = f"{ruta_db}.descargando"
    try:
        respuesta = requests.get(URL_CATALOGO_SAT, timeout=timeout, stream=True)
        respuesta.raise_for_status()

        descompresor = bz2.BZ2Decompressor()
        with open(ruta_temporal, 'wb') as destino:
            for bloque in respuesta.iter_content(chunk_size=1024 * 256):
                if bloque:
                    destino.write(descompresor.decompress(bloque))

        # Verificación mínima antes de reemplazar: que el archivo descargado
        # sea realmente una base SQLite con la tabla que esperamos.
        with sqlite3.connect(ruta_temporal) as conexion:
            tabla = _detectar_tabla_prodserv(conexion)
            if not tabla:
                raise ValueError("El archivo descargado no contiene la tabla de productos y servicios.")
            total = conexion.execute(f"SELECT COUNT(*) FROM {tabla}").fetchone()[0]
            if total < 1000:
                raise ValueError(f"El catálogo descargado sólo trae {total} claves; se esperaban miles.")

        os.replace(ruta_temporal, ruta_db)
        return True, f"Catálogo SAT actualizado correctamente ({total:,} claves)."

    except Exception as e:
        try:
            if os.path.exists(ruta_temporal):
                os.remove(ruta_temporal)
        except OSError:
            pass
        return False, f"No se pudo actualizar el catálogo SAT (se usará la copia local): {e}"


def _detectar_tabla_prodserv(conexion):
    """Devuelve el nombre de la tabla de productos y servicios presente en la
    base: primero la de CFDI 4.0 y, si no está, la de CFDI 3.3. None si no
    encuentra ninguna."""
    for tabla in (TABLA_CATALOGO_PRODSERV, TABLA_CATALOGO_PRODSERV_FALLBACK):
        existe = conexion.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (tabla,)
        ).fetchone()
        if existe:
            return tabla
    return None


def cargar_catalogo_prodserv(ruta_db=None, solo_vigentes=True):
    """
    Lee TODA la tabla de Claves ProdServ de la base local a un diccionario
    {clave: descripción oficial} — una sola lectura, sin queries por concepto.

    `solo_vigentes=True` descarta las claves cuya vigencia ya terminó
    (vigencia_hasta con fecha pasada), que es lo correcto para auditar
    facturas actuales. Si la columna no existe en tu esquema, se ignora el
    filtro sin fallar.

    Devuelve (catalogo: dict, mensaje_estatus: str). Si la base no existe o
    no se puede leer, devuelve ({}, motivo) y la app sigue funcionando con el
    catálogo interno NOMBRES_CLAVES_SAT como respaldo.
    """
    ruta_db = ruta_db or ruta_catalogo_por_defecto()

    if not os.path.exists(ruta_db):
        return {}, (
            f"No se encontró {NOMBRE_ARCHIVO_CATALOGO}. Se usará el catálogo interno reducido; "
            "muchas claves aparecerán como NO CATALOGADA."
        )

    try:
        with sqlite3.connect(f"file:{ruta_db}?mode=ro", uri=True) as conexion:
            tabla = _detectar_tabla_prodserv(conexion)
            if not tabla:
                return {}, (
                    f"{NOMBRE_ARCHIVO_CATALOGO} no contiene la tabla '{TABLA_CATALOGO_PRODSERV}'. "
                    "Verifique que sea la base de phpcfdi/resources-sat-catalogs."
                )

            columnas = {fila[1] for fila in conexion.execute(f"PRAGMA table_info({tabla})")}
            consulta = f"SELECT {COLUMNA_CATALOGO_CLAVE}, {COLUMNA_CATALOGO_TEXTO} FROM {tabla}"

            if solo_vigentes and 'vigencia_hasta' in columnas:
                # Se conservan las claves sin fecha de fin (vigentes) y las que
                # aún no vencen a la fecha de hoy.
                hoy = datetime.now().strftime('%Y-%m-%d')
                consulta += (
                    " WHERE vigencia_hasta IS NULL OR vigencia_hasta = '' "
                    f"OR vigencia_hasta >= '{hoy}'"
                )

            catalogo = {
                str(clave).strip(): str(texto).strip()
                for clave, texto in conexion.execute(consulta)
                if clave and texto
            }

        fecha_archivo = datetime.fromtimestamp(os.path.getmtime(ruta_db)).strftime('%d/%m/%Y')
        return catalogo, f"Catálogo SAT cargado: {len(catalogo):,} claves (actualizado el {fecha_archivo})."

    except Exception as e:
        return {}, f"No se pudo leer {NOMBRE_ARCHIVO_CATALOGO}: {e}"


def preparar_catalogo_sat(ruta_db=None, permitir_descarga=True):
    """
    Punto de entrada único que ejecuta el flujo completo del diagrama:
    revisa si toca actualizar, intenta descargar si hay internet, y carga el
    catálogo a memoria pase lo que pase.

    Devuelve (catalogo: dict, mensaje_estatus: str) listo para mostrarse en
    la barra lateral de la app.
    """
    ruta_db = ruta_db or ruta_catalogo_por_defecto()
    mensajes = []

    if permitir_descarga and catalogo_necesita_actualizacion(ruta_db):
        exito, mensaje = descargar_catalogo_sat(ruta_db)
        mensajes.append(mensaje)
        if not exito and not os.path.exists(ruta_db):
            # Sin internet y sin copia local: se sigue con el catálogo interno.
            return {}, " ".join(mensajes)

    catalogo, mensaje_carga = cargar_catalogo_prodserv(ruta_db)
    mensajes.append(mensaje_carga)
    return catalogo, " ".join(mensajes)
