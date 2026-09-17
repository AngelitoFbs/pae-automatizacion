"""
Configuración central PAE - Estructura de 4 bloques según prompt original.
"""
from dataclasses import dataclass
from typing import Dict, List, Optional
from pathlib import Path


# ─── Configuración de bloques según plantilla Cobertura ───
# Cada bloque tiene: nombre, fila inicio, fila fin, columnas por nivel, tarifas

BLOQUES = {
    1: {
        "nombre": "PAE_REGULAR_COMPLEMENTO_AM_PM_PS",
        "descripcion": "PAE REGULAR COMPLEMENTO AM / PM RACION PREPARADA EN SITIO",
        "fila_inicio": 4,
        "fila_fin": 46,
        "total_fila": 47,
        "valores_fila": 48,
        "resumen_fila": 50,
        "niveles": ["A", "B", "C", "D"],
        "columnas_por_nivel": {
            "A": {"AM": 3, "PM": 4, "TOTAL_COB": 5, "DIAS": 6, "TOTAL_RAC": 7, "VALOR": 23},
            "B": {"AM": 8, "PM": 9, "TOTAL_COB": 9, "DIAS": 10, "TOTAL_RAC": 11, "VALOR": 24},
            "C": {"AM": 13, "PM": 14, "TOTAL_COB": 14, "DIAS": 15, "TOTAL_RAC": 16, "VALOR": 25},
            "D": {"AM": 18, "PM": 19, "TOTAL_COB": 19, "DIAS": 20, "TOTAL_RAC": 20, "VALOR": 26},
        },
        "valor_cols": {"A": 23, "B": 24, "C": 25, "D": 26},
        "total_valor_fila": 48,
        "tarifas": {"A": 4298, "B": 4685, "C": 5039, "D": 5638},
    },
    2: {
        "nombre": "COMPLEMENTO_ALIMENTARIO_TRANSPORTADO_CALIENTE",
        "descripcion": "COMPLEMENTO ALIMENTARIO TRANSPORTADO EN CALIENTE",
        "fila_inicio": 56,
        "fila_fin": 66,
        "total_fila": 67,
        "valores_fila": 68,
        "resumen_fila": 70,
        "niveles": ["A", "B", "C", "D"],
        "columnas_por_nivel": {
            "A": {"AM": 3, "PM": 4, "TOTAL_COB": 5, "DIAS": 6, "TOTAL_RAC": 7, "VALOR": 23},
            "B": {"AM": 8, "PM": 9, "TOTAL_COB": 9, "DIAS": 10, "TOTAL_RAC": 10, "VALOR": 24},
            "C": {"AM": 13, "PM": 14, "TOTAL_COB": 14, "DIAS": 15, "TOTAL_RAC": 15, "VALOR": 25},
            "D": {"AM": 18, "PM": 19, "TOTAL_COB": 19, "DIAS": 20, "TOTAL_RAC": 20, "VALOR": 26},
        },
        "valor_cols": {"A": 23, "B": 24, "C": 25, "D": 26},
        "total_valor_fila": 68,
        "tarifas": {"A": 4517, "B": 4905, "C": 5258, "D": 5856},
    },
    3: {
        "nombre": "ALMUERZO_JORNADA_UNICA_PS",
        "descripcion": "ALMUERZO JORNADA UNICA PREPARADO EN SITIO",
        "fila_inicio": 75,
        "fila_fin": 110,
        "total_fila": 111,
        "valores_fila": 112,
        "resumen_fila": 113,
        "niveles": ["A", "B", "C", "D", "E"],
        "columnas_por_nivel": {
            "A": {"COB": 3, "DIAS": 4, "TOTAL_RAC": 5, "VALOR": 18},
            "B": {"COB": 6, "DIAS": 7, "TOTAL_RAC": 8, "VALOR": 19},
            "C": {"COB": 9, "DIAS": 10, "TOTAL_RAC": 11, "VALOR": 20},
            "D": {"COB": 12, "DIAS": 13, "TOTAL_RAC": 13, "VALOR": 21},
            "E": {"COB": 15, "DIAS": 16, "TOTAL_RAC": 16, "VALOR": 17},
        },
        "valor_cols": {"A": 18, "B": 19, "C": 20, "D": 21, "E": 22},
        "total_valor_fila": 112,
        "tarifas": {"A": 4500, "B": 4880, "C": 5270, "D": 5886, "E": 6386},
    },
    4: {
        "nombre": "ALMUERZO_JORNADA_UNICA_CCT",
        "descripcion": "ALMUERZO JORNADA UNICA TRANSPORTADO CALIENTE",
        "fila_inicio": 119,
        "fila_fin": 125,
        "total_fila": 126,
        "valores_fila": 127,
        "resumen_fila": 129,
        "niveles": ["A", "B", "C", "D", "E"],
        "columnas_por_nivel": {
            "A": {"COB": 3, "DIAS": 4, "TOTAL_RAC": 5, "VALOR": 18},
            "B": {"COB": 6, "DIAS": 7, "TOTAL_RAC": 7, "VALOR": 19},
            "C": {"COB": 9, "DIAS": 10, "TOTAL_RAC": 10, "VALOR": 20},
            "D": {"COB": 12, "DIAS": 13, "TOTAL_RAC": 13, "VALOR": 21},
            "E": {"COB": 15, "DIAS": 16, "TOTAL_RAC": 16, "VALOR": 17},
        },
        "valor_cols": {"A": 18, "B": 19, "C": 20, "D": 21, "E": 22},
        "total_valor_fila": 127,
        "tarifas": {"A": 4718, "B": 5097, "C": 5487, "D": 6106, "E": 6603},
    },
}

# Mapeo de modalidad en certificado → bloque destino
MODALIDAD_A_BLOQUE = {
    ("CAJM", "PS"): 1,
    ("CAJT", "PS"): 1,
    ("CAJM", "CCT"): 2,
    ("CAJT", "CCT"): 2,
    ("ALMUERZO", "PS"): 3,
    ("ALMUERZO", "CCT"): 4,
}

# ─── Certificado: detección dinámica de celdas fijas ───
# La plantilla inicial tiene institucion en B6/I6, la completa en B7/I7
# Buscamos por texto de encabezado para ser robustos

def detectar_celdas_fijas_certificado(ws) -> Dict[str, str]:
    """Detecta celdas fijas buscando por texto de encabezado."""
    import re
    celdas = {}
    for row in ws.iter_rows(min_row=1, max_row=20, max_col=10, values_only=False):
        for cell in row:
            if not cell.value:
                continue
            val = str(cell.value).strip().upper()
            coord = cell.coordinate
            if 'OPERADOR' in val and 'CONTRATO' not in val:
                celdas['operador'] = coord
            elif 'CONTRATO' in val and 'N' in val:
                celdas['contrato'] = coord
            elif ('INSTITUCION' in val or 'CENTRO EDUCATIVO' in val) and 'institucion' not in celdas:
                celdas['institucion'] = coord
            elif 'CODIGO DANE' in val or 'CÓDIGO DANE' in val:
                # El valor está en la columna siguiente
                valor_cell = cell.offset(column=1)
                celdas['codigo_dane'] = valor_cell.coordinate
                # Extraer número DANE del valor
                valor_val = str(valor_cell.value).strip() if valor_cell.value else ""
                nums = re.findall(r'\d{10,}', valor_val)
                if nums and 'codigo_dane_valor' not in celdas:
                    celdas['codigo_dane_valor'] = nums[0]
            elif 'DEPARTAMENTO' in val:
                celdas['departamento'] = coord
            elif 'MUNICIPIO' in val:
                celdas['municipio'] = coord
            elif 'DESDE' == val:  # Exact match for "DESDE"
                celdas['fecha_desde'] = cell.offset(column=1).coordinate
            elif 'HASTA' == val:  # Exact match for "HASTA"
                celdas['fecha_hasta'] = cell.offset(column=1).coordinate
            elif 'FECHA' in val and 'DESDE' in val:
                celdas['fecha_desde'] = cell.offset(column=1).coordinate
            elif 'FECHA' in val and 'HASTA' in val:
                celdas['fecha_hasta'] = cell.offset(column=1).coordinate
            elif 'RECTOR' in val:
                celdas['rector'] = coord
    return celdas

# Columnas de la tabla de raciones en Certificado
CERT_TABLA_COLS = {
    'tipo': 2,        # B
    'clasificacion': 3,  # C
    'raciones_dia': 4,   # D
    'dias': 5,           # E
    'total': 6,          # F
}

# Columnas fijas en Cobertura
NAME_COL = 2
DANE_COL = 3
DATA_START_ROW = 4

# Rutas
BASE_DIR = Path(__file__).parent.parent.parent.parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"
MAPEO_FILE = DATA_DIR / "mapeo_colegios.csv"
TARIFAS_FILE = DATA_DIR / "tarifas.csv"

for d in [DATA_DIR, TEMPLATES_DIR, OUTPUT_DIR, LOGS_DIR]:
    d.mkdir(parents=True, exist_ok=True)


@dataclass
class ColegioData:
    """Datos extraídos de un colegio en Certificado."""
    hoja: str
    nombre: str
    codigo_dane: str
    departamento: str
    municipio: str
    fecha_desde: str
    fecha_hasta: str
    rector: str
    raciones: Dict[str, Dict[str, Dict[str, int]]]


def get_bloque_por_tipo_modalidad(tipo_racion: str, modalidad: str) -> Optional[int]:
    """Determina bloque destino según tipo de ración y modalidad."""
    return MODALIDAD_A_BLOQUE.get((tipo_racion, modalidad))


def normalizar_modalidad(texto: str) -> str:
    """Normaliza texto de modalidad a PS o CCT."""
    if not texto:
        return "PS"
    t = texto.upper().replace(" ", "").replace("-", "").replace("_", "")
    if any(x in t for x in ["CCT", "TCC", "TRANSPORTADO", "CALIENTE", "TRANSPORTADOCALIENTE"]):
        return "CCT"
    return "PS"


def normalizar_tipo_racion(texto: str) -> str:
    """Normaliza tipo de ración a CAJM, CAJT, ALMUERZO."""
    if not texto:
        return ""
    t = texto.upper().replace(" ", "").replace("-", "").replace("_", "").replace("/", "")
    if "ALMUERZO" in t:
        return "ALMUERZO"
    if "CAJM" in t and "CAJT" in t:
        return "CAJM/JT"
    if "CAJM" in t:
        return "CAJM"
    if "CAJT" in t:
        return "CAJT"
    return ""


def detectar_fila_encabezado_tabla(ws) -> Optional[int]:
    """Detecta fila donde está el encabezado 'TIPO RACIÓN'."""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=False):
        for cell in row:
            if cell.value and "TIPO RACI" in str(cell.value).upper():
                return cell.row
    return None


def detectar_fila_total(ws, start_row: int) -> Optional[int]:
    """Detecta fila de TOTAL al final de la tabla."""
    for row in range(start_row, ws.max_row + 1):
        val = ws.cell(row=row, column=2).value
        if val and "TOTAL" in str(val).upper():
            return row
    return None