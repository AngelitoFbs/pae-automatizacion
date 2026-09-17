"""Configuración central PAE - Multi-operador, multi-mes"""
from dataclasses import dataclass
from typing import Dict, List
from pathlib import Path

# ─── Operadores ───
OPERADORES = {
    "kelly_primo": {
        "nombre": "Kelly Primo",
        "codigo": "KP",
        "color": "#E7B52A",
        "icon": "👩‍💼",
        "descripcion": "Operación y supervisión de entrega PAE",
    },
    "laura_jimenez": {
        "nombre": "Laura Jiménez",
        "codigo": "LJ",
        "color": "#3B82F6",
        "icon": "📊",
        "descripcion": "Consolidación y reporte de cobertura",
    },
}

# ─── Meses PAE ───
MESES_PAE = [
    ("ENERO", 1), ("FEBRERO", 2), ("MARZO", 3), ("ABRIL", 4),
    ("MAYO", 5), ("JUNIO", 6), ("JULIO", 7), ("AGOSTO", 8),
    ("SEPTIEMBRE", 9), ("OCTUBRE", 10), ("NOVIEMBRE", 11), ("DICIEMBRE", 12),
]

# ─── Tipos de ración ───
TIPOS_RACION = ["CAJM", "CAJT", "ALMUERZO"]
NIVELES = ["A", "B", "C", "D", "E"]

# ─── Columnas Cobertura estándar ───
NIVEL_COLS = {
    'A': {'AM': 3, 'PM': 4, 'Días': 6, 'Total_Cob': 5, 'Total_Rac': 7, 'Valor': 23},
    'B': {'AM': 8, 'PM': 9, 'Días': 11, 'Total_Cob': 10, 'Total_Rac': 12, 'Valor': 24},
    'C': {'AM': 13, 'PM': 14, 'Días': 16, 'Total_Cob': 15, 'Total_Rac': 17, 'Valor': 25},
    'D': {'AM': 18, 'PM': 19, 'Días': 21, 'Total_Cob': 20, 'Total_Rac': 22, 'Valor': 26},
}
VALOR_TOTAL_COL = 27
DATA_START_ROW = 4
NAME_COL = 2
DANE_COL = 28

# ─── Rutas base ───
BASE_DIR = Path(__file__).parent.parent.parent.parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"

for d in [DATA_DIR, TEMPLATES_DIR, OUTPUT_DIR, LOGS_DIR]:
    d.mkdir(parents=True, exist_ok=True)


@dataclass
class AppConfig:
    operador: str = "kelly_primo"
    mes: str = "JULIO"
    mes_num: int = 7
    anio: int = 2026
    
    @property
    def mes_key(self) -> str:
        return f"{self.mes}_{self.anio}"
    
    @property
    def output_prefix(self) -> str:
        op = OPERADORES[self.operador]["codigo"]
        return f"{op}_{self.mes}_{self.anio}"


def get_operador_config(operador: str) -> dict:
    return OPERADORES.get(operador, OPERADORES["kelly_primo"])


def get_template_names(mes: str) -> tuple:
    """Nombres esperados de plantillas por mes"""
    return (
        f"2_CERTIFICACIONES_MES_DE_{mes.upper()}.xlsx",
        f"3_COBERTURA_EJECUTADA_{mes.upper()}.xlsx",
    )