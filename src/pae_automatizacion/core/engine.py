"""Motor base PAE - Lógica compartida de lectura/escritura de plantillas Excel.

Este módulo proporciona las clases principales para:
- Leer plantillas Certificado (multi-hoja)
- Escribir en plantillas Cobertura preservando fórmulas y formato
- Gestionar tarifas por grupo y nivel
- Orquestar el proceso completo Certificado → Cobertura
"""
import openpyxl
import pandas as pd
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Set
from dataclasses import dataclass, asdict
from collections import defaultdict

from .config import (
    NIVEL_COLS, DATA_START_ROW, NAME_COL, DANE_COL, 
    TIPOS_RACION, NIVELES
)

# Configurar logging solo una vez
if not logging.getLogger().handlers:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


@dataclass
class ColegioData:
    """Datos extraídos de un colegio en la plantilla Certificado."""
    nombre: str
    codigo_dane: str
    departamento: str
    municipio: str
    fecha_desde: str
    fecha_hasta: str
    rector: str
    raciones: Dict[str, Dict[str, Dict[str, int]]]


class CertificadoReader:
    """Lee plantilla Certificado (multi-hoja) y extrae datos por colegio"""
    
    CELDAS_FIJAS = {
        'nombre': 'B7',
        'codigo_dane': 'I7',
        'departamento': 'B8',
        'municipio': 'B9',
        'fecha_desde': 'C10',
        'fecha_hasta': 'I10',
        'rector': 'B11',
    }

    def __init__(self, filepath: str):
        self.filepath = Path(filepath)
        self.wb = openpyxl.load_workbook(filepath, data_only=True)

    def read_all(self) -> Dict[str, ColegioData]:
        colegios = {}
        for sheet_name in self.wb.sheetnames:
            ws = self.wb[sheet_name]
            try:
                colegio = self._read_sheet(ws, sheet_name)
                if colegio and colegio.codigo_dane:
                    colegios[colegio.codigo_dane] = colegio
                    logger.info(f"Leído: {colegio.nombre} (DANE: {colegio.codigo_dane})")
                else:
                    logger.warning(f"Hoja '{sheet_name}' sin DANE válido, omitida")
            except Exception as e:
                logger.error(f"Error leyendo hoja '{sheet_name}': {e}")
        return colegios

    def _read_sheet(self, ws, sheet_name: str) -> Optional[ColegioData]:
        datos_fijos = {}
        for key, cell in self.CELDAS_FIJAS.items():
            val = ws[cell].value
            datos_fijos[key] = str(val).strip() if val else ""

        if not datos_fijos['codigo_dane']:
            return None

        raciones = self._extract_raciones(ws)
        return ColegioData(
            nombre=datos_fijos['nombre'],
            codigo_dane=datos_fijos['codigo_dane'],
            departamento=datos_fijos['departamento'],
            municipio=datos_fijos['municipio'],
            fecha_desde=datos_fijos['fecha_desde'],
            fecha_hasta=datos_fijos['fecha_hasta'],
            rector=datos_fijos['rector'],
            raciones=raciones
        )

    def _extract_raciones(self, ws) -> Dict[str, Dict[str, Dict[str, int]]]:
        raciones = defaultdict(lambda: defaultdict(dict))

        header_row = None
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=False):
            for cell in row:
                if cell.value and 'TIPO RACI' in str(cell.value).upper():
                    header_row = cell.row
                    break
            if header_row:
                break

        if not header_row:
            logger.warning("No se encontró fila de encabezado 'TIPO RACIÓN'")
            return dict(raciones)

        sub_header_row = header_row + 1

        col_tipo = None
        col_nivel = None
        col_raciones_dia = None
        col_dias = None
        col_total = None

        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=sub_header_row, column=col).value
            if val:
                v = str(val).upper()
                if 'RACIONES' in v and 'DÍA' in v:
                    col_raciones_dia = col
                elif 'DÍAS' in v and 'ATEND' in v:
                    col_dias = col
                elif 'TOTAL' in v and 'RACIONES' in v and 'ENTREG' in v:
                    col_total = col

        val_tipo = ws.cell(row=header_row, column=2).value
        val_nivel = ws.cell(row=header_row, column=3).value
        if val_tipo and 'TIPO' in str(val_tipo).upper() and 'RACI' in str(val_tipo).upper():
            col_tipo = 2
        if val_nivel and ('CLASIFIC' in str(val_nivel).upper() or 'NIVEL' in str(val_nivel).upper()):
            col_nivel = 3

        if not all([col_tipo, col_nivel, col_raciones_dia, col_dias]):
            logger.warning("No se encontraron todas las columnas de la tabla de raciones")
            return dict(raciones)

        current_tipo = None
        for row in range(sub_header_row + 1, ws.max_row + 1):
            tipo_val = ws.cell(row=row, column=col_tipo).value
            nivel_val = ws.cell(row=row, column=col_nivel).value
            rac_dia_val = ws.cell(row=row, column=col_raciones_dia).value
            dias_val = ws.cell(row=row, column=col_dias).value
            total_val = ws.cell(row=row, column=col_total).value if col_total else None

            if tipo_val:
                tipo_str = str(tipo_val).strip().upper()
                if 'TOTAL' in tipo_str:
                    break
                if tipo_str in ['CAJM/JT PS', 'CAJM/JT', 'CAJM', 'CAJT']:
                    current_tipo = 'CAJM' if 'CAJM' in tipo_str else 'CAJT'
                elif 'ALMUERZO' in tipo_str or 'JORNADA' in tipo_str:
                    current_tipo = 'ALMUERZO'

            if not current_tipo or not nivel_val:
                continue

            nivel_str = str(nivel_val).strip().upper()
            nivel = None
            if 'NIVEL A' in nivel_str or 'NIVEL A-' in nivel_str:
                nivel = 'A'
            elif 'NIVEL B' in nivel_str:
                nivel = 'B'
            elif 'NIVEL C' in nivel_str:
                nivel = 'C'
            elif 'NIVEL D' in nivel_str:
                nivel = 'D'
            elif 'NIVEL E' in nivel_str:
                nivel = 'E'

            if nivel and nivel in NIVELES:
                rac_dia = int(rac_dia_val) if rac_dia_val else 0
                dias = int(dias_val) if dias_val else 0
                total = int(total_val) if total_val else 0
                if rac_dia > 0 or dias > 0 or total > 0:
                    raciones[current_tipo][nivel] = {
                        'raciones_dia': rac_dia,
                        'dias': dias,
                        'total': total
                    }

        return dict(raciones)


class CoberturaWriter:
    """Escribe en plantilla Cobertura preservando fórmulas y formato"""

    def __init__(self, template_path: str):
        self.template_path = Path(template_path)
        self.wb = openpyxl.load_workbook(template_path)
        self.ws = self.wb.active
        self.name_to_row: Dict[str, int] = {}
        # Detectar celdas combinadas para evitar escribir en ellas
        self._merged_ranges = set()
        for merged_range in self.ws.merged_cells.ranges:
            for row in range(merged_range.min_row, merged_range.max_row + 1):
                for col in range(merged_range.min_col, merged_range.max_col + 1):
                    self._merged_ranges.add((row, col))
        self._build_name_map()
        self.template_path = Path(template_path)
        self.wb = openpyxl.load_workbook(template_path)
        self.ws = self.wb.active
        self.name_to_row: Dict[str, int] = {}
        self._build_name_map()

    def _normalize_name(self, nombre: str) -> str:
        n = nombre.upper()
        for term in ['INSTITUCION EDUCATIVA', 'INSTITUCION E.', 'INSTITUCION', 
                     'I.E.', 'I.E', 'INST.', 'INST', 'SEDE', 'Nº', 'NO.', '.']:
            n = n.replace(term, '')
        return ' '.join(n.split()).strip()

    def _build_name_map(self):
        for row in range(DATA_START_ROW, self.ws.max_row + 1):
            name_cell = self.ws.cell(row=row, column=NAME_COL)
            if name_cell.value:
                nombre = str(name_cell.value).strip()
                if nombre:
                    norm = self._normalize_name(nombre)
                    if norm:
                        self.name_to_row[norm] = row

    def find_or_create_row(self, dane: str, nombre: str) -> Tuple[int, bool]:
        norm_nombre = self._normalize_name(nombre)

        # Match exacto normalizado
        if norm_nombre in self.name_to_row:
            row = self.name_to_row[norm_nombre]
            self._safe_write_cell(row, DANE_COL, dane)
            return row, False

        # Match fuzzy (substring)
        for norm_name, row in self.name_to_row.items():
            if norm_nombre in norm_name or norm_name in norm_nombre:
                self._safe_write_cell(row, DANE_COL, dane)
                return row, True

        # Crear nueva fila
        new_row = self.ws.max_row + 1
        self._safe_write_cell(new_row, 1, new_row - DATA_START_ROW + 1)
        self._safe_write_cell(new_row, NAME_COL, nombre)
        self.ws.cell(row=new_row, column=DANE_COL, value=dane)
        self.name_to_row[norm_nombre] = new_row
        return new_row, True

    def write_colegio(self, colegio: ColegioData) -> List[dict]:
        rows_written = []
        cajm = colegio.raciones.get('CAJM', {})
        cajt = colegio.raciones.get('CAJT', {})

        # Determinar niveles con datos en CAJM o CAJT
        niveles_con_datos = set()
        for nivel in ['A', 'B', 'C', 'D']:
            if nivel in cajm and cajm[nivel].get('raciones_dia', 0) > 0:
                niveles_con_datos.add(nivel)
            if nivel in cajt and cajt[nivel].get('raciones_dia', 0) > 0:
                niveles_con_datos.add(nivel)

        if not niveles_con_datos:
            logger.warning(f"Colegio {colegio.nombre} sin datos CAJM/CAJT para niveles A-D")
            return rows_written

        # Recolectar días únicos por tipo
        dias_cajm = {cajm[n].get('dias', 0) for n in niveles_con_datos if n in cajm and cajm[n].get('dias', 0) > 0}
        dias_cajt = {cajt[n].get('dias', 0) for n in niveles_con_datos if n in cajt and cajt[n].get('dias', 0) > 0}

        usar_dos_filas = len(dias_cajm) > 0 and len(dias_cajt) > 0 and dias_cajm != dias_cajt

        if usar_dos_filas:
            row1, _ = self.find_or_create_row(colegio.codigo_dane, colegio.nombre)
            row2 = self.ws.max_row + 1
            self._copy_row_format(row1, row2)
            self.ws.cell(row=row2, column=1, value=row2 - DATA_START_ROW + 1)
            self.ws.cell(row=row2, column=NAME_COL, value=f"{colegio.nombre} (TARDE)")
            self.ws.cell(row=row2, column=DANE_COL, value=colegio.codigo_dane)

            rows_written.append(self._write_fila_am(row1, cajm))
            rows_written.append(self._write_fila_pm(row2, cajt))
        else:
            row, _ = self.find_or_create_row(colegio.codigo_dane, colegio.nombre)
            dias_comunes = (dias_cajm | dias_cajt) - {0}
            dias_val = list(dias_comunes)[0] if dias_comunes else 0

            rows_written.append(self._write_fila_completa(row, cajm, cajt, dias_val))

        return rows_written

    def _safe_write_cell(self, row: int, col: int, value: Any):
        """Escribe valor en celda solo si no es una MergedCell."""
        if (row, col) not in self._merged_ranges:
            self.ws.cell(row=row, column=col, value=value)

    def _write_fila_am(self, row: int, cajm: Dict) -> dict:
        data = {'fila': row, 'dane': '', 'nombre': '', 'tipo': 'AM'}
        for nivel in ['A', 'B', 'C', 'D']:
            cols = NIVEL_COLS[nivel]
            if nivel in cajm:
                rac_dia = cajm[nivel].get('raciones_dia', 0)
                dias = cajm[nivel].get('dias', 0)
                if rac_dia > 0:
                    self._safe_write_cell(row, cols['AM'], rac_dia)
                    data[f'{nivel}_AM'] = rac_dia
                if dias > 0:
                    self._safe_write_cell(row, cols['Días'], dias)
                    data[f'{nivel}_Días'] = dias
        return data

    def _write_fila_pm(self, row: int, cajt: Dict) -> dict:
        data = {'fila': row, 'dane': '', 'nombre': '', 'tipo': 'PM'}
        for nivel in ['A', 'B', 'C', 'D']:
            cols = NIVEL_COLS[nivel]
            if nivel in cajt:
                rac_dia = cajt[nivel].get('raciones_dia', 0)
                dias = cajt[nivel].get('dias', 0)
                if rac_dia > 0:
                    self._safe_write_cell(row, cols['PM'], rac_dia)
                    data[f'{nivel}_PM'] = rac_dia
                if dias > 0:
                    self._safe_write_cell(row, cols['Días'], dias)
                    data[f'{nivel}_Días'] = dias
        return data

    def _write_fila_completa(self, row: int, cajm: Dict, cajt: Dict, dias: int) -> dict:
        data = {'fila': row, 'dane': '', 'nombre': '', 'tipo': 'AM+PM'}
        for nivel in ['A', 'B', 'C', 'D']:
            cols = NIVEL_COLS[nivel]
            if nivel in cajm:
                rac_dia = cajm[nivel].get('raciones_dia', 0)
                if rac_dia > 0:
                    self._safe_write_cell(row, cols['AM'], rac_dia)
                    data[f'{nivel}_AM'] = rac_dia
            if nivel in cajt:
                rac_dia = cajt[nivel].get('raciones_dia', 0)
                if rac_dia > 0:
                    self._safe_write_cell(row, cols['PM'], rac_dia)
                    data[f'{nivel}_PM'] = rac_dia
            if dias > 0:
                self._safe_write_cell(row, cols['Días'], dias)
                data[f'{nivel}_Días'] = dias
        return data

    def _copy_row_format(self, src_row: int, dst_row: int):
        for col in range(1, self.ws.max_column + 1):
            if (dst_row, col) in self._merged_ranges:
                continue
            src = self.ws.cell(row=src_row, column=col)
            dst = self.ws.cell(row=dst_row, column=col)
            if src.has_style:
                dst.font = src.font.copy()
                dst.border = src.border.copy()
                dst.fill = src.fill.copy()
                dst.number_format = src.number_format
                dst.alignment = src.alignment.copy()
                dst.protection = src.protection.copy()

    def save(self, output_path: str):
        self.wb.save(output_path)
        logger.info(f"Guardado: {output_path}")

    def get_preview_data(self) -> List[dict]:
        preview = []
        for row in range(DATA_START_ROW, self.ws.max_row + 1):
            dane = self.ws.cell(row=row, column=DANE_COL).value
            nombre = self.ws.cell(row=row, column=NAME_COL).value
            if not dane and not nombre:
                continue
            row_data = {'fila': row, 'dane': dane, 'nombre': nombre}
            for nivel in ['A', 'B', 'C', 'D']:
                cols = NIVEL_COLS[nivel]
                row_data[f'{nivel}_AM'] = self.ws.cell(row=row, column=cols['AM']).value
                row_data[f'{nivel}_PM'] = self.ws.cell(row=row, column=cols['PM']).value
                row_data[f'{nivel}_Días'] = self.ws.cell(row=row, column=cols['Días']).value
                row_data[f'{nivel}_Total_Cob'] = self.ws.cell(row=row, column=cols['Total_Cob']).value
                row_data[f'{nivel}_Total_Rac'] = self.ws.cell(row=row, column=cols['Total_Rac']).value
                row_data[f'{nivel}_Valor'] = self.ws.cell(row=row, column=cols['Valor']).value
            row_data['Valor_Total'] = self.ws.cell(row=row, column=27).value
            preview.append(row_data)
        return preview


class TarifasManager:
    """Gestiona tarifas por grupo y nivel"""
    
    def __init__(self, tarifas_path: str = "tarifas.csv", colegios_tarifas_path: str = "colegios_tarifas.csv"):
        self.tarifas_path = Path(tarifas_path)
        self.colegios_tarifas_path = Path(colegios_tarifas_path)
        self.tarifas: Dict[str, Dict[str, int]] = {}
        self.colegio_grupo: Dict[str, str] = {}
        self.load_tarifas()
        self.load_colegios_tarifas()

    def load_tarifas(self):
        if not self.tarifas_path.exists():
            logger.warning(f"Tarifas no encontradas: {self.tarifas_path}")
            return
        df = pd.read_csv(self.tarifas_path)
        for _, row in df.iterrows():
            grupo = row['grupo']
            nivel = row['nivel']
            tarifa = int(row['tarifa'])
            self.tarifas.setdefault(grupo, {})[nivel] = tarifa

    def load_colegios_tarifas(self):
        if not self.colegios_tarifas_path.exists():
            logger.info("Mapeo colegios-tarifas no encontrado (opcional)")
            return
        try:
            df = pd.read_csv(self.colegios_tarifas_path, comment='#')
            if 'codigo_dane' in df.columns and 'grupo_tarifa' in df.columns:
                self.colegio_grupo = dict(zip(df['codigo_dane'].astype(str), df['grupo_tarifa']))
                logger.info(f"Mapeo cargado: {len(self.colegio_grupo)} entradas")
        except Exception as e:
            logger.warning(f"Error cargando mapeo: {e}")

    def get_tarifa(self, grupo: str, nivel: str) -> Optional[int]:
        return self.tarifas.get(grupo, {}).get(nivel)

    def get_grupo_colegio(self, dane: str) -> str:
        return self.colegio_grupo.get(dane, "grupo_1")

    def get_grupos(self) -> List[str]:
        return list(self.tarifas.keys())


class PAEEngine:
    """Orquesta el proceso completo Certificado → Cobertura"""
    
    def __init__(self, certificado_path: str, cobertura_path: str, 
                 tarifas_path: str = "tarifas.csv", 
                 colegios_tarifas_path: str = "colegios_tarifas.csv"):
        self.reader = CertificadoReader(certificado_path)
        self.writer = CoberturaWriter(cobertura_path)
        self.tarifas = TarifasManager(tarifas_path, colegios_tarifas_path)
        self.colegios: Dict[str, ColegioData] = {}
        self.resultado: List[dict] = []

    def procesar(self) -> List[dict]:
        logger.info("Leyendo Certificado...")
        self.colegios = self.reader.read_all()
        logger.info(f"Total colegios: {len(self.colegios)}")

        logger.info("Escribiendo Cobertura...")
        self.resultado = []
        for dane, colegio in self.colegios.items():
            rows = self.writer.write_colegio(colegio)
            if rows:
                self.resultado.extend(rows)
            else:
                logger.warning(f"No se escribieron datos para {colegio.nombre}")

        return self.resultado

    def save_output(self, output_path: str):
        self.writer.save(output_path)

    def get_preview(self) -> pd.DataFrame:
        preview = self.writer.get_preview_data()
        return pd.DataFrame(preview) if preview else pd.DataFrame()

    def get_log_data(self) -> dict:
        certificado_names = {self.writer._normalize_name(c.nombre) for c in self.colegios.values()}
        cobertura_names = set(self.writer.name_to_row.keys())
        
        return {
            'colegios_procesados': len(self.colegios),
            'filas_escritas': len(self.resultado),
            'colegios_sin_cobertura': list(certificado_names - cobertura_names),
            'colegios_sin_certificado': list(cobertura_names - certificado_names),
            'detalle': self.resultado,
        }

    def save_log(self, log_path: str):
        log = self.get_log_data()
        with open(log_path, 'w', encoding='utf-8') as f:
            json.dump(log, f, ensure_ascii=False, indent=2)