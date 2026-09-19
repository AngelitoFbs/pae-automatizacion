"""
Motor PAE - Procesamiento Certificado → Cobertura (4 bloques).
Siguiendo especificación original: extraer → generar, mapeo persistente, 4 bloques.
"""
import openpyxl
import pandas as pd
import json
import logging
import csv
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Set
from dataclasses import dataclass, asdict
from collections import defaultdict
from copy import copy

from .config import (
    BLOQUES, MODALIDAD_A_BLOQUE, CERT_TABLA_COLS,
    NAME_COL, DANE_COL, DATA_START_ROW,
    MAPEO_FILE, TARIFAS_FILE,
    normalizar_modalidad, normalizar_tipo_racion,
    detectar_fila_encabezado_tabla, get_bloque_por_tipo_modalidad,
    detectar_celdas_fijas_certificado, detectar_fila_total,
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


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


@dataclass
class RegistroBorrador:
    """Registro para el Excel borrador (Paso A: extraer)."""
    dane: str
    hoja_certificado: str
    nombre_certificado: str
    nombre_cobertura: str
    bloque: int
    bloque_nombre: str
    fila_am: int
    fila_pm: int
    tipo_racion: str
    modalidad: str
    nivel: str
    raciones_dia: int
    dias: int
    total_raciones: int
    estado: str  # "ok", "revisar", "vacío"
    observaciones: str


class CertificadoReader:
    """Lee plantilla Certificado (multi-hoja) y extrae datos por institución."""

    def __init__(self, filepath: str):
        self.filepath = Path(filepath)
        self.wb = openpyxl.load_workbook(filepath, data_only=True)

    def read_all(self) -> Dict[str, ColegioData]:
        colegios = {}
        for sheet_name in self.wb.sheetnames:
            ws = self.wb[sheet_name]
            try:
                # Detectar celdas fijas dinámicamente
                self.celdas_fijas = detectar_celdas_fijas_certificado(ws)
                if not self.celdas_fijas.get('codigo_dane'):
                    logger.warning(f"Hoja '{sheet_name}': no se encontró DANE, omitida")
                    continue
                colegio = self._read_sheet(ws, sheet_name)
                if colegio and colegio.codigo_dane:
                    colegios[colegio.codigo_dane] = colegio
                    logger.info(f"Leído: {colegio.nombre} (DANE: {colegio.codigo_dane}) Hoja: {sheet_name}")
                else:
                    logger.warning(f"Hoja '{sheet_name}' sin DANE válido, omitida")
            except Exception as e:
                logger.error(f"Error leyendo hoja '{sheet_name}': {e}")
        return colegios

    def _read_sheet(self, ws, sheet_name: str) -> Optional[ColegioData]:
        # Leer celdas fijas detectadas dinámicamente
        datos_fijos = {}
        for key, cell in self.celdas_fijas.items():
            if key == 'codigo_dane_valor':
                continue  # Skip internal key
            val = ws[cell].value
            datos_fijos[key] = str(val).strip() if val else ""

        # Use extracted DANE value if available, otherwise read from cell
        if 'codigo_dane_valor' in self.celdas_fijas:
            datos_fijos['codigo_dane'] = self.celdas_fijas['codigo_dane_valor']
        elif not datos_fijos.get('codigo_dane'):
            return None

        if not datos_fijos.get('codigo_dane'):
            return None

        # Detectar tabla de raciones
        header_row = detectar_fila_encabezado_tabla(ws)
        if not header_row:
            logger.warning(f"Hoja '{sheet_name}': no se encontró encabezado tabla")
            return None

        # Detectar columnas dinámicamente
        cols = self._detectar_columnas_tabla(ws, header_row)
        if not cols:
            return None

        # Extraer filas hasta TOTAL
        total_row = detectar_fila_total(ws, header_row + 1)
        if not total_row:
            total_row = ws.max_row

        raciones = defaultdict(lambda: defaultdict(dict))
        current_tipo = None

        for row in range(header_row + 1, total_row + 1):
            tipo_val = ws.cell(row=row, column=CERT_TABLA_COLS['tipo']).value
            clasif_val = ws.cell(row=row, column=CERT_TABLA_COLS['clasificacion']).value
            rac_dia_val = ws.cell(row=row, column=CERT_TABLA_COLS['raciones_dia']).value
            dias_val = ws.cell(row=row, column=CERT_TABLA_COLS['dias']).value
            total_val = ws.cell(row=row, column=CERT_TABLA_COLS['total']).value

            # Detectar tipo de ración
            if tipo_val:
                tipo_str = str(tipo_val).strip()
                tipo_norm = normalizar_tipo_racion(tipo_str)
                modalidad = normalizar_modalidad(tipo_str)
                if tipo_norm:
                    current_tipo = tipo_norm

            if not current_tipo or not clasif_val:
                continue

            # Detectar nivel
            clasif_str = str(clasif_val).strip().upper()
            nivel = None
            for n in ["A", "B", "C", "D", "E"]:
                if f"NIVEL {n}" in clasif_str or f"NIVEL {n}-" in clasif_str:
                    nivel = n
                    break

            if not nivel:
                continue

            rac_dia = int(rac_dia_val) if rac_dia_val else 0
            dias = int(dias_val) if dias_val else 0
            total = int(total_val) if total_val else 0

            if rac_dia > 0 or dias > 0 or total > 0:
                raciones[current_tipo][nivel] = {
                    'raciones_dia': rac_dia,
                    'dias': dias,
                    'total': total,
                }

        return ColegioData(
            hoja=ws.title,
            nombre=datos_fijos['institucion'],
            codigo_dane=datos_fijos['codigo_dane'],
            departamento=datos_fijos['departamento'],
            municipio=datos_fijos['municipio'],
            fecha_desde=datos_fijos['fecha_desde'],
            fecha_hasta=datos_fijos['fecha_hasta'],
            rector=datos_fijos['rector'],
            raciones=dict(raciones)
        )

    def _detectar_columnas_tabla(self, ws, header_row: int) -> Optional[Dict]:
        """Detecta columnas de la tabla por encabezados en sub-header row."""
        sub_header = header_row + 1
        cols = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=sub_header, column=col).value
            if val:
                v = str(val).upper()
                if 'RACIONES' in v and 'DÍA' in v:
                    cols['raciones_dia'] = col
                elif 'DÍAS' in v and 'ATEND' in v:
                    cols['dias'] = col
                elif 'TOTAL' in v and 'RACIONES' in v and 'ENTREG' in v:
                    cols['total'] = col
        return cols if all(k in cols for k in ['raciones_dia', 'dias', 'total']) else None


class CoberturaWriter:
    """Escribe en plantilla Cobertura (4 bloques) preservando fórmulas y formato."""

    def __init__(self, template_path: str):
        self.template_path = Path(template_path)
        self.wb = openpyxl.load_workbook(template_path)
        self.ws = self.wb.active
        # Detectar celdas combinadas
        self._merged_ranges = set()
        for mr in self.ws.merged_cells.ranges:
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    self._merged_ranges.add((r, c))

    def _safe_write(self, row: int, col: int, value: Any):
        if (row, col) not in self._merged_ranges:
            self.ws.cell(row=row, column=col, value=value)

    def _write_formula(self, row: int, col: int, formula: str):
        if (row, col) not in self._merged_ranges:
            self.ws.cell(row=row, column=col, value=formula)

    def escribir_registro(self, bloque: int, fila: int, nivel: str, 
                          am: int = 0, pm: int = 0, cob: int = 0,
                          dias: int = 0, es_pm: bool = False) -> Dict:
        """Escribe un registro en el bloque especificado."""
        b = BLOQUES[bloque]
        cols = b["columnas_por_nivel"][nivel]
        data = {'fila': fila, 'bloque': bloque, 'nivel': nivel}

        if bloque in [1, 2]:  # Bloques con AM/PM
            if not es_pm:
                if am > 0:
                    self._safe_write(fila, cols['AM'], am)
                    data['AM'] = am
            else:
                if pm > 0:
                    self._safe_write(fila, cols['PM'], pm)
                    data['PM'] = pm
            if dias > 0:
                self._safe_write(fila, cols['DIAS'], dias)
                data['DIAS'] = dias
            # Fórmulas
            self._write_formula(fila, cols['TOTAL_COB'], f"={self._col_letter(cols['AM'])}{fila}+{self._col_letter(cols['PM'])}{fila}")
            self._write_formula(fila, cols['TOTAL_RAC'], f"={self._col_letter(cols['TOTAL_COB'])}{fila}*{self._col_letter(cols['DIAS'])}{fila}")
            self._write_formula(fila, cols['VALOR'], f"={self._col_letter(cols['TOTAL_RAC'])}{fila}*{b['tarifas'][nivel]}")
        else:  # Bloques 3, 4 - solo COBERTURA
            if cob > 0:
                self._safe_write(fila, cols['COB'], cob)
                data['COB'] = cob
            if dias > 0:
                self._safe_write(fila, cols['DIAS'], dias)
                data['DIAS'] = dias
            self._write_formula(fila, cols['TOTAL_RAC'], f"={self._col_letter(cols['COB'])}{fila}*{self._col_letter(cols['DIAS'])}{fila}")
            self._write_formula(fila, cols['VALOR'], f"={self._col_letter(cols['TOTAL_RAC'])}{fila}*{b['tarifas'][nivel]}")

        data['fila'] = fila
        return data

    def _col_letter(self, col: int) -> str:
        """Convierte número de columna a letra (A, B, ..., Z, AA, AB...)."""
        result = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def save(self, output_path: str):
        self.wb.save(output_path)
        logger.info(f"Guardado: {output_path}")

    def regenerar_totales_bloque(self, bloque: int):
        """Regenera fórmulas SUM de totales del bloque con rangos correctos."""
        b = BLOQUES[bloque]
        start = b['fila_inicio']
        end = b['fila_fin']
        total_row = b['total_fila']

        for nivel in b['niveles']:
            cols = b['columnas_por_nivel'][nivel]
            col_letter = self._col_letter
            # Total Cobertura
            if 'TOTAL_COB' in cols:
                self._write_formula(total_row, cols['TOTAL_COB'],
                    f"=SUM({col_letter(cols['AM'])}{start}:{col_letter(cols['AM'])}{end})")
            # Total Rac
            self._write_formula(total_row, cols['TOTAL_RAC'],
                f"=SUM({col_letter(cols['TOTAL_RAC'])}{start}:{col_letter(cols['TOTAL_RAC'])}{end})")
            # Valor
            if 'VALOR' in cols:
                self._write_formula(b['total_valor_fila'], b['valor_cols'][nivel],
                    f"={self._col_letter(b['columnas_por_nivel'][nivel]['TOTAL_RAC'])}{total_row}*{b['tarifas'][nivel]}")


class PAEEngine:
    """Orquesta el proceso completo: extraer → generar."""

    def __init__(self, certificado_path: str, cobertura_path: str,
                 mapeo_path: str = None, tarifas_path: str = None):
        self.reader = CertificadoReader(certificado_path)
        self.writer = CoberturaWriter(cobertura_path)
        self.cert_path = Path(certificado_path)
        self.cob_path = Path(cobertura_path)
        self.mapeo_path = Path(mapeo_path) if mapeo_path else MAPEO_FILE
        self.tarifas_path = Path(tarifas_path) if tarifas_path else TARIFAS_FILE
        self.mapeo = self._cargar_mapeo()
        self.colegios: Dict[str, ColegioData] = {}
        self.registros_borrador: List[RegistroBorrador] = []

    def _cargar_mapeo(self) -> Dict:
        """Carga mapeo DANE → (nombre_cobertura, bloque, fila_am, fila_pm)."""
        mapeo = {}
        if self.mapeo_path.exists():
            try:
                df = pd.read_csv(self.mapeo_path, dtype={'dane': str})
                for _, row in df.iterrows():
                    mapeo[str(row['dane'])] = {
                        'nombre_cobertura': row.get('nombre_cobertura', ''),
                        'bloque': int(row.get('bloque', 0)),
                        'fila_am': int(row.get('fila_am', 0)) if row.get('fila_am') else 0,
                        'fila_pm': int(row.get('fila_pm', 0)) if row.get('fila_pm') else 0,
                    }
                logger.info(f"Mapeo cargado: {len(mapeo)} entradas")
            except Exception as e:
                logger.warning(f"Error cargando mapeo: {e}")
        return mapeo

    def extraer(self) -> List[RegistroBorrador]:
        """Paso A: Extrae datos y genera borrador para revisión manual."""
        logger.info("=== PASO A: EXTRAER ===")
        self.colegios = self.reader.read_all()
        logger.info(f"Total colegios leídos: {len(self.colegios)}")

        self.registros_borrador = []

        for dane, colegio in self.colegios.items():
            # Obtener mapeo del colegio
            m = self.mapeo.get(dane, {})
            fila_am_base = m.get('fila_am', 0)
            fila_pm_base = m.get('fila_pm', 0)
            nombre_cobertura = m.get('nombre_cobertura', colegio.nombre)
            bloque_mapa = m.get('bloque', 0)

            # Procesar cada tipo de ración
            for tipo_racion, niveles in colegio.raciones.items():
                # Detectar modalidad para este tipo
                modalidad = "PS"  # default
                # Detectar modalidad desde el texto de la hoja (simplificado)
                # TODO: Mejorar detección de modalidad desde texto original
                
                # Determinar bloque destino
                bloque = get_bloque_por_tipo_modalidad(tipo_racion, "PS")
                if not bloque:
                    logger.warning(f"Sin bloque para {tipo_racion} DANE={dane}")
                    continue

                bloque_info = BLOQUES[bloque]
                bloque_nombre = bloque_info['nombre']

                # Obtener filas base del mapeo
                fila_am_base = self.mapeo.get(dane, {}).get('fila_am', 0)
                fila_pm_base = self.mapeo.get(dane, {}).get('fila_pm', 0)

                # Procesar cada nivel
                for nivel in ['A', 'B', 'C', 'D', 'E']:
                    # Saltar nivel E en bloques 1 y 2
                    if bloque in [1, 2] and nivel == 'E':
                        continue
                    if bloque in [3, 4] and nivel not in ['A', 'B', 'C', 'D', 'E']:
                        continue

                    # Obtener datos de CAJM y CAJT
                    cajm = colegio.raciones.get('CAJM', {}).get(nivel, {})
                    cajt = colegio.raciones.get('CAJT', {}).get(nivel, {})
                    alm = colegio.raciones.get('ALMUERZO', {}).get(nivel, {})

                    # Determinar qué tipo de ración tiene datos para este nivel
                    tiene_cajm = nivel in colegio.raciones.get('CAJM', {}) and colegio.raciones['CAJM'][nivel].get('raciones_dia', 0) > 0
                    tiene_cajt = nivel in colegio.raciones.get('CAJT', {}) and colegio.raciones['CAJT'][nivel].get('raciones_dia', 0) > 0
                    tiene_alm = nivel in colegio.raciones.get('ALMUERZO', {}) and colegio.raciones['ALMUERZO'][nivel].get('raciones_dia', 0) > 0

                    if not (tiene_cajm or tiene_cajt or tiene_alm):
                        continue

                    # Determinar filas AM/PM según mapeo
                    fila_am = 0
                    fila_pm = 0
                    es_pm = False
                    usa_dos_filas = False

                    # Obtener filas del mapeo
                    fila_am = self.mapeo.get(dane, {}).get('fila_am', 0)
                    fila_pm = self.mapeo.get(dane, {}).get('fila_pm', 0)

                    # Determinar días atendidos
                    dias = 0
                    if tiene_cajm:
                        dias = cajm.get('dias', 0)
                    if tiene_cajt:
                        dias = max(dias, cajt.get('dias', 0))
                    if tiene_alm:
                        dias = max(dias, alm.get('dias', 0))

                    # Determinar si usa una o dos filas (regla: si días CAJM != días CAJT)
                    dias_cajm = cajm.get('dias', 0) if tiene_cajm else 0
                    dias_cajt = cajt.get('dias', 0) if tiene_cajt else 0
                    usa_dos_filas = tiene_cajm and tiene_cajt and dias_cajm != dias_cajt

                    if bloque in [1, 2]:  # Bloques con AM/PM
                        if usa_dos_filas:
                            # Usar fila_am para CAJM, fila_pm para CAJT
                            if tiene_cajm:
                                raciones_dia = cajm.get('raciones_dia', 0)
                                dias = cajm.get('dias', 0)
                                self.registros_borrador.append(RegistroBorrador(
                                    dane=colegio.codigo_dane,
                                    hoja_certificado="",
                                    nombre_certificado="",
                                    nombre_cobertura="",
                                    bloque=bloque,
                                    bloque_nombre=BLOQUES[bloque]['nombre'],
                                    fila_am=fila_am,
                                    fila_pm=0,
                                    tipo_racion="CAJM",
                                    modalidad="PS",
                                    nivel=nivel,
                                    raciones_dia=cajm.get('raciones_dia', 0),
                                    dias=dias,
                                    total_raciones=cajm.get('total', 0),
                                    estado="ok",
                                    observaciones=""
                                ))
                            if tiene_cajt:
                                self.registros_borrador.append(RegistroBorrador(
                                    dane=colegio.codigo_dane,
                                    hoja_certificado="",
                                    nombre_certificado="",
                                    nombre_cobertura="",
                                    bloque=bloque,
                                    bloque_nombre=BLOQUES[bloque]['nombre'],
                                    fila_am=0,
                                    fila_pm=fila_pm,
                                    tipo_racion="CAJT",
                                    modalidad="PS",
                                    nivel=nivel,
                                    raciones_dia=cajt.get('raciones_dia', 0),
                                    dias=cajt.get('dias', 0),
                                    total_raciones=cajt.get('total', 0),
                                    estado="ok",
                                    observaciones=""
                                ))
                        else:
                            # Una sola fila con AM y PM
                            if tiene_cajm or tiene_cajt:
                                raciones_dia_am = cajm.get('raciones_dia', 0) if tiene_cajm else 0
                                raciones_dia_pm = cajt.get('raciones_dia', 0) if tiene_cajt else 0
                                dias = max(dias_cajm, dias_cajt) if (dias_cajm := cajm.get('dias', 0) if tiene_cajm else 0) or (dias_cajt := cajt.get('dias', 0) if tiene_cajt else 0) else 0
                                self.registros_borrador.append(RegistroBorrador(
                                    dane=colegio.codigo_dane,
                                    hoja_certificado="",
                                    nombre_certificado="",
                                    nombre_cobertura="",
                                    bloque=bloque,
                                    bloque_nombre=BLOQUES[bloque]['nombre'],
                                    fila_am=fila_am,
                                    fila_pm=fila_pm if not usa_dos_filas else 0,
                                    tipo_racion="CAJM/CAJT",
                                    modalidad="PS",
                                    nivel=nivel,
                                    raciones_dia=raciones_dia_am + raciones_dia_pm,
                                    dias=dias,
                                    total_raciones=(cajm.get('total', 0) if tiene_cajm else 0) + (cajt.get('total', 0) if tiene_cajt else 0),
                                    estado="ok",
                                    observaciones=""
                                ))
                    else:  # Bloques 3, 4 - solo COBERTURA
                        if tiene_alm:
                            raciones_dia = alm.get('raciones_dia', 0)
                            dias = alm.get('dias', 0)
                            self.registros_borrador.append(RegistroBorrador(
                                dane=colegio.codigo_dane,
                                hoja_certificado="",
                                nombre_certificado="",
                                nombre_cobertura="",
                                bloque=bloque,
                                bloque_nombre=BLOQUES[bloque]['nombre'],
                                fila_am=fila_am,
                                fila_pm=0,
                                tipo_racion="ALMUERZO",
                                modalidad="PS",
                                nivel=nivel,
                                raciones_dia=raciones_dia,
                                dias=dias,
                                total_raciones=alm.get('total', 0),
                                estado="ok",
                                observaciones=""
                            ))

        return self.registros_borrador

    def generar_borrador_excel(self, output_path: str):
        """Genera Excel borrador para revisión manual."""
        df = pd.DataFrame([asdict(r) for r in self.registros_borrador])
        df.to_excel(output_path, index=False)
        logger.info(f"Borrador generado: {output_path}")

    def procesar(self, output_path: str = None, log_path: str = None):
        """Procesa el certificado y cobertura generando el borrador."""
        self.extraer()
        if output_path is None:
            output_path = "output/borrador_revision.xlsx"
        self.generar_borrador_excel(output_path)
        if log_path is None:
            log_path = "logs/borrador_revision.log"
        logger.info(f"Procesamiento completado. Output: {output_path}")

    def save_output(self, output_path: str):
        """Guarda el output Excel."""
        self.generar_borrador_excel(output_path)

    def save_log(self, log_path: str):
        """Guarda el log JSON."""
        import json
        from pathlib import Path
        log_data = {
            'colegios_procesados': len(self.colegios),
            'filas_escritas': len(self.registros_borrador),
            'colegios_sin_cobertura': [
                dane for dane in self.colegios.keys()
                if not any(r.dane == dane for r in self.registros_borrador)
            ]
        }
        Path(log_path).parent.mkdir(parents=True, exist_ok=True)
        with open(log_path, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)
        logger.info(f"Log guardado: {log_path}")

    @property
    def resultado(self):
        """Propiedad que devuelve los registros de proceso para reportes."""
        return self.registros_borrador

    def get_preview(self):
        """Devuelve un dataframe con la vista previa de los datos procesados."""
        import pandas as pd
        if not self.registros_borrador:
            return pd.DataFrame()
        
        # Construir dataframe desde los registros
        data = []
        for r in self.registros_borrador:
            row = {
                'fila': r.fila_am,  # Usar fila_am como fila principal
                'dane': r.dane,
                'nombre': r.nombre_cobertura,
                'bloque': r.bloque,
                'nivel': r.nivel,
                'tipo_racion': r.tipo_racion,
                'modalidad': r.modalidad,
            }
            # Agregar datos de ración por nivel (AM/PM/Días)
            row['AM'] = r.raciones_dia  # Usar raciones_dia como valor AM
            row['PM'] = 0  # Valor por defecto para PM
            row['Días'] = r.dias  # Usar dias como días
            
            data.append(row)
        
        df = pd.DataFrame(data)
        if df.empty:
            return df
        
        # Filtrar filas con datos
        data_cols = [c for c in df.columns if any(x in c for x in ['AM', 'PM', 'Días'])]
        if data_cols:
            df['_has'] = df[data_cols].notna().any(axis=1)
            show = df[df['_has']].drop(columns=['_has'])
        else:
            show = df
        
        return show
    
    def get_log_data(self):
        """Devuelve diccionario con estadísticas del log para métricas."""
        colegios_sin_cobertura = [
            dane for dane in self.colegios.keys()
            if not any(r.dane == dane for r in self.registros_borrador)
        ]
        return {
            'colegios_procesados': len(self.colegios),
            'filas_escritas': len(self.registros_borrador),
            'colegios_sin_cobertura': colegios_sin_cobertura
        }


def main():
    import click

    @click.group()
    def cli():
        """PAE Automatización: Certificado → Cobertura (extraer → generar)"""
        pass

    @cli.command()
    @click.argument('certificado', type=click.Path(exists=True))
    @click.argument('cobertura', type=click.Path(exists=True))
    @click.option('--mapeo', default=str(MAPEO_FILE), help='CSV mapeo colegios')
    @click.option('--output', default='borrador_revision.xlsx', help='Excel borrador salida')
    def extraer(certificado, cobertura, mapeo, output):
        """Paso A: Extrae datos y genera borrador para revisión manual."""
        engine = PAEEngine(certificado, cobertura, mapeo)
        engine.extraer()
        engine.generar_borrador_excel(output)
        click.echo(f"Borrador generado: {output}")
        click.echo("Revise, corrija y complete el Excel, luego ejecute 'generar'.")

    @cli.command()
    @click.argument('borrador', type=click.Path(exists=True))
    @click.argument('cobertura', type=click.Path(exists=True))
    @click.option('--output', default='COBERTURA_FINAL.xlsx', help='Archivo final')
    def generar(borrador, cobertura, output):
        """Paso B: Genera Cobertura final desde borrador corregido."""
        click.echo("Funcionalidad pendiente: leer borrador corregido y escribir Cobertura final")

    cli()


if __name__ == '__main__':
    main()