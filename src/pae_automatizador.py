import openpyxl
import pandas as pd
import csv
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, asdict
from collections import defaultdict

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


@dataclass
class ColegioData:
    nombre: str
    codigo_dane: str
    departamento: str
    municipio: str
    fecha_desde: str
    fecha_hasta: str
    rector: str
    raciones: Dict[str, Dict[str, Dict[str, int]]]


@dataclass
class CoberturaRow:
    item: int
    dane: str
    nombre: str
    nivel_a_am: int = 0
    nivel_a_pm: int = 0
    nivel_a_dias: int = 0
    nivel_b_am: int = 0
    nivel_b_pm: int = 0
    nivel_b_dias: int = 0
    nivel_c_am: int = 0
    nivel_c_pm: int = 0
    nivel_c_dias: int = 0
    nivel_d_am: int = 0
    nivel_d_pm: int = 0
    nivel_d_dias: int = 0
    grupo_tarifa: str = "grupo_1"
    advertencia: str = ""


class TarifasManager:
    def __init__(self, tarifas_path: str = "tarifas.csv", colegios_tarifas_path: str = "colegios_tarifas.csv"):
        self.tarifas_path = Path(tarifas_path)
        self.colegios_tarifas_path = Path(colegios_tarifas_path)
        self.tarifas: Dict[str, Dict[str, int]] = {}
        self.colegio_grupo: Dict[str, str] = {}
        self.load_tarifas()
        self.load_colegios_tarifas()

    def load_tarifas(self):
        if not self.tarifas_path.exists():
            logger.warning(f"Archivo de tarifas no encontrado: {self.tarifas_path}")
            return
        df = pd.read_csv(self.tarifas_path)
        for _, row in df.iterrows():
            grupo = row['grupo']
            nivel = row['nivel']
            tarifa = int(row['tarifa'])
            if grupo not in self.tarifas:
                self.tarifas[grupo] = {}
            self.tarifas[grupo][nivel] = tarifa

    def load_colegios_tarifas(self):
        if not self.colegios_tarifas_path.exists():
            logger.info(f"Archivo de mapeo colegios-tarifas no encontrado: {self.colegios_tarifas_path} (opcional)")
            return
        try:
            df = pd.read_csv(self.colegios_tarifas_path, comment='#')
            if 'codigo_dane' not in df.columns or 'grupo_tarifa' not in df.columns:
                logger.warning(f"Columnas esperadas no encontradas en {self.colegios_tarifas_path}. Columnas: {list(df.columns)}")
                return
            for _, row in df.iterrows():
                dane = str(row['codigo_dane']).strip()
                grupo = str(row['grupo_tarifa']).strip()
                if dane and grupo:
                    self.colegio_grupo[dane] = grupo
            logger.info(f"Mapeo colegios-tarifas cargado: {len(self.colegio_grupo)} entradas")
        except Exception as e:
            logger.warning(f"Error leyendo {self.colegios_tarifas_path}: {e}")

    def get_tarifa(self, grupo: str, nivel: str) -> Optional[int]:
        return self.tarifas.get(grupo, {}).get(nivel)

    def get_grupo_colegio(self, dane: str) -> str:
        return self.colegio_grupo.get(dane, "grupo_1")

    def get_grupos(self) -> List[str]:
        return list(self.tarifas.keys())

    def get_niveles(self, grupo: str) -> List[str]:
        return list(self.tarifas.get(grupo, {}).keys())

    def get_tarifa(self, grupo: str, nivel: str) -> Optional[int]:
        return self.tarifas.get(grupo, {}).get(nivel)

    def get_grupos(self) -> List[str]:
        return list(self.tarifas.keys())

    def get_niveles(self, grupo: str) -> List[str]:
        return list(self.tarifas.get(grupo, {}).keys())


class CertificadoReader:
    CELDAS_FIJAS = {
        'nombre': 'B7',
        'codigo_dane': 'I7',
        'departamento': 'B8',
        'municipio': 'B9',
        'fecha_desde': 'C10',
        'fecha_hasta': 'I10',
        'rector': 'B11',
    }

    TIPOS_RACION = ['CAJM', 'CAJT', 'ALMUERZO']
    NIVELES = ['A', 'B', 'C', 'D', 'E']

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
                elif 'ALMUERZO' in tipo_str:
                    current_tipo = 'ALMUERZO'
                elif 'JORNADA' in tipo_str:
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

            if nivel and nivel in self.NIVELES:
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
    NIVEL_COLS = {
        'A': {'am': 3, 'pm': 4, 'total_cob': 5, 'dias': 6, 'total_rac': 7, 'valor': 23},
        'B': {'am': 8, 'pm': 9, 'total_cob': 10, 'dias': 11, 'total_rac': 12, 'valor': 24},
        'C': {'am': 13, 'pm': 14, 'total_cob': 15, 'dias': 16, 'total_rac': 17, 'valor': 25},
        'D': {'am': 18, 'pm': 19, 'total_cob': 20, 'dias': 21, 'total_rac': 22, 'valor': 26},
    }
    VALOR_TOTAL_COL = 27
    DATA_START_ROW = 4
    NAME_COL = 2  # Column B has school names in the template
    DANE_COL = 28  # Column AB for DANE code (new)

    def __init__(self, template_path: str, tarifas_mgr: TarifasManager):
        self.template_path = Path(template_path)
        self.tarifas_mgr = tarifas_mgr
        self.wb = openpyxl.load_workbook(template_path)
        self.ws = self.wb.active
        self.name_to_row: Dict[str, int] = {}
        self._build_name_map()

    def _normalize_name(self, nombre: str) -> str:
        n = nombre.upper()
        n = n.replace('INSTITUCION EDUCATIVA', '')
        n = n.replace('INSTITUCION E.', '')
        n = n.replace('INSTITUCION', '')
        n = n.replace('I.E.', '')
        n = n.replace('I.E', '')
        n = n.replace('INST.', '')
        n = n.replace('INST', '')
        n = n.replace('SEDE', '')
        n = n.replace('Nº', '')
        n = n.replace('NO.', '')
        n = n.replace('.', '')
        n = n.replace('  ', ' ')
        return n.strip()

    def _build_name_map(self):
        for row in range(self.DATA_START_ROW, self.ws.max_row + 1):
            name_cell = self.ws.cell(row=row, column=self.NAME_COL)
            if name_cell.value:
                nombre = str(name_cell.value).strip()
                if nombre:
                    norm = self._normalize_name(nombre)
                    if norm:
                        self.name_to_row[norm] = row

    def find_or_create_row(self, dane: str, nombre: str) -> Tuple[int, bool]:
        norm_nombre = self._normalize_name(nombre)

        # Try exact normalized match
        if norm_nombre in self.name_to_row:
            row = self.name_to_row[norm_nombre]
            # Write DANE to DANE_COL for future reference
            self.ws.cell(row=row, column=self.DANE_COL, value=dane)
            return row, False

        # Try fuzzy match (substring)
        for norm_name, row in self.name_to_row.items():
            if norm_nombre in norm_name or norm_name in norm_nombre:
                self.ws.cell(row=row, column=self.DANE_COL, value=dane)
                return row, True

        # Create new row at the end
        new_row = self.ws.max_row + 1
        self.ws.cell(row=new_row, column=1, value=new_row - self.DATA_START_ROW + 1)
        self.ws.cell(row=new_row, column=self.NAME_COL, value=nombre)
        self.ws.cell(row=new_row, column=self.DANE_COL, value=dane)
        self.name_to_row[norm_nombre] = new_row
        return new_row, True

    def write_colegio(self, colegio: ColegioData) -> List[CoberturaRow]:
        rows_written = []

        cajm = colegio.raciones.get('CAJM', {})
        cajt = colegio.raciones.get('CAJT', {})

        grupo_tarifa = self.tarifas_mgr.get_grupo_colegio(colegio.codigo_dane)
        colegio.raciones['_grupo_tarifa'] = grupo_tarifa

        niveles_con_datos = set()
        for nivel in ['A', 'B', 'C', 'D']:
            if nivel in cajm and cajm[nivel].get('raciones_dia', 0) > 0:
                niveles_con_datos.add(nivel)
            if nivel in cajt and cajt[nivel].get('raciones_dia', 0) > 0:
                niveles_con_datos.add(nivel)

        if not niveles_con_datos:
            logger.warning(f"Colegio {colegio.nombre} sin datos de CAJM/CAJT para niveles A-D")
            return rows_written

        dias_cajm = set()
        dias_cajt = set()
        for nivel in niveles_con_datos:
            if nivel in cajm:
                dias_cajm.add(cajm[nivel].get('dias', 0))
            if nivel in cajt:
                dias_cajt.add(cajt[nivel].get('dias', 0))

        dias_cajm = dias_cajm - {0}
        dias_cajt = dias_cajt - {0}

        usar_dos_filas = len(dias_cajm) > 0 and len(dias_cajt) > 0 and dias_cajm != dias_cajt

        if usar_dos_filas:
            row1, _ = self.find_or_create_row(colegio.codigo_dane, colegio.nombre)
            row2 = self.ws.max_row + 1
            self._copy_row_format(row1, row2)
            self.ws.cell(row=row2, column=1, value=row2 - self.DATA_START_ROW + 1)
            self.ws.cell(row=row2, column=2, value=colegio.codigo_dane)
            self.ws.cell(row=row2, column=3, value=f"{colegio.nombre} (TARDE)")

            rows_written.append(self._write_fila_am(row1, cajm, colegio))
            rows_written.append(self._write_fila_pm(row2, cajt, colegio))
        else:
            row, _ = self.find_or_create_row(colegio.codigo_dane, colegio.nombre)
            dias_comunes = (dias_cajm | dias_cajt) - {0}
            dias_val = list(dias_comunes)[0] if dias_comunes else 0

            row_data = self._write_fila_completa(row, cajm, cajt, dias_val, colegio)
            rows_written.append(row_data)

        return rows_written

    def _write_fila_am(self, row: int, cajm: Dict, colegio: ColegioData) -> CoberturaRow:
        row_data = CoberturaRow(
            item=self.ws.cell(row=row, column=1).value or 0,
            dane=colegio.codigo_dane,
            nombre=colegio.nombre,
            grupo_tarifa=colegio.raciones.get('_grupo_tarifa', 'grupo_1')
        )
        for nivel in ['A', 'B', 'C', 'D']:
            cols = self.NIVEL_COLS[nivel]
            if nivel in cajm:
                rac_dia = cajm[nivel].get('raciones_dia', 0)
                dias = cajm[nivel].get('dias', 0)
                if rac_dia > 0:
                    self.ws.cell(row=row, column=cols['am'], value=rac_dia)
                    row_data.__dict__[f'nivel_{nivel.lower()}_am'] = rac_dia
                if dias > 0:
                    self.ws.cell(row=row, column=cols['dias'], value=dias)
                    row_data.__dict__[f'nivel_{nivel.lower()}_dias'] = dias
        return row_data

    def _write_fila_pm(self, row: int, cajt: Dict, colegio: ColegioData) -> CoberturaRow:
        row_data = CoberturaRow(
            item=self.ws.cell(row=row, column=1).value or 0,
            dane=colegio.codigo_dane,
            nombre=f"{colegio.nombre} (TARDE)",
            grupo_tarifa=colegio.raciones.get('_grupo_tarifa', 'grupo_1')
        )
        for nivel in ['A', 'B', 'C', 'D']:
            cols = self.NIVEL_COLS[nivel]
            if nivel in cajt:
                rac_dia = cajt[nivel].get('raciones_dia', 0)
                dias = cajt[nivel].get('dias', 0)
                if rac_dia > 0:
                    self.ws.cell(row=row, column=cols['pm'], value=rac_dia)
                    row_data.__dict__[f'nivel_{nivel.lower()}_pm'] = rac_dia
                if dias > 0:
                    self.ws.cell(row=row, column=cols['dias'], value=dias)
                    row_data.__dict__[f'nivel_{nivel.lower()}_dias'] = dias
        return row_data

    def _write_fila_completa(self, row: int, cajm: Dict, cajt: Dict, dias: int, colegio: ColegioData) -> CoberturaRow:
        row_data = CoberturaRow(
            item=self.ws.cell(row=row, column=1).value or 0,
            dane=colegio.codigo_dane,
            nombre=colegio.nombre,
            grupo_tarifa=colegio.raciones.get('_grupo_tarifa', 'grupo_1')
        )
        for nivel in ['A', 'B', 'C', 'D']:
            cols = self.NIVEL_COLS[nivel]
            if nivel in cajm:
                rac_dia = cajm[nivel].get('raciones_dia', 0)
                if rac_dia > 0:
                    self.ws.cell(row=row, column=cols['am'], value=rac_dia)
                    row_data.__dict__[f'nivel_{nivel.lower()}_am'] = rac_dia
            if nivel in cajt:
                rac_dia = cajt[nivel].get('raciones_dia', 0)
                if rac_dia > 0:
                    self.ws.cell(row=row, column=cols['pm'], value=rac_dia)
                    row_data.__dict__[f'nivel_{nivel.lower()}_pm'] = rac_dia
            if dias > 0:
                self.ws.cell(row=row, column=cols['dias'], value=dias)
                row_data.__dict__[f'nivel_{nivel.lower()}_dias'] = dias
        return row_data

    def _copy_row_format(self, src_row: int, dst_row: int):
        for col in range(1, self.ws.max_column + 1):
            src_cell = self.ws.cell(row=src_row, column=col)
            dst_cell = self.ws.cell(row=dst_row, column=col)
            if src_cell.has_style:
                dst_cell.font = src_cell.font.copy()
                dst_cell.border = src_cell.border.copy()
                dst_cell.fill = src_cell.fill.copy()
                dst_cell.number_format = src_cell.number_format
                dst_cell.alignment = src_cell.alignment.copy()
                dst_cell.protection = src_cell.protection.copy()

    def save(self, output_path: str):
        self.wb.save(output_path)
        logger.info(f"Guardado: {output_path}")

    def get_preview_data(self) -> List[Dict]:
        preview = []
        for row in range(self.DATA_START_ROW, self.ws.max_row + 1):
            dane = self.ws.cell(row=row, column=2).value
            nombre = self.ws.cell(row=row, column=3).value
            if not dane and not nombre:
                continue
            row_data = {'fila': row, 'dane': dane, 'nombre': nombre}
            for nivel in ['A', 'B', 'C', 'D']:
                cols = self.NIVEL_COLS[nivel]
                row_data[f'nivel_{nivel}_am'] = self.ws.cell(row=row, column=cols['am']).value
                row_data[f'nivel_{nivel}_pm'] = self.ws.cell(row=row, column=cols['pm']).value
                row_data[f'nivel_{nivel}_dias'] = self.ws.cell(row=row, column=cols['dias']).value
                row_data[f'nivel_{nivel}_total_cob'] = self.ws.cell(row=row, column=cols['total_cob']).value
                row_data[f'nivel_{nivel}_total_rac'] = self.ws.cell(row=row, column=cols['total_rac']).value
                row_data[f'nivel_{nivel}_valor'] = self.ws.cell(row=row, column=cols['valor']).value
            row_data['valor_total'] = self.ws.cell(row=row, column=self.VALOR_TOTAL_COL).value
            preview.append(row_data)
        return preview


class PAEAutomatizador:
    def __init__(self, certificado_path: str, cobertura_path: str, tarifas_path: str = "tarifas.csv", colegios_tarifas_path: str = "colegios_tarifas.csv"):
        self.certificado_path = Path(certificado_path)
        self.cobertura_path = Path(cobertura_path)
        self.tarifas_mgr = TarifasManager(tarifas_path, colegios_tarifas_path)
        self.reader = CertificadoReader(certificado_path)
        self.writer = CoberturaWriter(cobertura_path, self.tarifas_mgr)
        self.colegios: Dict[str, ColegioData] = {}
        self.resultado: List[CoberturaRow] = []

    def procesar(self) -> List[CoberturaRow]:
        logger.info("Leyendo plantilla Certificado...")
        self.colegios = self.reader.read_all()
        logger.info(f"Total colegios leídos: {len(self.colegios)}")

        logger.info("Escribiendo en plantilla Cobertura...")
        self.resultado = []
        for dane, colegio in self.colegios.items():
            rows = self.writer.write_colegio(colegio)
            if rows:
                self.resultado.extend(rows)
            else:
                logger.warning(f"No se pudieron escribir datos para {colegio.nombre} (DANE: {dane})")

        # Check for schools in Cobertura without certificado
        for norm_name, row in self.writer.name_to_row.items():
            # This is a bit tricky since we don't have a reverse map easily
            pass

        return self.resultado

    def generar_borrador(self, output_path: str):
        self.procesar()
        self.writer.save(output_path)
        self._exportar_log(output_path.replace('.xlsx', '_log.json'))
        logger.info(f"Borrador generado: {output_path}")

    def exportar_final(self, borrador_path: str, output_path: str):
        wb = openpyxl.load_workbook(borrador_path)
        wb.save(output_path)
        logger.info(f"Archivo final exportado: {output_path}")

    def _exportar_log(self, log_path: str):
        # Get school names from Cobertura that weren't matched
        certificado_names = {self.writer._normalize_name(c.nombre) for c in self.colegios.values()}
        cobertura_names = set(self.writer.name_to_row.keys())
        sin_certificado = cobertura_names - certificado_names

        log_data = {
            'colegios_procesados': len(self.colegios),
            'filas_escritas': len(self.resultado),
            'colegios_sin_cobertura': list(certificado_names - cobertura_names),
            'colegios_sin_certificado': list(sin_certificado),
            'detalle': [asdict(r) for r in self.resultado]
        }
        with open(log_path, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)

    def mostrar_vista_previa(self):
        preview = self.writer.get_preview_data()
        if preview:
            df = pd.DataFrame(preview)
            print("\n=== VISTA PREVIA COBERTURA ===")
            print(df.to_string(index=False))
        else:
            print("No hay datos para mostrar")


def main():
    import click

    @click.group()
    def cli():
        """PAE Automatizacion: Certificado -> Cobertura"""
        pass

    @cli.command()
    @click.argument('certificado', type=click.Path(exists=True))
    @click.argument('cobertura', type=click.Path(exists=True))
    @click.option('--tarifas', default='tarifas.csv', help='Archivo CSV de tarifas')
    @click.option('--colegios-tarifas', default='colegios_tarifas.csv', help='Archivo CSV mapeo colegio->grupo tarifa')
    @click.option('--output', default='cobertura_borrador.xlsx', help='Archivo de salida (borrador)')
    def generar(certificado, cobertura, tarifas, colegios_tarifas, output):
        """Genera borrador de Cobertura a partir de Certificado"""
        automatizador = PAEAutomatizador(certificado, cobertura, tarifas, colegios_tarifas)
        automatizador.generar_borrador(output)
        automatizador.mostrar_vista_previa()
        click.echo(f"\nBorrador generado: {output}")
        click.echo("Revise y corrija el archivo en Excel, luego use 'confirmar' para exportar final.")

    @cli.command()
    @click.argument('borrador', type=click.Path(exists=True))
    @click.argument('output', type=click.Path())
    def confirmar(borrador, output):
        """Confirma borrador y exporta archivo final"""
        wb = openpyxl.load_workbook(borrador)
        wb.save(output)
        click.echo(f"Archivo final exportado: {output}")

    @cli.command()
    @click.argument('certificado', type=click.Path(exists=True))
    @click.argument('cobertura', type=click.Path(exists=True))
    @click.option('--tarifas', default='tarifas.csv', help='Archivo CSV de tarifas')
    @click.option('--colegios-tarifas', default='colegios_tarifas.csv', help='Archivo CSV mapeo colegio->grupo tarifa')
    def preview(certificado, cobertura, tarifas, colegios_tarifas):
        """Muestra vista previa sin guardar"""
        automatizador = PAEAutomatizador(certificado, cobertura, tarifas, colegios_tarifas)
        automatizador.procesar()
        automatizador.mostrar_vista_previa()

    cli()


if __name__ == '__main__':
    main()