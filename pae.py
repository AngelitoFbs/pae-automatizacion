#!/usr/bin/env python
"""Entry point for PAE Automatizacion CLI - extraer -> generar"""
import sys
import os
import click
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pae_automatizacion.core.engine import PAEEngine
from pae_automatizacion.core.config import MAPEO_FILE, TARIFAS_FILE


@click.group()
def cli():
    """PAE Automatizacion: Certificado -> Cobertura (extraer -> generar)"""
    pass


@cli.command()
@click.argument('certificado', type=click.Path(exists=True))
@click.argument('cobertura', type=click.Path(exists=True))
@click.option('--mapeo', default=str(MAPEO_FILE), help='CSV mapeo colegios (dane,nombre_cobertura,bloque,fila_am,fila_pm)')
@click.option('--tarifas', default=str(TARIFAS_FILE), help='CSV tarifas (bloque,nivel,valor) - opcional, usa defaults si no existe')
@click.option('--output', default='borrador_revision.xlsx', help='Excel borrador para revision manual')
def extraer(certificado, cobertura, mapeo, tarifas, output):
    """Paso A: Extrae datos del Certificado y genera borrador para revision manual."""
    engine = PAEEngine(certificado, cobertura, mapeo, tarifas)
    engine.extraer()
    engine.generar_borrador_excel(output)
    print(f"\nBorador generado: {output}")
    print("Revise, corrija y complete el Excel, luego ejecute 'pae generar'.")


@cli.command()
@click.argument('borrador', type=click.Path(exists=True))
@click.argument('cobertura', type=click.Path(exists=True))
@click.option('--mapeo', default='mapeo_colegios.csv', help='CSV mapeo colegios')
@click.option('--tarifas', default='tarifas.csv', help='CSV tarifas')
@click.option('--output', default='COBERTURA_FINAL.xlsx', help='Archivo final Cobertura')
def generar(borrador, cobertura, mapeo, tarifas, output):
    """Paso B: Genera Cobertura final desde borrador corregido."""
    print("Funcionalidad 'generar' en desarrollo. Use el borrador para revision manual.")


@cli.command()
@click.argument('certificado', type=click.Path(exists=True))
@click.argument('cobertura', type=click.Path(exists=True))
@click.option('--mapeo', default='mapeo_colegios.csv', help='CSV mapeo colegios')
@click.option('--tarifas', default='tarifas.csv', help='CSV tarifas')
def preview(certificado, cobertura, mapeo, tarifas):
    """Muestra vista previa de extraccion sin guardar."""
    from pae_automatizacion.core.engine import PAEEngine
    engine = PAEEngine(certificado, cobertura, mapeo, tarifas)
    engine.extraer()
    preview = engine.get_preview()
    if not preview.empty:
        print(preview.to_string(index=False))
    else:
        print("No hay datos para mostrar")


if __name__ == '__main__':
    cli()