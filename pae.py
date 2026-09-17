#!/usr/bin/env python
"""Entry point for PAE Automatizacion CLI - uses new engine"""
import sys
import os
import click
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pae_automatizacion.core.engine import PAEEngine

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
    engine = PAEEngine(certificado, cobertura, tarifas, colegios_tarifas)
    engine.procesar()
    engine.save_output(output)
    engine.save_log(output.replace('.xlsx', '_log.json'))
    click.echo(f"\nBorrador generado: {output}")
    click.echo("Revise y corrija el archivo en Excel.")

@cli.command()
@click.argument('borrador', type=click.Path(exists=True))
@click.argument('output', type=click.Path())
def confirmar(borrador, output):
    """Confirma borrador y exporta archivo final"""
    import openpyxl
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
    engine = PAEEngine(certificado, cobertura, tarifas, colegios_tarifas)
    engine.procesar()
    preview = engine.get_preview()
    if not preview.empty:
        print(preview.to_string(index=False))
    else:
        print("No hay datos para mostrar")

if __name__ == '__main__':
    cli()