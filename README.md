# PAE Automatización - Certificado → Cobertura

Herramienta en Python para automatizar el llenado de la **Plantilla Cobertura** a partir de la **Plantilla Certificado** para el programa PAE (Alimentación Escolar) de la UT Alianza Integral.

## Estructura del proyecto

```
pae_automatizacion/
├── pae.py                    # Punto de entrada CLI
├── tarifas.csv               # Tabla de tarifas parametrizable
├── colegios_tarifas.csv      # Mapeo colegio → grupo de tarifa (opcional)
├── requirements.txt
├── README.md
└── src/
    └── pae_automatizacion/
        ├── __init__.py
        └── pae_automatizador.py
```

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

### 1. Generar borrador

```bash
python pae.py generar "2_CERTIFICACIONES_MES_DE_JULIO.xlsx" "3_COBERTURA_EJECUTADA.xlsx" --output cobertura_borrador.xlsx
```

Esto:
- Lee todas las hojas del Certificado
- Extrae datos por colegio (usando código DANE como llave)
- Aplica la regla de 1 vs 2 filas (según días CAJM vs CAJT)
- Escribe en la plantilla Cobertura **solo valores de entrada** (AM, PM, días)
- **Preserva fórmulas y formato** existentes
- Genera un archivo `cobertura_borrador.xlsx` para revisión manual
- Genera un log `cobertura_borrador_log.json`

### 2. Revisar y corregir en Excel

Abra `cobertura_borrador.xlsx` en Excel y:
- Verifique los valores mapeados automáticamente
- Complete/corrija datos faltantes o incorrectos
- Asigne grupo de tarifa correcto por colegio (columna oculta o auxiliar)
- Guarde los cambios

### 3. Exportar final

```bash
python pae.py confirmar cobertura_borrador.xlsx cobertura_final.xlsx
```

### 4. Solo vista previa (sin guardar)

```bash
python pae.py preview "2_CERTIFICACIONES_MES_DE_JULIO.xlsx" "3_COBERTURA_EJECUTADA.xlsx"
```

## Archivos de configuración

### tarifas.csv

Tabla de tarifas por grupo y nivel. **No modificar los valores numéricos sin confirmación.**

```csv
grupo,nivel,tarifa
grupo_1,A,4298
grupo_1,B,4685
grupo_1,C,5039
grupo_1,D,5638
grupo_2,A,4517
grupo_2,B,4905
grupo_2,C,5258
grupo_2,D,5856
...
```

### colegios_tarifas.csv (opcional)

Mapeo de código DANE → grupo de tarifa. Si existe, se usa automáticamente.

```csv
codigo_dane,grupo_tarifa
12345678,grupo_1
87654321,grupo_2
```

## Reglas de mapeo

| Certificado | → | Cobertura |
|-------------|---|-----------|
| CAJM Nivel A-D (raciones/día) | → | Columna AM (C, H, M, R) |
| CAJT Nivel A-D (raciones/día) | → | Columna PM (D, I, N, S) |
| Días atendidos CAJM/CAJT | → | Columna Días (F, K, P, U) |

**Regla 1 vs 2 filas:**
- Si días CAJM == días CAJT → **1 fila** (AM y PM en misma fila, comparten columna días)
- Si días CAJM != días CAJT → **2 filas** (separadas, cada una con su columna días)

## Detección automática

- Busca encabezados "TIPO RACIÓN", "NIVEL", "RACIONES/DÍA", "DÍAS ATENDIDOS", "TOTAL RACIONES" en cada hoja
- No asume fila fija final (maneja sedes adicionales)
- Usa código DANE (celda I7) como llave principal
- Alerta si DANE no coincide entre plantillas

## Logs y reportes

Cada ejecución genera:
- `cobertura_borrador_log.json`: detalle de mapeos, advertencias, colegios sin coincidencia
- Consola: resumen de procesamiento

## Preguntas pendientes (confirmar con usuario)

1. **Criterio de tarifa**: ¿Qué define grupo_1 vs grupo_2 vs otras variantes? (zona, contrato, modalidad)
2. **Almuerzo Jornada Única**: ¿Se incluye en otro reporte o se descarta?
3. **Plantilla Cobertura base**: ¿Viene pre-creada cada mes o hay que generarla?
4. **Interfaz**: ¿Es suficiente Excel intermedio o se necesita GUI?

## Dependencias

- `openpyxl` - Lectura/escritura Excel preservando fórmulas y formato
- `pandas` - Manipulación y previsualización de datos
- `click` - CLI
- `tabulate` - Tablas en consola