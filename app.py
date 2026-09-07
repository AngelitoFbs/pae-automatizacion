import streamlit as st
import pandas as pd
import openpyxl
import json
import tempfile
import os
from pathlib import Path
from typing import Dict, List, Optional
import io

# Import backend logic
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
from pae_automatizador import PAEAutomatizador, CertificadoReader, CoberturaWriter, TarifasManager, ColegioData


st.set_page_config(
    page_title="PAE Automatización - Certificado → Cobertura",
    page_icon="📊",
    layout="wide"
)

st.title("📊 PAE Automatización: Certificado → Cobertura")
st.caption("UT Alianza Integral - Programa de Alimentación Escolar")


# Initialize session state
if 'automatizador' not in st.session_state:
    st.session_state.automatizador = None
if 'colegios_data' not in st.session_state:
    st.session_state.colegios_data = {}
if 'cobertura_preview' not in st.session_state:
    st.session_state.cobertura_preview = None
if 'output_file' not in st.session_state:
    st.session_state.output_file = None
if 'tarifas_df' not in st.session_state:
    st.session_state.tarifas_df = None
if 'colegios_tarifas_df' not in st.session_state:
    st.session_state.colegios_tarifas_df = None


# Sidebar - File Uploads
with st.sidebar:
    st.header("📁 Cargar Archivos")
    
    certificado_file = st.file_uploader(
        "Plantilla Certificado (multi-hoja)",
        type=['xlsx'],
        help="Archivo con una hoja por colegio (ej: 2_CERTIFICACIONES_MES_DE_JULIO.xlsx)"
    )
    
    cobertura_file = st.file_uploader(
        "Plantilla Cobertura (hoja JULIO)",
        type=['xlsx'],
        help="Plantilla base con fórmulas y formato (ej: 3_COBERTURA_EJECUTADA.xlsx)"
    )
    
    st.divider()
    
    st.subheader("⚙️ Configuración Tarifas")
    
    tarifas_file = st.file_uploader(
        "Tabla de Tarifas (tarifas.csv)",
        type=['csv'],
        help="Grupos de tarifa por nivel (A, B, C, D)"
    )
    
    colegios_tarifas_file = st.file_uploader(
        "Mapeo Colegio → Tarifa (colegios_tarifas.csv)",
        type=['csv'],
        help="Opcional: DANE → grupo_tarifa"
    )
    
    # Load default configs if available
    if tarifas_file is None:
        default_tarifas = Path("tarifas.csv")
        if default_tarifas.exists():
            st.session_state.tarifas_df = pd.read_csv(default_tarifas)
            st.info("✅ tarifas.csv cargado por defecto")
    
    if colegios_tarifas_file is None:
        default_colegios = Path("colegios_tarifas.csv")
        if default_colegios.exists():
            st.session_state.colegios_tarifas_df = pd.read_csv(default_colegios, comment='#')
            st.info("✅ colegios_tarifas.csv cargado por defecto")


# Main area - Step by step workflow
tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Cargar y Procesar", 
    "2️⃣ Vista Previa y Editar", 
    "3️⃣ Asignar Tarifas", 
    "4️⃣ Exportar Final"
])


with tab1:
    st.header("Paso 1: Procesar Archivos")
    
    if certificado_file and cobertura_file:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📄 Certificado")
            st.write(f"**Archivo:** {certificado_file.name}")
            st.write(f"**Tamaño:** {certificado_file.size / 1024:.1f} KB")
        
        with col2:
            st.subheader("📋 Cobertura")
            st.write(f"**Archivo:** {cobertura_file.name}")
            st.write(f"**Tamaño:** {cobertura_file.size / 1024:.1f} KB")
        
        if st.button("🚀 Procesar y Generar Borrador", type="primary", use_container_width=True):
            with st.spinner("Leyendo Certificado y escribiendo en Cobertura..."):
                # Save uploaded files to temp
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_cert:
                    tmp_cert.write(certificado_file.getvalue())
                    cert_path = tmp_cert.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_cob:
                    tmp_cob.write(cobertura_file.getvalue())
                    cob_path = tmp_cob.name
                
                try:
                    # Prepare tarifas paths
                    tarifas_path = "tarifas.csv"
                    colegios_tarifas_path = "colegios_tarifas.csv"
                    
                    if tarifas_file:
                        with open(tarifas_path, 'wb') as f:
                            f.write(tarifas_file.getvalue())
                    
                    if colegios_tarifas_file:
                        with open(colegios_tarifas_path, 'wb') as f:
                            f.write(colegios_tarifas_file.getvalue())
                    
                    # Run automation
                    automatizador = PAEAutomatizador(cert_path, cob_path, tarifas_path, colegios_tarifas_path)
                    automatizador.procesar()
                    
                    # Save draft
                    output_path = "cobertura_borrador_streamlit.xlsx"
                    automatizador.writer.save(output_path)
                    
                    st.session_state.automatizador = automatizador
                    st.session_state.colegios_data = {dane: vars(c) for dane, c in automatizador.colegios.items()}
                    st.session_state.output_file = output_path
                    
                    # Get preview
                    preview = automatizador.writer.get_preview_data()
                    st.session_state.cobertura_preview = pd.DataFrame(preview) if preview else pd.DataFrame()
                    
                    st.success(f"✅ Procesados {len(automatizador.colegios)} colegios, {len(automatizador.resultado)} filas escritas")
                    
                    # Show summary
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Colegios en Certificado", len(automatizador.colegios))
                    with col2:
                        st.metric("Filas escritas en Cobertura", len(automatizador.resultado))
                    with col3:
                        sin_cobertura = len(set(automatizador.colegios.keys()) - set([c.codigo_dane for c in automatizador.colegios.values() if hasattr(c, 'codigo_dane')]))
                        st.metric("Sin coincidencia", "Ver log")
                    
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.exception(e)
                finally:
                    # Cleanup temp files
                    try:
                        os.unlink(cert_path)
                        os.unlink(cob_path)
                    except:
                        pass
    else:
        st.info("👈 Carga ambos archivos en el panel lateral para comenzar")


with tab2:
    st.header("Paso 2: Vista Previa y Edición")
    
    if st.session_state.cobertura_preview is not None and not st.session_state.cobertura_preview.empty:
        df = st.session_state.cobertura_preview.copy()
        
        # Filter to show only rows with data
        data_cols = [c for c in df.columns if c.startswith('nivel_') and ('am' in c or 'pm' in c or 'dias' in c)]
        if data_cols:
            df['has_data'] = df[data_cols].notna().any(axis=1)
            df_data = df[df['has_data']].drop(columns=['has_data'])
        else:
            df_data = df
        
        st.subheader(f"📋 Datos a escribir en Cobertura ({len(df_data)} filas con datos)")
        
        # Display editable dataframe
        edited_df = st.data_editor(
            df_data,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "fila": st.column_config.NumberColumn("Fila", disabled=True),
                "dane": st.column_config.TextColumn("DANE", disabled=True),
                "nombre": st.column_config.TextColumn("Colegio", disabled=True, width="large"),
            }
        )
        
        if st.button("💾 Guardar Cambios en Borrador", use_container_width=True):
            # Apply edits to the workbook
            wb = openpyxl.load_workbook(st.session_state.output_file)
            ws = wb.active
            
            # Map columns back to Excel
            nivel_cols = {
                'A': {'am': 3, 'pm': 4, 'dias': 6},
                'B': {'am': 8, 'pm': 9, 'dias': 11},
                'C': {'am': 13, 'pm': 14, 'dias': 16},
                'D': {'am': 18, 'pm': 19, 'dias': 21},
            }
            
            for _, row in edited_df.iterrows():
                fila = int(row['fila'])
                for nivel in ['A', 'B', 'C', 'D']:
                    cols = nivel_cols[nivel]
                    for campo, col_idx in cols.items():
                        val = row.get(f'nivel_{nivel.lower()}_{campo}')
                        if pd.notna(val) and val != '':
                            ws.cell(row=fila, column=col_idx, value=val)
            
            wb.save(st.session_state.output_file)
            st.success("✅ Cambios guardados en el borrador")
            st.rerun()
    else:
        st.info("Primero procesa los archivos en la pestaña 1")


with tab3:
    st.header("Paso 3: Asignar Grupos de Tarifa")
    
    if st.session_state.colegios_data:
        # Load current tarifas
        if st.session_state.tarifas_df is not None:
            grupos_disponibles = st.session_state.tarifas_df['grupo'].unique().tolist()
        else:
            grupos_disponibles = ['grupo_1', 'grupo_2', 'grupo_3', 'grupo_4']
        
        # Load existing mappings
        if st.session_state.colegios_tarifas_df is not None:
            mapping = dict(zip(
                st.session_state.colegios_tarifas_df['codigo_dane'].astype(str),
                st.session_state.colegios_tarifas_df['grupo_tarifa']
            ))
        else:
            mapping = {}
        
        st.subheader("📝 Asignar grupo de tarifa por colegio")
        st.caption("Cada grupo tiene precios distintos para niveles A, B, C, D. Ver tabla abajo.")
        
        # Show tariff table
        if st.session_state.tarifas_df is not None:
            with st.expander("📊 Ver tabla de tarifas"):
                st.dataframe(st.session_state.tarifas_df.pivot(index='nivel', columns='grupo', values='tarifa'), use_container_width=True)
        
        # Create mapping editor
        mapping_data = []
        for dane, colegio in st.session_state.colegios_data.items():
            nombre = colegio.get('nombre', '')
            current_grupo = mapping.get(str(dane), 'grupo_1')
            mapping_data.append({
                'DANE': dane,
                'Colegio': nombre[:60] + '...' if len(nombre) > 60 else nombre,
                'Grupo Tarifa': current_grupo
            })
        
        mapping_df = pd.DataFrame(mapping_data)
        
        edited_mapping = st.data_editor(
            mapping_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "DANE": st.column_config.TextColumn("DANE", disabled=True),
                "Colegio": st.column_config.TextColumn("Colegio", disabled=True, width="large"),
                "Grupo Tarifa": st.column_config.SelectboxColumn(
                    "Grupo Tarifa",
                    options=grupos_disponibles,
                    required=True
                ),
            }
        )
        
        if st.button("💾 Guardar Asignación de Tarifas", use_container_width=True):
            # Save to CSV
            output_df = edited_mapping[['DANE', 'Grupo Tarifa']].rename(columns={'Grupo Tarifa': 'grupo_tarifa'})
            output_df.to_csv("colegios_tarifas.csv", index=False)
            st.session_state.colegios_tarifas_df = output_df
            st.success("✅ Mapeo guardado en colegios_tarifas.csv")
            
            # Regenerate with new tarifas
            if st.session_state.automatizador:
                with st.spinner("Regenerando con nuevas tarifas..."):
                    # Re-process with updated tarifas
                    st.info("Vuelve a la pestaña 1 y presiona 'Procesar' nuevamente para aplicar las nuevas tarifas")
    else:
        st.info("Primero procesa los archivos en la pestaña 1")


with tab4:
    st.header("Paso 4: Exportar Archivo Final")
    
    if st.session_state.output_file and os.path.exists(st.session_state.output_file):
        st.subheader("📥 Descargar Cobertura Final")
        
        with open(st.session_state.output_file, 'rb') as f:
            file_bytes = f.read()
        
        st.download_button(
            label="⬇️ Descargar cobertura_final.xlsx",
            data=file_bytes,
            file_name="cobertura_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
        st.divider()
        
        # Show log
        log_path = st.session_state.output_file.replace('.xlsx', '_log.json')
        if os.path.exists(log_path):
            with open(log_path, 'r', encoding='utf-8') as f:
                log_data = json.load(f)
            
            with st.expander("📋 Ver Log de Procesamiento"):
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Colegios procesados", log_data.get('colegios_procesados', 0))
                    st.metric("Filas escritas", log_data.get('filas_escritas', 0))
                with col2:
                    st.write("**Colegios sin coincidencia en Cobertura:**")
                    for c in log_data.get('colegios_sin_cobertura', [])[:10]:
                        st.write(f"  - {c}")
                    if len(log_data.get('colegios_sin_cobertura', [])) > 10:
                        st.write(f"  ... y {len(log_data.get('colegios_sin_cobertura', [])) - 10} más")
                
                if log_data.get('detalle'):
                    st.write("**Detalle de filas escritas:**")
                    st.dataframe(pd.DataFrame(log_data['detalle']), use_container_width=True)
    else:
        st.info("Completa los pasos anteriores para generar el archivo final")


# Footer
st.divider()
st.caption("PAE Automatización v1.0 | UT Alianza Integral | Desarrollado con Streamlit + openpyxl + pandas")