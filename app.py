import streamlit as st
import pandas as pd
import openpyxl
import json
import tempfile
import os
from pathlib import Path
from typing import Dict, List
import io

# Import backend logic
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
from pae_automatizador import PAEAutomatizador, CertificadoReader, CoberturaWriter, TarifasManager, ColegioData


st.set_page_config(
    page_title="PAE - Certificado a Cobertura",
    page_icon="🍎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─── Estilos simples ───
st.markdown("""
<style>
    .step-badge {display:inline-block;padding:4px 12px;border-radius:20px;font-weight:600;font-size:0.85rem;}
    .step-done {background:#e8f5e9;color:#2e7d32;}
    .step-active {background:#e3f2fd;color:#1565c0;}
    .step-pending {background:#f5f5f5;color:#9e9e9e;}
    .card {border:1px solid #e0e0e0;border-radius:12px;padding:1.2rem;margin-bottom:1rem;background:#fafafa;}
    .metric-big {font-size:2rem;font-weight:700;color:#1565c0;}
    .stButton>button {border-radius:8px;font-weight:600;}
    .warning-box {background:#fff3e0;border-left:4px solid #ff9800;padding:1rem;border-radius:4px;margin:1rem 0;}
    .success-box {background:#e8f5e9;border-left:4px solid #4caf50;padding:1rem;border-radius:4px;margin:1rem 0;}
</style>
""", unsafe_allow_html=True)

# ─── Estado ───
defaults = {
    'step': 1,
    'automatizador': None,
    'cobertura_df': None,
    'output_file': None,
    'colegios_data': {},
    'tarifas_mapping': {},
    'edited': False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─── Helpers ───
NIVEL_COLS = {
    'A': {'AM': 3, 'PM': 4, 'Días': 6},
    'B': {'AM': 8, 'PM': 9, 'Días': 11},
    'C': {'AM': 13, 'PM': 14, 'Días': 16},
    'D': {'AM': 18, 'PM': 19, 'Días': 21},
}

def step_badge(n, label):
    cls = 'step-done' if st.session_state.step > n else ('step-active' if st.session_state.step == n else 'step-pending')
    return f'<span class="step-badge {cls}">Paso {n}</span> {label}'

def load_default_configs():
    if st.session_state.get('tarifas_df') is None and Path("tarifas.csv").exists():
        try:
            st.session_state.tarifas_df = pd.read_csv("tarifas.csv")
        except:
            pass
    if st.session_state.get('colegios_tarifas_df') is None and Path("colegios_tarifas.csv").exists():
        try:
            df = pd.read_csv("colegios_tarifas.csv", comment='#')
            if 'codigo_dane' in df.columns and 'grupo_tarifa' in df.columns:
                st.session_state.colegios_tarifas_df = df
        except:
            pass

load_default_configs()

# ─── Header ───
st.title("🍎 PAE: Certificado → Cobertura")
st.caption("UT Alianza Integral • Programa de Alimentación Escolar")

# ─── Stepper visual ───
cols = st.columns(4)
steps = [
    ("1️⃣ Subir archivos", 1),
    ("2️⃣ Revisar datos", 2),
    ("3️⃣ Tarifas", 3),
    ("4️⃣ Descargar", 4),
]
for col, (label, n) in zip(cols, steps):
    col.markdown(step_badge(n, label), unsafe_allow_html=True)

st.divider()

# ═══════════════════════════════════════════
# PASO 1: SUBIR ARCHIVOS
# ═══════════════════════════════════════════
if st.session_state.step == 1:
    st.markdown("### 📁 Paso 1: Subir los dos archivos Excel")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**📄 Plantilla Certificado**")
        st.caption("Una hoja por colegio • Meses: Julio, Agosto, etc.")
        cert_file = st.file_uploader(
            "Certificado",
            type=['xlsx'],
            label_visibility="collapsed",
            help="Ej: 2_CERTIFICACIONES_MES_DE_JULIO.xlsx"
        )
        if cert_file:
            st.success(f"✅ {cert_file.name} ({cert_file.size/1024:.0f} KB)")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**📋 Plantilla Cobertura**")
        st.caption("Hoja 'JULIO' con fórmulas y formato listo")
        cob_file = st.file_uploader(
            "Cobertura",
            type=['xlsx'],
            label_visibility="collapsed",
            help="Ej: 3_COBERTURA_EJECUTADA_JULIO_2026.xlsx"
        )
        if cob_file:
            st.success(f"✅ {cob_file.name} ({cob_file.size/1024:.0f} KB)")
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.divider()
    
    # Configuración opcional en expander
    with st.expander("⚙️ Configuración avanzada (opcional)"):
        c1, c2 = st.columns(2)
        with c1:
            tarifas_file = st.file_uploader("tarifas.csv", type=['csv'], help="Grupos de tarifa por nivel A-D")
            if tarifas_file:
                st.session_state.tarifas_df = pd.read_csv(tarifas_file)
                st.success("Tarifas personalizadas cargadas")
        with c2:
            map_file = st.file_uploader("colegios_tarifas.csv", type=['csv'], help="DANE → grupo_tarifa")
            if map_file:
                st.session_state.colegios_tarifas_df = pd.read_csv(map_file, comment='#')
                st.success("Mapeo colegio-tarifa cargado")
    
    if cert_file and cob_file:
        if st.button("🚀 Procesar y Continuar", type="primary", use_container_width=True):
            with st.spinner("Leyendo certificados y llenando cobertura..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_cert:
                    tmp_cert.write(cert_file.getvalue()); cert_path = tmp_cert.name
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_cob:
                    tmp_cob.write(cob_file.getvalue()); cob_path = tmp_cob.name
                try:
                    tarifas_path = "tarifas.csv"
                    map_path = "colegios_tarifas.csv"
                    if 'tarifas_df' in st.session_state and st.session_state.tarifas_df is not None:
                        st.session_state.tarifas_df.to_csv(tarifas_path, index=False)
                    if 'colegios_tarifas_df' in st.session_state and st.session_state.colegios_tarifas_df is not None:
                        st.session_state.colegios_tarifas_df.to_csv(map_path, index=False)
                    
                    aut = PAEAutomatizador(cert_path, cob_path, tarifas_path, map_path)
                    aut.procesar()
                    
                    out_path = "cobertura_borrador.xlsx"
                    aut.writer.save(out_path)
                    
                    st.session_state.automatizador = aut
                    st.session_state.colegios_data = {d: vars(c) for d, c in aut.colegios.items()}
                    st.session_state.output_file = out_path
                    
                    preview = aut.writer.get_preview_data()
                    st.session_state.cobertura_df = pd.DataFrame(preview) if preview else pd.DataFrame()
                    
                    # Cargar mapeo actual de tarifas
                    if Path(map_path).exists():
                        dfm = pd.read_csv(map_path, comment='#')
                        st.session_state.tarifas_mapping = dict(zip(dfm['codigo_dane'].astype(str), dfm['grupo_tarifa']))
                    
                    st.session_state.step = 2
                    st.session_state.edited = False
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.exception(e)
                finally:
                    try: os.unlink(cert_path); os.unlink(cob_path)
                    except: pass

# ═══════════════════════════════════════════
# PASO 2: REVISAR Y EDITAR
# ═══════════════════════════════════════════
elif st.session_state.step == 2:
    st.markdown("### 📋 Paso 2: Revisar y corregir datos")
    
    if st.session_state.cobertura_df is None or st.session_state.cobertura_df.empty:
        st.warning("No hay datos. Vuelve al Paso 1.")
        if st.button("← Volver"): st.session_state.step = 1; st.rerun()
    else:
        df = st.session_state.cobertura_df.copy()
        
        # Solo filas con datos
        data_cols = [c for c in df.columns if c.startswith('nivel_') and ('am' in c or 'pm' in c or 'dias' in c)]
        if data_cols:
            df['_tiene_datos'] = df[data_cols].notna().any(axis=1)
            df_show = df[df['_tiene_datos']].drop(columns=['_tiene_datos']).reset_index(drop=True)
        else:
            df_show = df.reset_index(drop=True)
        
        # Renombrar columnas a nombres amigables (solo las que existan)
        rename_map = {
            'fila': 'Fila',
            'dane': 'DANE',
            'nombre': 'Colegio',
        }
        for n in ['A','B','C','D']:
            rename_map[f'nivel_{n.lower()}_am'] = f'{n} - AM'
            rename_map[f'nivel_{n.lower()}_pm'] = f'{n} - PM'
            rename_map[f'nivel_{n.lower()}_dias'] = f'{n} - Días'
        # Solo renombrar columnas que existan
        existing_rename = {k: v for k, v in rename_map.items() if k in df_show.columns}
        df_show = df_show.rename(columns=existing_rename)
        
        # Columnas a mostrar (solo las que existan tras renombrar)
        show_cols = ['Fila', 'DANE', 'Colegio']
        for n in ['A','B','C','D']:
            show_cols += [f'{n} - AM', f'{n} - PM', f'{n} - Días']
        show_cols = [c for c in show_cols if c in df_show.columns]
        
        # Editor
        st.caption("✏️ **Edita directamente en la tabla** • Celdas vacías = 0 • Los totales se recalculan solos en Excel")
        edited = st.data_editor(
            df_show[show_cols],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Fila": st.column_config.NumberColumn("Fila", disabled=True, width="small"),
                "DANE": st.column_config.TextColumn("DANE", disabled=True, width="medium"),
                "Colegio": st.column_config.TextColumn("Colegio", disabled=True, width="large"),
                **{f'{n} - AM': st.column_config.NumberColumn(f'{n} AM', min_value=0, step=1, width="small") for n in ['A','B','C','D'] if f'{n} - AM' in show_cols},
                **{f'{n} - PM': st.column_config.NumberColumn(f'{n} PM', min_value=0, step=1, width="small") for n in ['A','B','C','D'] if f'{n} - PM' in show_cols},
                **{f'{n} - Días': st.column_config.NumberColumn(f'{n} Días', min_value=0, max_value=31, step=1, width="small") for n in ['A','B','C','D'] if f'{n} - Días' in show_cols},
            },
            key="data_editor"
        )
        
        # Detectar cambios
        if not edited.equals(df_show[show_cols]):
            st.session_state.edited = True
            st.session_state.cobertura_df = edited
        
        col1, col2, col3 = st.columns([1,1,1])
        with col1:
            if st.button("← Volver", use_container_width=True):
                st.session_state.step = 1; st.rerun()
        with col2:
            if st.button("💾 Guardar cambios", type="secondary", use_container_width=True, disabled=not st.session_state.edited):
                wb = openpyxl.load_workbook(st.session_state.output_file)
                ws = wb.active
                # Mapear de vuelta a columnas Excel
                for _, row in edited.iterrows():
                    fila = int(row['Fila'])
                    for n in ['A','B','C','D']:
                        cols = NIVEL_COLS[n]
                        for campo_excel, col_idx in cols.items():
                            col_name = f'{n} - {campo_excel}'
                            val = row.get(col_name)
                            if pd.notna(val) and val != '':
                                ws.cell(row=fila, column=col_idx, value=int(val))
                wb.save(st.session_state.output_file)
                st.session_state.edited = False
                st.success("✅ Guardado")
        with col3:
            if st.button("Continuar →", type="primary", use_container_width=True):
                st.session_state.step = 3; st.rerun()
        
        # Resumen rápido
        st.divider()
        c1, c2, c3 = st.columns(3)
        c1.metric("Filas con datos", len(df_show))
        
        def safe_sum(df, col):
            if col in df.columns:
                return int(pd.to_numeric(df[col], errors='coerce').fillna(0).sum())
            return 0
        
        total_am = sum(safe_sum(df_show, f'{n} - AM') for n in ['A','B','C','D'])
        total_pm = sum(safe_sum(df_show, f'{n} - PM') for n in ['A','B','C','D'])
        c2.metric("Total AM", f"{total_am:,}")
        c3.metric("Total PM", f"{total_pm:,}")

# ═══════════════════════════════════════════
# PASO 3: TARIFAS
# ═══════════════════════════════════════════
elif st.session_state.step == 3:
    st.markdown("### 💰 Paso 3: Asignar grupo de tarifa por colegio")
    st.caption("Cada grupo tiene precios distintos. Ver tabla de referencia abajo.")
    
    if not st.session_state.colegios_data:
        st.warning("Sin datos. Vuelve al Paso 1.")
        if st.button("← Volver"): st.session_state.step = 1; st.rerun()
    else:
        # Tabla de referencia
        if st.session_state.get('tarifas_df') is not None:
            with st.expander("📊 Ver tabla de precios por grupo"):
                tdf = st.session_state.tarifas_df.pivot(index='nivel', columns='grupo', values='tarifa')
                tdf.columns.name = None
                st.dataframe(tdf, use_container_width=True)
                st.caption("Niveles E no se usan en Cobertura JULIO (solo A-D)")
        
        # Obtener grupos disponibles
        grupos = ['grupo_1','grupo_2','grupo_3','grupo_4']
        if st.session_state.get('tarifas_df') is not None:
            try:
                grupos = st.session_state.tarifas_df['grupo'].unique().tolist()
            except:
                pass
        
        # Build mapping table
        rows = []
        for dane, col in st.session_state.colegios_data.items():
            nombre = col.get('nombre', '')
            rows.append({
                'DANE': str(dane),
                'Colegio': nombre,
                'Grupo': st.session_state.tarifas_mapping.get(str(dane), 'grupo_1')
            })
        
        map_df = pd.DataFrame(rows)
        
        edited_map = st.data_editor(
            map_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "DANE": st.column_config.TextColumn("DANE", disabled=True, width="small"),
                "Colegio": st.column_config.TextColumn("Colegio", disabled=True, width="large"),
                "Grupo": st.column_config.SelectboxColumn("Grupo tarifa", options=grupos, required=True, width="medium"),
            },
            key="tarifa_editor"
        )
        
        if st.button("💾 Guardar y continuar", type="primary", use_container_width=True):
            out = edited_map[['DANE', 'Grupo']].rename(columns={'Grupo': 'grupo_tarifa'})
            out.to_csv("colegios_tarifas.csv", index=False)
            st.session_state.tarifas_mapping = dict(zip(out['DANE'], out['grupo_tarifa']))
            st.session_state.colegios_tarifas_df = out
            st.success("✅ Tarifas guardadas")
            st.session_state.step = 4
            st.rerun()
        
        if st.button("← Volver", use_container_width=True):
            st.session_state.step = 2; st.rerun()

# ═══════════════════════════════════════════
# PASO 4: DESCARGAR
# ═══════════════════════════════════════════
elif st.session_state.step == 4:
    st.markdown("### 📥 Paso 4: Descargar archivo final")
    
    if st.session_state.output_file and Path(st.session_state.output_file).exists():
        with open(st.session_state.output_file, 'rb') as f:
            data = f.read()
        
        st.markdown('<div class="success-box">✅ Archivo listo para descargar</div>', unsafe_allow_html=True)
        
        st.download_button(
            "⬇️ Descargar COBERTURA_FINAL.xlsx",
            data=data,
            file_name="COBERTURA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
        st.divider()
        
        # Resumen final
        c1, c2, c3 = st.columns(3)
        c1.metric("Colegios procesados", len(st.session_state.colegios_data))
        c2.metric("Filas en Cobertura", len(st.session_state.cobertura_df) if st.session_state.cobertura_df is not None else 0)
        c3.metric("Grupos de tarifa usados", len(set(st.session_state.tarifas_mapping.values())))
        
        # Log
        log_path = Path(st.session_state.output_file).with_name(
            Path(st.session_state.output_file).stem + '_log.json'
        )
        if log_path.exists():
            with st.expander("📋 Ver detalle y advertencias"):
                log = json.loads(log_path.read_text(encoding='utf-8'))
                
                if log.get('colegios_sin_cobertura'):
                    st.markdown("**⚠️ Colegios en Certificado SIN coincidencia en Cobertura:**")
                    for c in log['colegios_sin_cobertura']:
                        st.write(f"  • {c}")
                
                if log.get('colegios_sin_certificado'):
                    st.markdown("**ℹ️ Colegios en Cobertura SIN certificado:**")
                    for c in log['colegios_sin_certificado'][:15]:
                        st.write(f"  • {c}")
                    if len(log['colegios_sin_certificado']) > 15:
                        st.write(f"  ... y {len(log['colegios_sin_certificado']) - 15} más")
                
                if log.get('detalle'):
                    st.markdown("**📄 Filas generadas:**")
                    st.dataframe(pd.DataFrame(log['detalle']), use_container_width=True)
        
        st.divider()
        if st.button("🔄 Procesar otro mes", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()
    else:
        st.warning("No hay archivo generado. Completa los pasos anteriores.")
        if st.button("← Volver al inicio"): 
            for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()

# ─── Footer ───
st.divider()
st.caption("PAE Automatización v1.1 | UT Alianza Integral | Dudas: soporte@utalianzaintegral.gov.co")