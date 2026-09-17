"""Módulo Laura Jiménez - Consolidación y reporte de cobertura PAE"""
import streamlit as st
import pandas as pd
from pathlib import Path

# Absolute imports for Streamlit Cloud compatibility
from pae_automatizacion.core.engine import PAEEngine
from pae_automatizacion.core.config import OPERADORES, MESES_PAE, get_template_names, NIVEL_COLS

OPERADOR_KEY = "laura_jimenez"
OPERADOR = OPERADORES[OPERADOR_KEY]


def render_laura_module():
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:1rem;">
        <div style="width:40px;height:40px;border-radius:10px;background:{OPERADOR['color']};display:flex;align-items:center;justify-content:center;">
            <span style="font-size:18px;">{OPERADOR['icon']}</span>
        </div>
        <div>
            <h2 style="margin:0;font-size:1.25rem;font-weight:700;">{OPERADOR['nombre']}</h2>
            <p style="margin:0;color:var(--text-muted);font-size:0.85rem;">{OPERADOR['descripcion']}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### ⚙️ Configuración")
        
        mes_options = [m[0] for m in MESES_PAE]
        mes_idx = st.selectbox("Mes", range(len(mes_options)), 
                               format_func=lambda i: mes_options[i], index=6)
        mes = mes_options[mes_idx]
        
        anio = st.number_input("Año", min_value=2024, max_value=2030, value=2026)
        
        st.divider()
        st.caption(f"📧 {OPERADOR['email']}")

    cert_name, cob_name = get_template_names(mes)
    
    st.markdown(f"""
    <div style="margin-bottom:1.5rem;">
        <h2 style="margin:0 0 0.25rem;font-size:1.25rem;font-weight:700;">Consolidar {mes} {anio}</h2>
        <p style="margin:0;color:var(--text-muted);font-size:0.9rem;">Revisar y consolidar cobertura final</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**📄 Certificado (referencia)**")
        st.caption(f"Opcional: {cert_name}")
        cert_file = st.file_uploader("Certificado", type=['xlsx'], label_visibility="collapsed")
        if cert_file:
            st.success(f"✅ {cert_file.name}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**📋 Cobertura a consolidar**")
        st.caption(f"Plantilla: {cob_name}")
        cob_file = st.file_uploader("Cobertura", type=['xlsx'], label_visibility="collapsed")
        if cob_file:
            st.success(f"✅ {cob_file.name}")
        st.markdown('</div>', unsafe_allow_html=True)

    if cob_file:
        if cert_file:
            # Proceso completo con certificado
            if st.button("🔄 Procesar completo (Cert + Cobertura)", type="primary", use_container_width=True):
                process_full(st, cert_file, cob_file, mes, anio)
        else:
            # Solo abrir cobertura existente para revisión
            if st.button("📋 Abrir Cobertura para revisión", type="primary", use_container_width=True):
                process_review_only(st, cob_file, mes, anio)


def process_full(st, cert_file, cob_file, mes, anio):
    import tempfile, os
    with st.spinner("Procesando..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tc:
            tc.write(cert_file.getvalue()); cert_path = tc.name
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tc:
            tc.write(cob_file.getvalue()); cob_path = tc.name
        try:
            engine = PAEEngine(cert_path, cob_path)
            engine.procesar()
            
            out_name = f"{OPERADOR['codigo']}_{mes}_{anio}_cobertura_final.xlsx"
            out_path = str(Path("output") / out_name)
            engine.save_output(out_path)
            
            log_path = str(Path("logs") / f"{OPERADOR['codigo']}_{mes}_{anio}_log.json")
            engine.save_log(log_path)
            
            st.session_state[f'{OPERADOR_KEY}_output'] = out_path
            st.session_state[f'{OPERADOR_KEY}_engine'] = engine
            st.success(f"✅ Consolidado: {len(engine.colegios)} colegios")
            st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")
        finally:
            try: os.unlink(cert_path); os.unlink(cob_path)
            except: pass


def process_review_only(st, cob_file, mes, anio):
    import tempfile, os
    with st.spinner("Abriendo cobertura..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tc:
            tc.write(cob_file.getvalue()); cob_path = tc.name
        try:
            # Solo cargar para previsualizar
            from ..core.engine import CoberturaWriter
            writer = CoberturaWriter(cob_path)
            preview = writer.get_preview_data()
            df = pd.DataFrame(preview) if preview else pd.DataFrame()
            
            st.session_state[f'{OPERADOR_KEY}_writer'] = writer
            st.session_state[f'{OPERADOR_KEY}_df'] = df
            st.session_state[f'{OPERADOR_KEY}_cob_path'] = cob_path
            st.success("✅ Cobertura cargada para revisión")
            st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")
        finally:
            try: os.unlink(cob_path)
            except: pass


def show_review_interface(st, mes, anio):
    """Interfaz de revisión y edición de cobertura"""
    writer = st.session_state[f'{OPERADOR_KEY}_writer']
    df = st.session_state[f'{OPERADOR_KEY}_df']
    
    st.divider()
    st.markdown("### 📋 Revisión y Edición de Cobertura")
    
    if df.empty:
        st.warning("No hay datos en la cobertura")
        return
    
    # Filtrar filas con datos
    data_cols = [c for c in df.columns if any(x in c for x in ['_AM', '_PM', '_Días'])]
    if data_cols:
        df['_has'] = df[data_cols].notna().any(axis=1)
        show_df = df[df['_has']].drop(columns=['_has'])
    else:
        show_df = df
    
    # Renombrar columnas amigables
    rename_map = {'fila': 'Fila', 'dane': 'DANE', 'nombre': 'Colegio'}
    for n in ['A','B','C','D']:
        rename_map[f'{n}_AM'] = f'{n} - AM'
        rename_map[f'{n}_PM'] = f'{n} - PM'
        rename_map[f'{n}_Días'] = f'{n} - Días'
    existing = {k:v for k,v in rename_map.items() if k in show_df.columns}
    show_df = show_df.rename(columns=existing)
    
    show_cols = ['Fila', 'DANE', 'Colegio']
    for n in ['A','B','C','D']:
        show_cols += [f'{n} - AM', f'{n} - PM', f'{n} - Días']
    show_cols = [c for c in show_cols if c in show_df.columns]
    
    st.caption("Edita valores • Celdas vacías = 0 • Totales se recalculan en Excel")
    edited = st.data_editor(
        show_df[show_cols], use_container_width=True, hide_index=True, num_rows="dynamic",
        column_config={
            "Fila": st.column_config.NumberColumn("Fila", disabled=True, width="small"),
            "DANE": st.column_config.TextColumn("DANE", disabled=True, width="medium"),
            "Colegio": st.column_config.TextColumn("Colegio", disabled=True, width="large"),
            **{f'{n} - AM': st.column_config.NumberColumn(f'{n} AM', min_value=0, step=1, width="small") 
               for n in ['A','B','C','D'] if f'{n} - AM' in show_cols},
            **{f'{n} - PM': st.column_config.NumberColumn(f'{n} PM', min_value=0, step=1, width="small") 
               for n in ['A','B','C','D'] if f'{n} - PM' in show_cols},
            **{f'{n} - Días': st.column_config.NumberColumn(f'{n} Días', min_value=0, max_value=31, step=1, width="small") 
               for n in ['A','B','C','D'] if f'{n} - Días' in show_cols},
        }, key="laura_editor"
    )
    
    # Guardar cambios
    if st.button("💾 Guardar cambios en Excel", type="secondary", use_container_width=True):
        wb = writer.wb
        ws = writer.ws
        for _, row in edited.iterrows():
            fila = int(row['Fila'])
            for n in ['A','B','C','D']:
                cols = NIVEL_COLS[n]
                for campo_excel, col_idx in cols.items():
                    col_name = f'{n} - {campo_excel}'
                    val = row.get(col_name)
                    if pd.notna(val) and val != '':
                        ws.cell(row=fila, column=col_idx, value=int(val))
        wb.save(st.session_state[f'{OPERADOR_KEY}_cob_path'])
        st.toast("✅ Cambios guardados", icon="✅")
    
    # Exportar final
    st.divider()
    cob_path = st.session_state[f'{OPERADOR_KEY}_cob_path']
    with open(cob_path, 'rb') as f:
        st.download_button(
            f"⬇️ Descargar {OPERADOR['codigo']}_{mes}_{anio}_cobertura_revisada.xlsx",
            data=f.read(),
            file_name=f"{OPERADOR['codigo']}_{mes}_{anio}_cobertura_revisada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True
        )
    
    # Resumen
    c1, c2, c3 = st.columns(3)
    c1.metric("Filas con datos", len(df))
    c2.metric("Colegios únicos", df['nombre'].nunique() if 'nombre' in df.columns else 0)


if __name__ == "__main__":
    render_laura_module()