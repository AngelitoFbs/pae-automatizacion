"""Módulo Kelly Primo - Operación y supervisión de entrega PAE"""
import streamlit as st
import pandas as pd
import traceback
import tempfile
import os
from pathlib import Path

from pae_automatizacion.core.engine import PAEEngine
from pae_automatizacion.core.config import OPERADORES, MESES_PAE, get_template_names
from pae_automatizacion.core.utils import (
    secure_temp_files, validate_excel_file, safe_save_uploaded_file
)

OPERADOR_KEY = "kelly_primo"
OPERADOR = OPERADORES[OPERADOR_KEY]


def render_kelly_module():
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

    # Sidebar config
    with st.sidebar:
        st.markdown("### ⚙️ Configuración")
        
        mes_options = [m[0] for m in MESES_PAE]
        mes_idx = st.selectbox("Mes", range(len(mes_options)), 
                               format_func=lambda i: mes_options[i], index=6)
        mes = mes_options[mes_idx]
        
        anio = st.number_input("Año", min_value=2024, max_value=2030, value=2026)
        
        st.divider()
        st.caption(f"📧 {OPERADOR['email']}")

    # Main content
    cert_name, cob_name = get_template_names(mes)
    
    st.markdown(f"""
    <div style="margin-bottom:1.5rem;">
        <h2 style="margin:0 0 0.25rem;font-size:1.25rem;font-weight:700;">Procesar {mes} {anio}</h2>
        <p style="margin:0;color:var(--text-muted);font-size:0.9rem;">Certificado -> Cobertura ({OPERADOR['nombre']})</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**📄 Certificado**")
        st.caption(f"Esperado: {cert_name}")
        cert_file = st.file_uploader("Certificado", type=['xlsx'], label_visibility="collapsed")
        if cert_file:
            st.success(f"✅ {cert_file.name}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**📋 Cobertura**")
        st.caption(f"Esperado: {cob_name}")
        cob_file = st.file_uploader("Cobertura", type=['xlsx'], label_visibility="collapsed")
        if cob_file:
            st.success(f"✅ {cob_file.name}")
        st.markdown('</div>', unsafe_allow_html=True)

    if cert_file and cob_file:
        if st.button("🚀 Procesar y Generar", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tc:
                        tc.write(cert_file.getvalue()); cert_path = tc.name
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tc:
                        tc.write(cob_file.getvalue()); cob_path = tc.name
                    try:
                        engine = PAEEngine(cert_path, cob_path)
                        engine.procesar()
                        
                        out_name = f"{OPERADOR['codigo']}_{mes}_{anio}_cobertura.xlsx"
                        out_path = str(Path("output") / out_name)
                        engine.save_output(out_path)
                        
                        log_path = str(Path("logs") / f"{OPERADOR['codigo']}_{mes}_{anio}_log.json")
                        engine.save_log(log_path)
                        
                        st.session_state[f'{OPERADOR_KEY}_output'] = out_path
                        st.session_state[f'{OPERADOR_KEY}_engine'] = engine
                        st.success(f"✅ Procesado: {len(engine.colegios)} colegios, {len(engine.resultado)} filas")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
                    finally:
                        try: os.unlink(cert_path); os.unlink(cob_path)
                        except: pass
                except Exception as e:
                    st.error(f"Error inesperado: {e}")

    # Show preview if processed
    if f'{OPERADOR_KEY}_engine' in st.session_state:
        engine = st.session_state[f'{OPERADOR_KEY}_engine']
        
        st.divider()
        st.markdown("### 📋 Vista Previa")
        
        preview = engine.get_preview()
        if not preview.empty:
            # Show only rows with data
            data_cols = [c for c in preview.columns if any(x in c for x in ['_AM', '_PM', '_Días'])]
            if data_cols:
                preview['_has'] = preview[data_cols].notna().any(axis=1)
                show = preview[preview['_has']].drop(columns=['_has'])
            else:
                show = preview
            
            st.dataframe(show, use_container_width=True, hide_index=True)
            
            # Download
            out_path = st.session_state[f'{OPERADOR_KEY}_output']
            with open(out_path, 'rb') as f:
                st.download_button(
                    f"⬇️ Descargar {Path(out_path).name}",
                    data=f.read(),
                    file_name=Path(out_path).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", use_container_width=True
                )
        
        # Log summary
        log = engine.get_log_data()
        c1, c2, c3 = st.columns(3)
        c1.metric("Colegios", log['colegios_procesados'])
        c2.metric("Filas", log['filas_escritas'])
        c3.metric("Sin match", len(log['colegios_sin_cobertura']))
        
        if log['colegios_sin_cobertura']:
            with st.expander(f"⚠️ {len(log['colegios_sin_cobertura'])} colegios sin match en Cobertura"):
                for c in log['colegios_sin_cobertura']:
                    st.write(f"• {c}")


if __name__ == "__main__":
    render_kelly_module()