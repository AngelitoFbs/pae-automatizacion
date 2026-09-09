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

# ─── Estilos con paleta FOMBISOL + Animaciones + UX mejorado ───
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root {
    --navy: #08244A;
    --navy-dark: #102B50;
    --gold: #E7B52A;
    --gold-light: #F2C94C;
    --white: #FFFFFF;
    --green: #2E9B3F;
    --green-light: #4ade80;
    --gray-50: #F8FAFC;
    --gray-100: #F1F5F9;
    --gray-200: #E2E8F0;
    --gray-300: #CBD5E1;
    --gray-400: #94A3B8;
    --gray-600: #475569;
    --gray-700: #334155;
    --red: #EF4444;
    --red-light: #F87171;
}

* {font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;}

.stApp {background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);}

@keyframes fadeInUp {from {opacity:0;transform:translateY(20px);} to {opacity:1;transform:translateY(0);}}
@keyframes slideInRight {from {opacity:0;transform:translateX(30px);} to {opacity:1;transform:translateX(0);}}
@keyframes pulse {0%,100% {box-shadow:0 0 0 0 rgba(231,181,42,0.4);} 50% {box-shadow:0 0 0 12px rgba(231,181,42,0);}}
@keyframes shimmer {0% {background-position:-200% 0;} 100% {background-position:200% 0;}}
@keyframes spin {from {transform:rotate(0deg);} to {transform:rotate(360deg);}}
@keyframes bounce {0%,100% {transform:translateY(0);} 50% {transform:translateY(-4px);}}

.animate-in {animation:fadeInUp 0.5s ease-out forwards;}
.animate-delay-1 {animation-delay:0.05s;}
.animate-delay-2 {animation-delay:0.1s;}
.animate-delay-3 {animation-delay:0.15s;}
.animate-delay-4 {animation-delay:0.2s;}

.step-badge {
    display:inline-flex;align-items:center;gap:8px;
    padding:10px 18px;border-radius:999px;font-weight:600;font-size:0.85rem;
    transition:all 0.3s ease;
    white-space: nowrap;
}
.step-badge::before {
    content: attr(data-step);
    width:22px;height:22px;border-radius:50%;
    display:flex;align-items:center;justify-content:center;
    font-size:0.7rem;font-weight:700;
    background:currentColor;color:var(--white);
}
.step-done {background:linear-gradient(135deg,var(--green) 0%,var(--green-light) 100%);color:#fff;box-shadow:0 4px 14px rgba(46,155,63,0.3);}
.step-done::before {background:var(--white);color:var(--green);}
.step-active {background:linear-gradient(135deg,var(--navy) 0%,var(--navy-dark) 100%);color:#fff;animation:pulse 2s infinite;box-shadow:0 4px 20px rgba(8,36,74,0.3);}
.step-pending {background:var(--gray-100);color:var(--gray-500);}

.step-progress {
    height:4px;background:var(--gray-200);border-radius:2px;margin:0.5rem 0 1rem;
    overflow:hidden;position:relative;
}
.step-progress::after {
    content:'';position:absolute;top:0;left:0;height:100%;
    background:linear-gradient(90deg,var(--gold) 0%,var(--gold-light) 100%);
    border-radius:2px;transition:width 0.5s ease;
}

.card {
    background:var(--white);
    border:1px solid var(--gray-200);
    border-radius:16px;
    padding:1.5rem;
    margin-bottom:1rem;
    box-shadow:0 2px 8px rgba(8,36,74,0.04);
    transition:all 0.3s ease;
}
.card:hover {border-color:var(--gold-light);box-shadow:0 8px 24px rgba(8,36,74,0.08);transform:translateY(-2px);}
.card-gold {border:2px solid var(--gold);background:linear-gradient(135deg,#fffdf5 0%,#ffffff 100%);}
.card-glass {background:rgba(255,255,255,0.9);backdrop-filter:blur(8px);border:1px solid rgba(231,181,42,0.2);}

.metric-big {font-size:2.5rem;font-weight:700;color:var(--navy);background:linear-gradient(135deg,var(--navy) 0%,var(--navy-dark) 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;}

.stButton>button {
    border-radius:10px !important;
    font-weight:600 !important;
    transition:all 0.2s ease !important;
    border:none !important;
    position:relative;overflow:hidden;
}
.stButton>button[kind="primary"] {
    background:linear-gradient(135deg,var(--gold) 0%,var(--gold-light) 100%) !important;
    color:var(--navy) !important;
    box-shadow:0 4px 16px rgba(231,181,42,0.3) !important;
}
.stButton>button[kind="primary"]:hover {
    transform:translateY(-2px);
    box-shadow:0 8px 24px rgba(231,181,42,0.4) !important;
}
.stButton>button[kind="primary"]:active {transform:translateY(0);}
.stButton>button[kind="secondary"] {
    background:var(--white) !important;
    color:var(--navy) !important;
    border:2px solid var(--gray-200) !important;
}
.stButton>button[kind="secondary"]:hover {
    border-color:var(--gold) !important;
    background:linear-gradient(135deg,#fffdf5 0%,#ffffff 100%) !important;
}
.stButton>button:disabled {opacity:0.5;cursor:not-allowed;}

.warning-box {
    background:linear-gradient(135deg,#fff8ed 0%,#fffdf5 100%);
    border-left:4px solid var(--gold);
    padding:1rem 1.25rem;
    border-radius:12px;
    margin:1rem 0;
}
.success-box {
    background:linear-gradient(135deg,#f0fdf4 0%,#dcfce7 100%);
    border-left:4px solid var(--green);
    padding:1rem 1.25rem;
    border-radius:12px;
    margin:1rem 0;
}
.error-box {
    background:linear-gradient(135deg,#fef2f2 0%,#fee2e2 100%);
    border-left:4px solid var(--red);
    padding:1rem 1.25rem;
    border-radius:12px;
    margin:1rem 0;
}
.info-box {
    background:linear-gradient(135deg,#eff6ff 0%,#dbeafe 100%);
    border-left:4px solid var(--navy);
    padding:1rem 1.25rem;
    border-radius:12px;
    margin:1rem 0;
}

[data-testid="stFileUploader"] {
    border:2px dashed var(--gray-300) !important;
    border-radius:16px !important;
    padding:2rem !important;
    background:var(--white) !important;
    transition:all 0.3s ease !important;
}
[data-testid="stFileUploader"]:hover {
    border-color:var(--gold) !important;
    background:linear-gradient(135deg,#fffdf5 0%,#ffffff 100%) !important;
}
[data-testid="stFileUploader"]:focus-within {
    border-color:var(--gold) !important;
    box-shadow:0 0 0 4px rgba(231,181,42,0.15) !important;
}
[data-testid="stFileUploader"] section {padding:0 !important;}
[data-testid="stFileUploader"] small {color:var(--gray-600) !important;}

[data-testid="stDataFrame"] {
    border-radius:12px !important;
    overflow:hidden !important;
    border:1px solid var(--gray-200) !important;
}

.stMetric {
    background:var(--white);
    border:1px solid var(--gray-200);
    border-radius:12px;
    padding:1rem 1.25rem;
    box-shadow:0 2px 8px rgba(8,36,74,0.04);
}
.stMetric:hover {border-color:var(--gold-light);}

.stExpander {
    border:1px solid var(--gray-200) !important;
    border-radius:12px !important;
    background:var(--white) !important;
}
.stExpander summary {font-weight:600;color:var(--navy);}

header[data-testid="stHeader"] {background:transparent;}
footer {visibility:hidden;}

.block-container {padding-top:1.5rem;padding-bottom:2rem;max-width:1200px;}

.shimmer-bg {
    background:linear-gradient(90deg,var(--gray-100) 25%,var(--gray-50) 50%,var(--gray-100) 75%);
    background-size:200% 100%;
    animation:shimmer 1.5s infinite;
}

.tooltip {
    position:relative;cursor:help;
}
.tooltip:hover::after {
    content:attr(data-tip);
    position:absolute;bottom:125%;left:50%;transform:translateX(-50%);
    background:var(--navy);color:var(--white);
    padding:6px 10px;border-radius:6px;font-size:0.75rem;
    white-space:nowrap;z-index:100;
    box-shadow:0 4px 12px rgba(0,0,0,0.15);
    animation:fadeInUp 0.2s ease;
}
.tooltip:hover::before {
    content:'';position:absolute;bottom:115%;left:50%;transform:translateX(-50%);
    border:6px solid transparent;border-top-color:var(--navy);
}

.loading-overlay {
    position:fixed;top:0;left:0;right:0;bottom:0;
    background:rgba(255,255,255,0.95);z-index:9999;
    display:flex;flex-direction:column;align-items:center;justify-content:center;
    gap:1rem;
}
.loading-spinner {
    width:48px;height:48px;border:4px solid var(--gray-200);
    border-top-color:var(--gold);border-radius:50%;
    animation:spin 1s linear infinite;
}
.loading-dots {display:flex;gap:6px;}
.loading-dots span {
    width:10px;height:10px;border-radius:50%;background:var(--gold);
    animation:bounce 1.4s ease-in-out infinite both;
}
.loading-dots span:nth-child(2){animation-delay:0.2s;}
.loading-dots span:nth-child(3){animation-delay:0.4s;}

.empty-state {
    text-align:center;padding:3rem 2rem;color:var(--gray-500);
}
.empty-state svg {width:80px;height:80px;margin-bottom:1rem;opacity:0.5;}
.empty-state h3 {margin:0 0 0.5rem;color:var(--navy);font-size:1.25rem;}
.empty-state p {margin:0;color:var(--gray-600);}

@media (max-width: 768px) {
    .block-container {padding:1rem;}
    .step-badge {padding:8px 12px;font-size:0.75rem;}
    .step-badge::before {width:18px;height:18px;font-size:0.65rem;}
    .card {padding:1rem;border-radius:12px;}
    .stButton>button {padding:0.6rem 1rem;font-size:0.9rem;}
    [data-testid="stFileUploader"] {padding:1.5rem !important;}
    .stMetric {padding:0.75rem 1rem;}
}
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
    return f'<span class="step-badge {cls}" data-step="{n}">{label}</span>'

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

# ─── Progress bar helper ───
progress_pct = {1: 0, 2: 33, 3: 66, 4: 100}.get(st.session_state.step, 0)

# ─── Header con logo y animación ───
st.markdown(f"""
<div class="animate-in" style="text-align:center;padding:1rem 0 0.5rem;">
    <div style="display:inline-flex;align-items:center;gap:12px;margin-bottom:8px;">
        <div style="width:48px;height:48px;border-radius:12px;background:linear-gradient(135deg,var(--gold) 0%,var(--gold-light) 100%);display:flex;align-items:center;justify-content:center;box-shadow:0 4px 16px rgba(231,181,42,0.3);">
            <span style="font-size:24px;color:var(--navy);">🍎</span>
        </div>
        <div style="text-align:left;">
            <h1 style="margin:0;font-size:1.75rem;font-weight:700;color:var(--navy);letter-spacing:-0.5px;">PAE Automatización</h1>
            <p style="margin:2px 0 0 0;font-size:0.9rem;color:var(--gray-600);">Certificado → Cobertura</p>
        </div>
    </div>
    <p style="margin:8px 0 0 0;font-size:0.85rem;color:var(--gray-600);">UT Alianza Integral • Programa de Alimentación Escolar</p>
</div>
""", unsafe_allow_html=True)

# ─── Stepper visual con progress bar ───
st.markdown(f'''
<div class="animate-in animate-delay-1" style="margin-top:0.5rem;">
    <div class="step-progress" style="--progress:{progress_pct}%;">
        <div style="width:{progress_pct}%;"></div>
    </div>
    <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;">
        <span class="step-badge step-done" data-step="1">Subir archivos</span>
        <span class="step-badge step-done" data-step="2">Revisar datos</span>
        <span class="step-badge step-done" data-step="3">Tarifas</span>
        <span class="step-badge step-done" data-step="4">Descargar</span>
    </div>
</div>
''', unsafe_allow_html=True)

# Actualizar clases según step actual
if st.session_state.step == 1:
    stepper_html = '''
    <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;">
        <span class="step-badge step-active" data-step="1">Subir archivos</span>
        <span class="step-badge step-pending" data-step="2">Revisar datos</span>
        <span class="step-badge step-pending" data-step="3">Tarifas</span>
        <span class="step-badge step-pending" data-step="4">Descargar</span>
    </div>
    '''
elif st.session_state.step == 2:
    stepper_html = '''
    <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;">
        <span class="step-badge step-done" data-step="1">Subir archivos</span>
        <span class="step-badge step-active" data-step="2">Revisar datos</span>
        <span class="step-badge step-pending" data-step="3">Tarifas</span>
        <span class="step-badge step-pending" data-step="4">Descargar</span>
    </div>
    '''
elif st.session_state.step == 3:
    stepper_html = '''
    <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;">
        <span class="step-badge step-done" data-step="1">Subir archivos</span>
        <span class="step-badge step-done" data-step="2">Revisar datos</span>
        <span class="step-badge step-active" data-step="3">Tarifas</span>
        <span class="step-badge step-pending" data-step="4">Descargar</span>
    </div>
    '''
else:
    stepper_html = '''
    <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;">
        <span class="step-badge step-done" data-step="1">Subir archivos</span>
        <span class="step-badge step-done" data-step="2">Revisar datos</span>
        <span class="step-badge step-done" data-step="3">Tarifas</span>
        <span class="step-badge step-active" data-step="4">Descargar</span>
    </div>
    '''

st.markdown(f'<div class="animate-in animate-delay-1" style="margin-top:1rem;">{stepper_html}</div>', unsafe_allow_html=True)

st.divider()

# ═══════════════════════════════════════════
# PASO 1: SUBIR ARCHIVOS
# ═══════════════════════════════════════════
if st.session_state.step == 1:
    st.markdown('<div class="animate-in animate-delay-2">', unsafe_allow_html=True)
    st.markdown("### 📁 Paso 1: Subir los dos archivos Excel")
    st.caption("Arrastra o selecciona los archivos • Se procesarán automáticamente")
    st.markdown('</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="card card-gold animate-in animate-delay-2">', unsafe_allow_html=True)
        st.markdown("**📄 Plantilla Certificado** <span class='tooltip' data-tip='Archivo con una hoja por colegio, firmado por el rector'></span>", unsafe_allow_html=True)
        st.caption("Una hoja por colegio • Meses: Julio, Agosto, etc.")
        cert_file = st.file_uploader(
            "Certificado",
            type=['xlsx'],
            label_visibility="collapsed",
            help="Ej: 2_CERTIFICACIONES_MES_DE_JULIO.xlsx"
        )
        if cert_file:
            st.markdown(f'<div class="success-box">✅ {cert_file.name} <span style="color:var(--gray-600);">({cert_file.size/1024:.0f} KB)</span></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="empty-state"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg><h3>Sin archivo</h3><p>Sube el Certificado (.xlsx)</p></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="card card-gold animate-in animate-delay-3">', unsafe_allow_html=True)
        st.markdown("**📋 Plantilla Cobertura** <span class='tooltip' data-tip='Plantilla base con fórmulas y formato predefinido (hoja JULIO)'></span>", unsafe_allow_html=True)
        st.caption("Hoja 'JULIO' con fórmulas y formato listo")
        cob_file = st.file_uploader(
            "Cobertura",
            type=['xlsx'],
            label_visibility="collapsed",
            help="Ej: 3_COBERTURA_EJECUTADA_JULIO_2026.xlsx"
        )
        if cob_file:
            st.markdown(f'<div class="success-box">✅ {cob_file.name} <span style="color:var(--gray-600);">({cob_file.size/1024:.0f} KB)</span></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="empty-state"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg><h3>Sin archivo</h3><p>Sube la Cobertura (.xlsx)</p></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.divider()
    
    # Configuración opcional en expander
    with st.expander("⚙️ Configuración avanzada (opcional)"):
        c1, c2 = st.columns(2)
        with c1:
            tarifas_file = st.file_uploader("tarifas.csv", type=['csv'], help="Grupos de tarifa por nivel A-D")
            if tarifas_file:
                st.session_state.tarifas_df = pd.read_csv(tarifas_file)
                st.markdown('<div class="success-box">✅ Tarifas personalizadas cargadas</div>', unsafe_allow_html=True)
        with c2:
            map_file = st.file_uploader("colegios_tarifas.csv", type=['csv'], help="DANE → grupo_tarifa")
            if map_file:
                st.session_state.colegios_tarifas_df = pd.read_csv(map_file, comment='#')
                st.markdown('<div class="success-box">✅ Mapeo colegio-tarifa cargado</div>', unsafe_allow_html=True)
    
    # Keyboard shortcuts hint
    st.markdown("""
    <div style="text-align:center;padding:1rem;color:var(--gray-500);font-size:0.8rem;">
        <kbd style="background:var(--gray-100);border:1px solid var(--gray-300);border-radius:4px;padding:2px 6px;font-family:monospace;">Enter</kbd> Procesar&nbsp;&nbsp;
        <kbd style="background:var(--gray-100);border:1px solid var(--gray-300);border-radius:4px;padding:2px 6px;font-family:monospace;">←</kbd> / <kbd style="background:var(--gray-100);border:1px solid var(--gray-300);border-radius:4px;padding:2px 6px;font-family:monospace;">→</kbd> Navegar pasos
    </div>
    """, unsafe_allow_html=True)
    
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
    st.markdown('<div class="animate-in animate-delay-2">', unsafe_allow_html=True)
    st.markdown("### 📋 Paso 2: Revisar y corregir datos")
    st.caption("Edita directamente en la tabla • Celdas vacías = 0 • Los totales se recalculan solos en Excel")
    st.markdown('</div>', unsafe_allow_html=True)
    
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
        st.markdown('<div class="card animate-in animate-delay-2">', unsafe_allow_html=True)
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
        st.markdown('</div>', unsafe_allow_html=True)
        
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
                st.markdown('<div class="success-box">✅ Cambios guardados</div>', unsafe_allow_html=True)
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
    st.markdown('<div class="animate-in animate-delay-2">', unsafe_allow_html=True)
    st.markdown("### 💰 Paso 3: Asignar grupo de tarifa por colegio")
    st.caption("Cada grupo tiene precios distintos. Ver tabla de referencia abajo.")
    st.markdown('</div>', unsafe_allow_html=True)
    
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
        
        st.markdown('<div class="card animate-in animate-delay-2">', unsafe_allow_html=True)
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
        st.markdown('</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([1,1])
        with col1:
            if st.button("← Volver", use_container_width=True):
                st.session_state.step = 2; st.rerun()
        with col2:
            if st.button("💾 Guardar y continuar", type="primary", use_container_width=True):
                out = edited_map[['DANE', 'Grupo']].rename(columns={'Grupo': 'grupo_tarifa'})
                out.to_csv("colegios_tarifas.csv", index=False)
                st.session_state.tarifas_mapping = dict(zip(out['DANE'], out['grupo_tarifa']))
                st.session_state.colegios_tarifas_df = out
                st.markdown('<div class="success-box">✅ Tarifas guardadas</div>', unsafe_allow_html=True)
                st.session_state.step = 4
                st.rerun()

# ═══════════════════════════════════════════
# PASO 4: DESCARGAR
# ═══════════════════════════════════════════
elif st.session_state.step == 4:
    st.markdown('<div class="animate-in animate-delay-2">', unsafe_allow_html=True)
    st.markdown("### 📥 Paso 4: Descargar archivo final")
    st.caption("El archivo está listo con todas las fórmulas y formatos conservados")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if st.session_state.output_file and Path(st.session_state.output_file).exists():
        with open(st.session_state.output_file, 'rb') as f:
            data = f.read()
        
        st.markdown('<div class="success-box">✅ Archivo listo para descargar</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="card card-gold animate-in animate-delay-2">', unsafe_allow_html=True)
        st.download_button(
            "⬇️ Descargar COBERTURA_FINAL.xlsx",
            data=data,
            file_name="COBERTURA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.divider()
        
        # Resumen final
        st.markdown('<div class="animate-in animate-delay-3">', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        c1.metric("Colegios procesados", len(st.session_state.colegios_data))
        c2.metric("Filas en Cobertura", len(st.session_state.cobertura_df) if st.session_state.cobertura_df is not None else 0)
        c3.metric("Grupos de tarifa usados", len(set(st.session_state.tarifas_mapping.values())))
        st.markdown('</div>', unsafe_allow_html=True)
        
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

# ─── Keyboard shortcuts ───
st.markdown("""
<script>
document.addEventListener('keydown', function(e) {
    // Enter para procesar en paso 1
    if (e.key === 'Enter' && !e.target.matches('input, textarea, select')) {
        const btn = document.querySelector('button[kind="primary"]:not([disabled])');
        if (btn && btn.textContent.includes('Procesar')) btn.click();
    }
    // Flechas para navegar
    if (e.key === 'ArrowRight' && !e.target.matches('input, textarea, select')) {
        const btn = document.querySelector('button[kind="primary"]:not([disabled])');
        if (btn && (btn.textContent.includes('Continuar') || btn.textContent.includes('Guardar'))) btn.click();
    }
    if (e.key === 'ArrowLeft' && !e.target.matches('input, textarea, select')) {
        const btn = document.querySelector('button[kind="secondary"]');
        if (btn && btn.textContent.includes('Volver')) btn.click();
    }
});
</script>
""", unsafe_allow_html=True)