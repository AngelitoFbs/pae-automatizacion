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

# ─── Estilos estilo Finloop - Clean Fintech ───
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root {
    --bg: #0A0F1A;
    --bg-elevated: #111827;
    --bg-card: #141C2E;
    --border: #1F2A44;
    --border-light: #2D3A5A;
    --text: #F1F5F9;
    --text-muted: #8899B0;
    --text-dim: #5E7292;
    --gold: #E7B52A;
    --gold-hover: #F2C94C;
    --gold-dim: rgba(231,181,42,0.15);
    --green: #22C55E;
    --green-dim: rgba(34,197,94,0.15);
    --red: #EF4444;
    --red-dim: rgba(239,68,68,0.15);
    --blue: #3B82F6;
    --blue-dim: rgba(59,130,246,0.15);
    --radius-sm: 8px;
    --radius: 12px;
    --radius-lg: 16px;
    --shadow: 0 4px 24px rgba(0,0,0,0.3);
    --shadow-sm: 0 2px 12px rgba(0,0,0,0.2);
    --shadow-hover: 0 8px 32px rgba(0,0,0,0.4);
}

* {font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;}

.stApp {
    background: var(--bg);
    color: var(--text);
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1280px;
}

header[data-testid="stHeader"] {background: transparent;}
footer {visibility: hidden;}
#MainMenu {visibility: hidden;}

/* Scrollbar */
::-webkit-scrollbar {width: 8px;height: 8px;}
::-webkit-scrollbar-track {background: var(--bg);}
::-webkit-scrollbar-thumb {background: var(--border);border-radius: 4px;}
::-webkit-scrollbar-thumb:hover {background: var(--border-light);}

/* Animations */
@keyframes fadeInUp {from {opacity:0;transform:translateY(16px);} to {opacity:1;transform:translateY(0);}}
@keyframes fadeIn {from {opacity:0;} to {opacity:1;}}
@keyframes slideUp {from {opacity:0;transform:translateY(24px);} to {opacity:1;transform:translateY(0);}}
@keyframes pulse {0%,100% {box-shadow:0 0 0 0 var(--gold-dim);} 50% {box-shadow:0 0 0 16px transparent;}}
@keyframes shimmer {0% {background-position:-200% 0;} 100% {background-position:200% 0;}}
@keyframes spin {to {transform:rotate(360deg);}}

.animate-in {animation:slideUp 0.6s ease-out forwards;opacity:0;}
.animate-delay-1 {animation-delay:0.05s;}
.animate-delay-2 {animation-delay:0.1s;}
.animate-delay-3 {animation-delay:0.15s;}
.animate-delay-4 {animation-delay:0.2s;}
.animate-delay-5 {animation-delay:0.25s;}

/* Stepper */
.stepper-container {
    display:flex;align-items:center;gap:12px;margin-bottom:2.5rem;padding:0 4px;
}
.stepper-line {
    flex:1;height:2px;background:var(--border);border-radius:1px;position:relative;overflow:hidden;
}
.stepper-line::after {
    content:'';position:absolute;top:0;left:0;height:100%;
    background:linear-gradient(90deg,var(--gold),var(--gold-hover));
    border-radius:1px;transition:width 0.6s ease;
}
.stepper-step {
    display:flex;flex-direction:column;align-items:center;gap:6px;
    opacity:0.4;transition:all 0.3s ease;
}
.stepper-step.active {opacity:1;}
.stepper-step.done {opacity:1;}
.stepper-circle {
    width:40px;height:40px;border-radius:50%;
    display:flex;align-items:center;justify-content:center;
    font-weight:700;font-size:0.85rem;
    background:var(--bg-card);border:2px solid var(--border);
    transition:all 0.3s ease;
}
.stepper-step.done .stepper-circle {
    background:linear-gradient(135deg,var(--green),#16a34a);
    border-color:var(--green);color:#fff;
}
.stepper-step.active .stepper-circle {
    background:var(--navy);border-color:var(--gold);color:var(--gold);
    animation:pulse 2s infinite;
}
.stepper-step.pending .stepper-circle {color:var(--text-dim);}
.stepper-label {font-size:0.7rem;font-weight:500;color:var(--text-dim);text-align:center;white-space:nowrap;min-width:80px;}
.stepper-step.done .stepper-label {color:var(--green);}
.stepper-step.active .stepper-label {color:var(--gold);font-weight:600;}

/* Cards */
.card {
    background:var(--bg-card);
    border:1px solid var(--border);
    border-radius:var(--radius-lg);
    padding:1.5rem;
    transition:all 0.3s ease;
}
.card:hover {
    border-color:var(--border-light);
    box-shadow:var(--shadow-hover);
    transform:translateY(-2px);
}
.card-gold {border-color:var(--gold);background:linear-gradient(135deg,rgba(231,181,42,0.05),var(--bg-card));}
.card-glass {background:rgba(20,28,46,0.9);backdrop-filter:blur(12px);border:1px solid var(--border-light);}

/* Buttons */
.stButton>button {
    border-radius:var(--radius) !important;
    font-weight:600 !important;
    font-size:0.9rem !important;
    padding:0.75rem 1.5rem !important;
    transition:all 0.2s ease !important;
    border:none !important;
    letter-spacing:0.01em;
}
.stButton>button[kind="primary"] {
    background:linear-gradient(135deg,var(--gold),var(--gold-hover)) !important;
    color:#0A0F1A !important;
    box-shadow:0 4px 20px var(--gold-dim) !important;
}
.stButton>button[kind="primary"]:hover {
    transform:translateY(-2px);
    box-shadow:0 8px 32px var(--gold-dim) !important;
}
.stButton>button[kind="primary"]:active {transform:translateY(0);}
.stButton>button[kind="secondary"] {
    background:var(--bg-elevated) !important;
    color:var(--text) !important;
    border:1px solid var(--border) !important;
}
.stButton>button[kind="secondary"]:hover {
    border-color:var(--gold) !important;
    background:var(--bg-card) !important;
}
.stButton>button:disabled {opacity:0.4;cursor:not-allowed;}

/* File Uploader */
[data-testid="stFileUploader"] {
    border:2px dashed var(--border) !important;
    border-radius:var(--radius-lg) !important;
    padding:2.5rem 2rem !important;
    background:var(--bg-card) !important;
    transition:all 0.3s ease !important;
}
[data-testid="stFileUploader"]:hover {
    border-color:var(--gold) !important;
    background:var(--gold-dim) !important;
}
[data-testid="stFileUploader"]:focus-within {
    border-color:var(--gold) !important;
    box-shadow:0 0 0 4px var(--gold-dim) !important;
}
[data-testid="stFileUploader"] section {padding:0 !important;}
[data-testid="stFileUploader"] small {color:var(--text-dim) !important;}

/* DataFrame */
[data-testid="stDataFrame"] {
    border-radius:var(--radius) !important;
    overflow:hidden !important;
    border:1px solid var(--border) !important;
    background:var(--bg-card) !important;
}

/* Metrics */
.stMetric {
    background:var(--bg-card) !important;
    border:1px solid var(--border) !important;
    border-radius:var(--radius) !important;
    padding:1.25rem !important;
    box-shadow:var(--shadow-sm) !important;
}
.stMetric:hover {border-color:var(--border-light) !important;}
[data-testid="stMetricLabel"] {color:var(--text-muted) !important;font-size:0.75rem !important;font-weight:500 !important;}
[data-testid="stMetricValue"] {color:var(--text) !important;font-weight:700 !important;}

/* Expander */
.stExpander {
    border:1px solid var(--border) !important;
    border-radius:var(--radius) !important;
    background:var(--bg-card) !important;
}
.stExpander summary {font-weight:600;color:var(--text) !important;font-size:0.9rem !important;}
.stExpander details {border:none !important;}

/* Inputs */
.stTextInput>div>div>input,
.stSelectbox>div>div,
.stNumberInput>div>div>input {
    background:var(--bg-elevated) !important;
    border:1px solid var(--border) !important;
    border-radius:var(--radius) !important;
    color:var(--text) !important;
}
.stTextInput>div>div>input:focus,
.stSelectbox>div>div:focus-within,
.stNumberInput>div>div>input:focus {
    border-color:var(--gold) !important;
    box-shadow:0 0 0 3px var(--gold-dim) !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap:4px;background:var(--bg-elevated);padding:4px;
    border-radius:var(--radius);border:1px solid var(--border);
}
.stTabs [data-baseweb="tab"] {
    background:transparent !important;color:var(--text-muted) !important;
    font-weight:500 !important;padding:0.75rem 1.25rem !important;
    border-radius:var(--radius-sm) !important;transition:all 0.2s ease !important;
}
.stTabs [aria-selected="true"] {
    background:var(--bg-card) !important;color:var(--gold) !important;
    box-shadow:var(--shadow-sm) !important;
}

/* Progress */
.stProgress>div>div>div {
    background:linear-gradient(90deg,var(--gold),var(--gold-hover)) !important;
    border-radius:2px !important;
}

/* Selectbox */
[data-testid="stSelectbox"] [data-baseweb="select"] {
    background:var(--bg-elevated) !important;border:1px solid var(--border) !important;
}

/* Empty State */
.empty-state {
    text-align:center;padding:3rem 2rem;color:var(--text-dim);
}
.empty-state svg {width:72px;height:72px;margin-bottom:1rem;opacity:0.4;}
.empty-state h3 {margin:0 0 0.5rem;color:var(--text);font-size:1.1rem;font-weight:600;}
.empty-state p {margin:0;color:var(--text-dim);font-size:0.9rem;}

/* KBD */
kbd {
    background:var(--bg-elevated);border:1px solid var(--border);
    border-radius:6px;padding:2px 8px;font-family:monospace;
    font-size:0.75rem;color:var(--text-muted);
}

/* Tooltip */
.tooltip {position:relative;cursor:help;display:inline-block;}
.tooltip:hover::after {
    content:attr(data-tip);position:absolute;bottom:130%;left:50%;transform:translateX(-50%);
    background:var(--bg-elevated);color:var(--text);padding:8px 12px;
    border-radius:var(--radius-sm);font-size:0.75rem;white-space:nowrap;z-index:100;
    border:1px solid var(--border);box-shadow:var(--shadow);animation:fadeInUp 0.15s ease;
}
.tooltip:hover::before {
    content:'';position:absolute;bottom:120%;left:50%;transform:translateX(-50%);
    border:6px solid transparent;border-top-color:var(--bg-elevated);
}

/* Loading */
.loading-dots {display:flex;gap:6px;justify-content:center;padding:2rem;}
.loading-dots span {
    width:8px;height:8px;border-radius:50%;background:var(--gold);
    animation:bounce 1.4s ease-in-out infinite both;
}
.loading-dots span:nth-child(2){animation-delay:0.15s;}
.loading-dots span:nth-child(3){animation-delay:0.3s;}
@keyframes bounce {0%,100% {transform:translateY(0);opacity:0.5;} 50% {transform:translateY(-8px);opacity:1;}}

/* Responsive */
@media (max-width: 768px) {
    .block-container {padding:1rem;}
    .stepper-label {font-size:0.65rem;min-width:70px;}
    .stepper-circle {width:32px;height:32px;font-size:0.75rem;}
    .card {padding:1.25rem;border-radius:var(--radius);}
    .stButton>button {padding:0.6rem 1rem !important;font-size:0.85rem !important;}
    [data-testid="stFileUploader"] {padding:2rem 1.5rem !important;}
    .stMetric {padding:1rem !important;}
    .block-container {padding-top:1.5rem;}
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

# ─── Header Finloop Style ───
st.markdown("""
<div class="animate-in" style="padding:0.5rem 0 1.5rem;">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:1rem;">
        <div style="display:flex;align-items:center;gap:12px;">
            <div style="width:44px;height:44px;border-radius:10px;background:linear-gradient(135deg,var(--gold),var(--gold-hover));display:flex;align-items:center;justify-content:center;box-shadow:0 4px 16px var(--gold-dim);">
                <span style="font-size:20px;">🍎</span>
            </div>
            <div>
                <h1 style="margin:0;font-size:1.5rem;font-weight:700;color:var(--text);letter-spacing:-0.02em;">PAE Automatización</h1>
                <p style="margin:2px 0 0 0;font-size:0.85rem;color:var(--text-muted);">Certificado → Cobertura</p>
            </div>
        </div>
        <div style="text-align:right;">
            <p style="margin:0;font-size:0.75rem;color:var(--text-dim);">UT Alianza Integral</p>
            <p style="margin:2px 0 0 0;font-size:0.7rem;color:var(--gold);font-weight:600;">Programa Alimentación Escolar</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ─── Stepper Finloop Style ───
step = st.session_state.step
st.markdown(f"""
<div class="animate-in animate-delay-1 stepper-container">
    <div class="stepper-line" style="--progress:{ {1:0, 2:33, 3:66, 4:100}.get(step, 0) }%;"><div style="width:{ {1:0, 2:33, 3:66, 4:100}.get(step, 0) }%;"></div></div>
    <div class="stepper-step {'done' if step > 1 else ('active' if step == 1 else 'pending')}">
        <div class="stepper-circle">{'✓' if step > 1 else '1'}</div>
        <span class="stepper-label">1. Subir archivos</span>
    </div>
    <div class="stepper-step {'done' if step > 2 else ('active' if step == 2 else 'pending')}">
        <div class="stepper-circle">{'✓' if step > 2 else '2'}</div>
        <span class="stepper-label">2. Revisar datos</span>
    </div>
    <div class="stepper-step {'done' if step > 3 else ('active' if step == 3 else 'pending')}">
        <div class="stepper-circle">{'✓' if step > 3 else '3'}</div>
        <span class="stepper-label">3. Tarifas</span>
    </div>
    <div class="stepper-step {'done' if step > 4 else ('active' if step == 4 else 'pending')}">
        <div class="stepper-circle">{'✓' if step > 4 else '4'}</div>
        <span class="stepper-label">4. Descargar</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.divider()

# ═══════════════════════════════════════════
# PASO 1: SUBIR ARCHIVOS
# ═══════════════════════════════════════════
if st.session_state.step == 1:
    st.markdown("""
    <div class="animate-in" style="margin-bottom:1.5rem;">
        <h2 style="margin:0 0 0.25rem;font-size:1.25rem;font-weight:700;color:var(--text);">Subir archivos</h2>
        <p style="margin:0;color:var(--text-muted);font-size:0.9rem;">Arrastra o selecciona los dos archivos Excel requeridos</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**📄 Certificado**")
        st.caption("Hoja por colegio • Firmado por rector")
        cert_file = st.file_uploader("Certificado", type=['xlsx'], label_visibility="collapsed", help="2_CERTIFICACIONES_MES_DE_JULIO.xlsx")
        if cert_file:
            st.markdown(f'<div style="padding:0.75rem;background:var(--green-dim);border:1px solid var(--green);border-radius:var(--radius);margin-top:0.5rem;display:flex;align-items:center;gap:0.75rem;"><span style="font-size:1.25rem;">✓</span><div><strong style="color:var(--green);">{cert_file.name}</strong><br><small style="color:var(--text-muted);">{cert_file.size/1024:.0f} KB</small></div></div>', unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="empty-state" style="padding:2rem 1rem;">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:48px;height:48px;margin-bottom:0.5rem;opacity:0.4;">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
                </svg>
                <h3 style="font-size:0.95rem;">Sin archivo</h3>
                <p style="font-size:0.8rem;">Sube el Certificado (.xlsx)</p>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**📋 Cobertura**")
        st.caption("Hoja JULIO • Fórmulas y formato listos")
        cob_file = st.file_uploader("Cobertura", type=['xlsx'], label_visibility="collapsed", help="3_COBERTURA_EJECUTADA_JULIO_2026.xlsx")
        if cob_file:
            st.markdown(f'<div style="padding:0.75rem;background:var(--green-dim);border:1px solid var(--green);border-radius:var(--radius);margin-top:0.5rem;display:flex;align-items:center;gap:0.75rem;"><span style="font-size:1.25rem;">✓</span><div><strong style="color:var(--green);">{cob_file.name}</strong><br><small style="color:var(--text-muted);">{cob_file.size/1024:.0f} KB</small></div></div>', unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="empty-state" style="padding:2rem 1rem;">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:48px;height:48px;margin-bottom:0.5rem;opacity:0.4;">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                </svg>
                <h3 style="font-size:0.95rem;">Sin archivo</h3>
                <p style="font-size:0.8rem;">Sube la Cobertura (.xlsx)</p>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.divider()
    
    # Configuración opcional
    with st.expander("⚙️ Configuración avanzada"):
        c1, c2 = st.columns(2)
        with c1:
            tarifas_file = st.file_uploader("tarifas.csv", type=['csv'], help="Grupos de tarifa por nivel A-D")
            if tarifas_file:
                st.session_state.tarifas_df = pd.read_csv(tarifas_file)
                st.markdown('<div style="padding:0.5rem;background:var(--green-dim);border:1px solid var(--green);border-radius:var(--radius);color:var(--green);font-size:0.85rem;">✓ Tarifas personalizadas cargadas</div>', unsafe_allow_html=True)
        with c2:
            map_file = st.file_uploader("colegios_tarifas.csv", type=['csv'], help="DANE → grupo_tarifa")
            if map_file:
                st.session_state.colegios_tarifas_df = pd.read_csv(map_file, comment='#')
                st.markdown('<div style="padding:0.5rem;background:var(--green-dim);border:1px solid var(--green);border-radius:var(--radius);color:var(--green);font-size:0.85rem;">✓ Mapeo colegio-tarifa cargado</div>', unsafe_allow_html=True)
    
    # Keyboard hint
    st.markdown("""
    <div style="text-align:center;padding:1rem;color:var(--text-dim);font-size:0.75rem;">
        <kbd style="background:var(--bg-elevated);border:1px solid var(--border);border-radius:6px;padding:2px 8px;font-family:monospace;font-size:0.7rem;">Enter</kbd> Procesar&nbsp;&nbsp;
        <kbd style="background:var(--bg-elevated);border:1px solid var(--border);border-radius:6px;padding:2px 8px;font-family:monospace;font-size:0.7rem;">←</kbd> / <kbd style="background:var(--bg-elevated);border:1px solid var(--border);border-radius:6px;padding:2px 8px;font-family:monospace;font-size:0.7rem;">→</kbd> Navegar
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
    st.markdown("""
    <div class="animate-in" style="margin-bottom:1.5rem;">
        <h2 style="margin:0 0 0.25rem;font-size:1.25rem;font-weight:700;color:var(--text);">Revisar y editar</h2>
        <p style="margin:0;color:var(--text-muted);font-size:0.9rem;">Verifica los datos mapeados • Edita directamente en la tabla</p>
    </div>
    """, unsafe_allow_html=True)
    
    if st.session_state.cobertura_df is None or st.session_state.cobertura_df.empty:
        st.markdown("""
        <div class="card" style="text-align:center;padding:3rem;">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:48px;height:48px;margin:0 auto 1rem;opacity:0.4;">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
            </svg>
            <h3 style="margin:0 0 0.5rem;font-size:1.1rem;">Sin datos para revisar</h3>
            <p style="margin:0;color:var(--text-muted);">Vuelve al paso 1 y sube los archivos</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("← Volver al paso 1", use_container_width=True):
            st.session_state.step = 1; st.rerun()
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
            'fila': 'Fila', 'dane': 'DANE', 'nombre': 'Colegio',
        }
        for n in ['A','B','C','D']:
            rename_map[f'nivel_{n.lower()}_am'] = f'{n} - AM'
            rename_map[f'nivel_{n.lower()}_pm'] = f'{n} - PM'
            rename_map[f'nivel_{n.lower()}_dias'] = f'{n} - Días'
        existing_rename = {k: v for k, v in rename_map.items() if k in df_show.columns}
        df_show = df_show.rename(columns=existing_rename)
        
        # Columnas a mostrar (solo las que existan tras renombrar)
        show_cols = ['Fila', 'DANE', 'Colegio']
        for n in ['A','B','C','D']:
            show_cols += [f'{n} - AM', f'{n} - PM', f'{n} - Días']
        show_cols = [c for c in show_cols if c in df_show.columns]
        
        # Editor
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.caption("Edita directamente • Celdas vacías = 0 • Totales se recalculan en Excel")
        edited = st.data_editor(
            df_show[show_cols],
            use_container_width=True, hide_index=True, num_rows="dynamic",
            column_config={
                "Fila": st.column_config.NumberColumn("Fila", disabled=True, width="small"),
                "DANE": st.column_config.TextColumn("DANE", disabled=True, width="medium"),
                "Colegio": st.column_config.TextColumn("Colegio", disabled=True, width="large"),
                **{f'{n} - AM': st.column_config.NumberColumn(f'{n} AM', min_value=0, step=1, width="small") for n in ['A','B','C','D'] if f'{n} - AM' in show_cols},
                **{f'{n} - PM': st.column_config.NumberColumn(f'{n} PM', min_value=0, step=1, width="small") for n in ['A','B','C','D'] if f'{n} - PM' in show_cols},
                **{f'{n} - Días': st.column_config.NumberColumn(f'{n} Días', min_value=0, max_value=31, step=1, width="small") for n in ['A','B','C','D'] if f'{n} - Días' in show_cols},
            }, key="data_editor"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Detectar cambios
        if not edited.equals(df_show[show_cols]):
            st.session_state.edited = True
            st.session_state.cobertura_df = edited
        
        # Botones
        col1, col2, col3 = st.columns([1,1,1])
        with col1:
            if st.button("← Volver", use_container_width=True):
                st.session_state.step = 1; st.rerun()
        with col2:
            if st.button("💾 Guardar", type="secondary", use_container_width=True, disabled=not st.session_state.edited):
                wb = openpyxl.load_workbook(st.session_state.output_file)
                ws = wb.active
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
                st.toast("✅ Cambios guardados", icon="✅")
        with col3:
            if st.button("Continuar →", type="primary", use_container_width=True):
                st.session_state.step = 3; st.rerun()
        
        # Resumen rápido
        st.divider()
        c1, c2, c3 = st.columns(3)
        c1.metric("Filas", len(df_show))
        def safe_sum(df, col):
            if col in df.columns: return int(pd.to_numeric(df[col], errors='coerce').fillna(0).sum())
            return 0
        total_am = sum(safe_sum(df_show, f'{n} - AM') for n in ['A','B','C','D'])
        total_pm = sum(safe_sum(df_show, f'{n} - PM') for n in ['A','B','C','D'])
        c2.metric("Total AM", f"{total_am:,}")
        c3.metric("Total PM", f"{total_pm:,}")

# ═══════════════════════════════════════════
# PASO 3: TARIFAS
# ═══════════════════════════════════════════
elif st.session_state.step == 3:
    st.markdown("""
    <div class="animate-in" style="margin-bottom:1.5rem;">
        <h2 style="margin:0 0 0.25rem;font-size:1.25rem;font-weight:700;color:var(--text);">Asignar tarifas</h2>
        <p style="margin:0;color:var(--text-muted);font-size:0.9rem;">Define el grupo de tarifa por colegio • Ver tabla de precios abajo</p>
    </div>
    """, unsafe_allow_html=True)
    
    if not st.session_state.colegios_data:
        st.markdown("""
        <div class="card" style="text-align:center;padding:3rem;">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:48px;height:48px;margin:0 auto 1rem;opacity:0.4;">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
            </svg>
            <h3 style="margin:0 0 0.5rem;font-size:1.1rem;">Sin datos</h3>
            <p style="margin:0;color:var(--text-muted);">Vuelve al paso 1</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("← Volver al paso 1", use_container_width=True):
            st.session_state.step = 1; st.rerun()
    else:
        # Tabla de referencia
        if st.session_state.get('tarifas_df') is not None:
            with st.expander("📊 Ver tabla de precios por grupo"):
                tdf = st.session_state.tarifas_df.pivot(index='nivel', columns='grupo', values='tarifa')
                tdf.columns.name = None
                st.dataframe(tdf, use_container_width=True)
                st.caption("Niveles E no se usan en Cobertura JULIO (solo A-D)")
        
        grupos = ['grupo_1','grupo_2','grupo_3','grupo_4']
        if st.session_state.get('tarifas_df') is not None:
            try: grupos = st.session_state.tarifas_df['grupo'].unique().tolist()
            except: pass
        
        rows = []
        for dane, col in st.session_state.colegios_data.items():
            nombre = col.get('nombre', '')
            rows.append({
                'DANE': str(dane), 'Colegio': nombre,
                'Grupo': st.session_state.tarifas_mapping.get(str(dane), 'grupo_1')
            })
        map_df = pd.DataFrame(rows)
        
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.caption("Selecciona el grupo de tarifa para cada colegio")
        edited_map = st.data_editor(
            map_df, use_container_width=True, hide_index=True,
            column_config={
                "DANE": st.column_config.TextColumn("DANE", disabled=True, width="small"),
                "Colegio": st.column_config.TextColumn("Colegio", disabled=True, width="large"),
                "Grupo": st.column_config.SelectboxColumn("Grupo tarifa", options=grupos, required=True, width="medium"),
            }, key="tarifa_editor"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([1,1])
        with col1:
            if st.button("← Volver", use_container_width=True):
                st.session_state.step = 2; st.rerun()
        with col2:
            if st.button("Guardar y continuar →", type="primary", use_container_width=True):
                out = edited_map[['DANE', 'Grupo']].rename(columns={'Grupo': 'grupo_tarifa'})
                out.to_csv("colegios_tarifas.csv", index=False)
                st.session_state.tarifas_mapping = dict(zip(out['DANE'], out['grupo_tarifa']))
                st.session_state.colegios_tarifas_df = out
                st.toast("✅ Tarifas guardadas", icon="✅")
                st.session_state.step = 4; st.rerun()

# ═══════════════════════════════════════════
# PASO 4: DESCARGAR
# ═══════════════════════════════════════════
elif st.session_state.step == 4:
    st.markdown("""
    <div class="animate-in" style="margin-bottom:1.5rem;">
        <h2 style="margin:0 0 0.25rem;font-size:1.25rem;font-weight:700;color:var(--text);">Descargar final</h2>
        <p style="margin:0;color:var(--text-muted);font-size:0.9rem;">Archivo listo con fórmulas y formatos conservados</p>
    </div>
    """, unsafe_allow_html=True)
    
    if not (st.session_state.output_file and Path(st.session_state.output_file).exists()):
        st.markdown("""
        <div class="card" style="text-align:center;padding:3rem;">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:48px;height:48px;margin:0 auto 1rem;opacity:0.4;">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
            </svg>
            <h3 style="margin:0 0 0.5rem;font-size:1.1rem;">Sin archivo generado</h3>
            <p style="margin:0;color:var(--text-muted);">Completa los pasos anteriores</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("← Volver al inicio", use_container_width=True):
            for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()
    else:
        with open(st.session_state.output_file, 'rb') as f:
            data = f.read()
        
        st.markdown("""
        <div class="card card-gold" style="text-align:center;padding:2rem;">
            <div style="font-size:3rem;margin-bottom:0.5rem;">✅</div>
            <h3 style="margin:0 0 0.5rem;color:var(--text);">Archivo listo</h3>
            <p style="margin:0 0 1.5rem;color:var(--text-muted);">COBERTURA_FINAL.xlsx con todas las fórmulas</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.download_button(
            "⬇️ Descargar COBERTURA_FINAL.xlsx",
            data=data, file_name="COBERTURA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True
        )
        
        st.divider()
        
        # Resumen final
        c1, c2, c3 = st.columns(3)
        c1.metric("Colegios", len(st.session_state.colegios_data))
        c2.metric("Filas", len(st.session_state.cobertura_df) if st.session_state.cobertura_df is not None else 0)
        c3.metric("Grupos tarifa", len(set(st.session_state.tarifas_mapping.values())))
        
        # Log
        log_path = Path(st.session_state.output_file).with_name(
            Path(st.session_state.output_file).stem + '_log.json'
        )
        if log_path.exists():
            with st.expander("📋 Detalle y advertencias"):
                log = json.loads(log_path.read_text(encoding='utf-8'))
                
                if log.get('colegios_sin_cobertura'):
                    st.markdown("**⚠️ Sin coincidencia en Cobertura:**")
                    for c in log['colegios_sin_cobertura']:
                        st.write(f"  • {c}")
                
                if log.get('colegios_sin_certificado'):
                    st.markdown("**ℹ️ Sin certificado:**")
                    for c in log['colegios_sin_certificado'][:15]:
                        st.write(f"  • {c}")
                    if len(log['colegios_sin_certificado']) > 15:
                        st.write(f"  ... y {len(log['colegios_sin_certificado']) - 15} más")
                
                if log.get('detalle'):
                    st.markdown("**📄 Filas generadas:**")
                    st.dataframe(pd.DataFrame(log['detalle']), use_container_width=True)
        
        st.divider()
        if st.button("🔄 Procesar otro mes", use_container_width=True):
for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()
    
    # ─── Footer ───
    st.divider()
    st.markdown("""
    <div style="text-align:center;padding:1rem;color:var(--text-dim);font-size:0.75rem;">
        PAE Automatización v1.2 · UT Alianza Integral · Programa Alimentación Escolar
        <br><a href="mailto:soporte@utalianzaintegral.gov.co" style="color:var(--gold);">soporte@utalianzaintegral.gov.co</a>
    </div>
    """, unsafe_allow_html=True)