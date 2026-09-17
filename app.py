"""PAE Automatización - App principal multi-módulo"""
import streamlit as st
import sys
from pathlib import Path

# Add src to Python path for Streamlit Cloud
SRC_PATH = Path(__file__).parent / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

# Import modules
from pae_automatizacion.modules.kelly_primo import render_kelly_module
from pae_automatizacion.modules.laura_jimenez import render_laura_module

st.set_page_config(
    page_title="PAE Automatización",
    page_icon="🍎",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── Estilos PAE Dark Theme ───
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
    --blue: #3B82F6;
    --blue-dim: rgba(59,130,246,0.15);
    --radius-sm: 8px;
    --radius: 12px;
    --radius-lg: 16px;
    --shadow: 0 4px 24px rgba(0,0,0,0.3);
    --shadow-sm: 0 2px 12px rgba(0,0,0,0.2);
}

* {font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;}
.stApp {background: var(--bg); color: var(--text);}
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 1300px;}
header[data-testid="stHeader"] {background: transparent;}
footer, #MainMenu {visibility: hidden;}

.card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    padding: 1.5rem;
    transition: all 0.3s ease;
}
.card:hover {border-color: var(--border-light); box-shadow: var(--shadow); transform: translateY(-2px);}

.stButton>button {
    border-radius: var(--radius) !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    padding: 0.75rem 1.5rem !important;
    transition: all 0.2s ease !important;
    border: none !important;
}
.stButton>button[kind="primary"] {
    background: linear-gradient(135deg, var(--gold), #F2C94C) !important;
    color: #0A0F1A !important;
    box-shadow: 0 4px 20px rgba(231,181,42,0.3) !important;
}
.stButton>button[kind="primary"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 32px rgba(231,181,42,0.4) !important;
}
.stButton>button[kind="secondary"] {
    background: var(--bg-elevated) !important;
    color: var(--text) !important;
    border: 1px solid var(--border) !important;
}
.stButton>button[kind="secondary"]:hover {border-color: var(--gold) !important;}

[data-testid="stFileUploader"] {
    border: 2px dashed var(--border) !important;
    border-radius: var(--radius-lg) !important;
    padding: 2rem !important;
    background: var(--bg-card) !important;
}
[data-testid="stFileUploader"]:hover {border-color: var(--gold) !important;}

.stMetric {
    background: var(--bg-card) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 1rem !important;
}

.sidebar-content {padding: 1rem 0;}
.module-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    padding: 1.5rem;
    margin-bottom: 1rem;
    transition: all 0.3s ease;
    cursor: pointer;
}
.module-card:hover {border-color: var(--gold); transform: translateY(-2px); box-shadow: var(--shadow);}
.module-card.active {border-color: var(--gold); background: var(--gold-dim);}

@media (max-width: 768px) {
    .block-container {padding: 1rem;}
    .card {padding: 1rem;}
}
</style>
""", unsafe_allow_html=True)

# ─── Session State ───
if 'active_module' not in st.session_state:
    st.session_state.active_module = 'kelly_primo'

# ─── Sidebar Navigation ───
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:1rem 0 1.5rem;">
        <div style="width:48px;height:48px;border-radius:12px;background:linear-gradient(135deg,#E7B52A,#F2C94C);display:flex;align-items:center;justify-content:center;margin:0 auto 0.75rem;box-shadow:0 4px 16px rgba(231,181,42,0.3);">
            <span style="font-size:24px;">🍎</span>
        </div>
        <h2 style="margin:0;font-size:1.25rem;font-weight:700;">PAE Automatización</h2>
        <p style="margin:4px 0 0;color:var(--text-muted);font-size:0.8rem;">Programa Alimentación Escolar</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### Módulos")
    
    modules = [
        ("kelly_primo", "👩‍💼 Kelly Primo", "Operación y supervisión\nde entrega PAE", "#E7B52A"),
        ("laura_jimenez", "📊 Laura Jiménez", "Consolidación y reporte\nde cobertura", "#3B82F6"),
    ]
    
    for key, name, desc, color in modules:
        is_active = st.session_state.active_module == key
        card_class = "module-card active" if is_active else "module-card"
        if st.button(name, key=f"nav_{key}", use_container_width=True, 
                     type="primary" if is_active else "secondary"):
            _cleanup_module_state(st.session_state.active_module, key)
            st.session_state.active_module = key
            st.rerun()
        st.caption(desc)
    
    st.divider()
    
    # Global config
    st.markdown("### 📅 Periodo")
    from pae_automatizacion.core.config import MESES_PAE
    mes_options = [m[0] for m in MESES_PAE]
    mes_idx = st.selectbox("Mes", range(len(mes_options)), 
                           format_func=lambda i: mes_options[i], index=6)
    st.session_state.global_mes = mes_options[mes_idx]
    
    anio = st.number_input("Año", min_value=2024, max_value=2030, value=2026)
    st.session_state.global_anio = anio
    
    st.divider()
    st.markdown("""
    <div style="text-align:center;padding:1rem;color:var(--text-dim);font-size:0.7rem;">
        PAE v2.0 · Multi-operador · Multi-mes
    </div>
    """, unsafe_allow_html=True)

# ─── Main Content ───
st.markdown("""
<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:2rem;">
    <div>
        <h1 style="margin:0;font-size:1.75rem;font-weight:700;">PAE Automatización</h1>
        <p style="margin:4px 0 0;color:var(--text-muted);">Certificado → Cobertura · Flujo automatizado</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Cleanup session state when switching modules
def _cleanup_module_state(old_module: str, new_module: str):
    """Limpia el estado del módulo anterior al cambiar."""
    prefixes_to_clear = [
        f'{old_module}_output',
        f'{old_module}_engine',
        f'{old_module}_writer',
        f'{old_module}_df',
        f'{old_module}_cob_path',
    ]
    for key in prefixes_to_clear:
        if key in st.session_state:
            del st.session_state[key]

# Render active module
if st.session_state.active_module == 'kelly_primo':
    render_kelly_module()
elif st.session_state.active_module == 'laura_jimenez':
    render_laura_module()