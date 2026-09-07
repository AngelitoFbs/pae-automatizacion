# PAE Automatización - Deployment Guide

## Opción 1: Streamlit Community Cloud (GRATIS, MÁS FÁCIL)

1. Sube este repo a GitHub
2. Ve a https://share.streamlit.io/
3. Click "New app" → conecta tu repo GitHub
4. Main file: `app.py`
5. Deploy → obtienes URL pública `https://tu-app.streamlit.app`

## Opción 2: Hugging Face Spaces (GRATIS)

1. Crea cuenta en huggingface.co
2. New Space → SDK: Streamlit
3. Sube archivos o conecta GitHub
4. Auto-detecta `app.py` y deploya

## Opción 3: Docker (Railway, Render, Fly.io, VPS)

```bash
# Build local
docker build -t pae-automatizacion .

# Run local
docker run -p 8501:8501 pae-automatizacion

# Deploy a Railway
railway login
railway init
railway up

# Deploy a Render
# Conecta repo GitHub en render.com → New Web Service → Docker
```

## Opción 4: VPS propio (Ubuntu/Debian)

```bash
# En servidor
git clone <tu-repo>
cd pae_automatizacion
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py --server.port 8501 --server.address 0.0.0.0

# Con systemd para producción
sudo cp pae-automatizacion.service /etc/systemd/system/
sudo systemctl enable pae-automatizacion
sudo systemctl start pae-automatizacion
```

## Archivos necesarios para deploy:
- `app.py` - App principal Streamlit
- `requirements.txt` - Dependencias Python
- `.streamlit/config.toml` - Configuración servidor
- `src/pae_automatizador.py` - Backend logic
- `tarifas.csv` - Tabla de tarifas (configurable)
- `colegios_tarifas.csv` - Mapeo DANE→grupo (editable en UI)

## Variables de entorno (opcional):
```bash
STREAMLIT_SERVER_PORT=8501
STREAMLIT_SERVER_ADDRESS=0.0.0.0
STREAMLIT_SERVER_HEADLESS=true
```