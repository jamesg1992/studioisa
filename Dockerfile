# === BASE IMAGE ===
FROM python:3.11-slim

# === IMPOSTAZIONI DI BASE ===
WORKDIR /app

# Disattiva buffer per output immediato nei log
ENV PYTHONUNBUFFERED=1

# === COPIA FILE NEL CONTAINER ===
COPY . /app

# === INSTALLAZIONE DIPENDENZE ===
# (usa un requirements.txt già presente nel repo)
RUN pip install --no-cache-dir -r requirements.txt

# === CONFIGURAZIONE STREAMLIT ===
# Imposta la porta 8501 e disabilita la telemetria
ENV PORT=8501
ENV STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

# === AVVIO DELL'APPLICAZIONE ===
EXPOSE 8501
CMD ["streamlit", "run", "studio_isa_web.py", "--server.port=8501", "--server.address=0.0.0.0"]
