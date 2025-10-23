# === 1. Base image con Python ===
FROM python:3.10-slim

# === 2. Imposta la directory di lavoro ===
WORKDIR /app

# === 3. Copia i file del progetto ===
COPY . /app

# === 4. Installa le dipendenze ===
# NiceGUI richiede alcune librerie di sistema per il rendering (es. fontconfig)
RUN apt-get update && apt-get install -y \
    libglib2.0-0 libsm6 libxrender1 libxext6 libfontconfig1 \
    && pip install --no-cache-dir \
    nicegui==1.4.19 pandas openpyxl matplotlib requests \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# === 5. Esporta la porta per il web server ===
EXPOSE 8080

# === 6. Comando di avvio ===
# --no-reload per ridurre il consumo di risorse su hosting
CMD ["python", "studio_isa.py"]
