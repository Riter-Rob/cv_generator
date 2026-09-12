# CV Generator - container image (works on any Linux host / cloud).
# PDF export uses LibreOffice headless (no Microsoft Word needed).
FROM python:3.10-slim

# System deps: LibreOffice (docx->pdf), Arabic fonts, OpenCV runtime libs
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        fonts-noto fonts-noto-core \
        libgl1 libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# generate the placeholder template from original.docx at build time
RUN python src/build_template.py || true

EXPOSE 8501
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

CMD ["streamlit", "run", "src/app.py", \
     "--server.port=8501", "--server.address=0.0.0.0", \
     "--browser.gatherUsageStats=false"]
