# Gerador de Prestação de Contas — versão online
FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 DADOS_DIR=/dados PORT=8080

# LibreOffice (planilha em PDF) + fontes com as mesmas medidas de Arial/Calibri (a planilha não "estoura")
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc-nogui fonts-liberation fonts-crosextra-carlito fonts-dejavu-core \
        libglib2.0-0 libgomp1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && pip install --no-cache-dir --no-deps rapidocr_onnxruntime==1.4.4

COPY app ./app
COPY modelo ./modelo
COPY VERSAO.txt .

# roda sem ser root; /dados é onde ficam os usuários, notas e relatórios (montar um volume aqui)
RUN useradd --create-home gerador && mkdir -p /dados && chown gerador:gerador /dados
USER gerador
VOLUME ["/dados"]
EXPOSE 8080
HEALTHCHECK --interval=60s --timeout=5s CMD python -c "import urllib.request,os;urllib.request.urlopen('http://127.0.0.1:%s/api/ping'%os.environ.get('PORT','8080'),timeout=4)"
CMD ["python", "app/servidor.py"]
