FROM python:3.13.5-slim

WORKDIR /app

RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip3 install --no-cache-dir -r requirements.txt

COPY api.py build_report.py config.yaml ./
COPY template/ ./template/

EXPOSE 7860

HEALTHCHECK CMD curl --fail http://localhost:7860/health

CMD ["uvicorn", "api:app", "--host", "0.0.0.0", "--port", "7860"]
