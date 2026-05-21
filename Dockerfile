FROM node:20-slim AS frontend
WORKDIR /app/product-poc/frontend
COPY product-poc/frontend/package*.json ./
RUN npm ci
COPY product-poc/frontend/src ./src
COPY product-poc/frontend/index.html ./
COPY product-poc/frontend/vite.config.js ./
RUN npm run build

FROM python:3.11-slim
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    POC_HOST=0.0.0.0 \
    PORT=8787 \
    POC_RUNS_DIR=/data/runs \
    MANUAL_ANALYSIS_MODE=auto \
    GEMINI_MODEL=gemini-2.5-flash-lite

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY tools ./tools
COPY product-poc/backend/server.py ./product-poc/backend/server.py
COPY --from=frontend /app/product-poc/frontend/dist ./product-poc/frontend/dist

RUN mkdir -p /data/runs
VOLUME ["/data"]
EXPOSE 8787

CMD ["python", "product-poc/backend/server.py"]
