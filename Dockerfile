# ─────────────────────────────────────────────────────────────────────────────
# AI 벤치마킹 보고서 시스템  ·  Dockerfile
# Base: python:3.11-slim  |  Port: 8501
# ─────────────────────────────────────────────────────────────────────────────
FROM python:3.11-slim

# ── 시스템 의존성 ─────────────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential \
        curl \
        libxml2-dev \
        libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# ── Python 패키지 설치 (레이어 캐시 최적화) ──────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
 && pip install --no-cache-dir -r requirements.txt

# ── 애플리케이션 파일 복사 ────────────────────────────────────────────────────
COPY app.py .
COPY .streamlit/ .streamlit/

# ── 환경 변수 ─────────────────────────────────────────────────────────────────
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    STREAMLIT_SERVER_PORT=8501 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

# ── 보안: 비루트 사용자 ───────────────────────────────────────────────────────
RUN useradd -m -u 1000 appuser && chown -R appuser:appuser /app
USER appuser

# ── 헬스체크 ──────────────────────────────────────────────────────────────────
HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=3 \
    CMD curl -f http://localhost:8501/_stcore/health || exit 1

EXPOSE 8501

ENTRYPOINT ["streamlit", "run", "app.py", \
            "--server.port=8501", \
            "--server.address=0.0.0.0", \
            "--server.headless=true"]
