FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

WORKDIR /app

COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# Create a dedicated unprivileged runtime user/group.
RUN groupadd --gid 10001 lteuser \
    && useradd --uid 10001 --gid lteuser --create-home --shell /usr/sbin/nologin lteuser

COPY --chown=lteuser:lteuser . .

USER 10001:10001

EXPOSE 8501

CMD ["streamlit", "run", "streamlit_app.py", "--server.address=0.0.0.0", "--server.port=8501"]
