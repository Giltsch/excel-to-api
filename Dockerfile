FROM python:3.13-slim

WORKDIR /app

# Non-root user for security
RUN groupadd -r appuser && useradd -r -g appuser -d /app appuser

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Remove .env if accidentally copied (secrets must come from env vars)
RUN rm -f .env

RUN chown -R appuser:appuser /app
USER appuser

EXPOSE 8501

CMD ["streamlit", "run", "app/main.py", "--server.port=8501", "--server.address=0.0.0.0", "--server.headless=true", "--server.maxUploadSize=10"]
