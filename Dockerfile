FROM python:3.12-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY transit_report.py .

# Provide configuration at runtime via environment variables (or --env-file).
# Example:
#   docker run --rm --env-file .env -v "$PWD/output:/app/output" awa-transit-report
ENTRYPOINT ["python", "transit_report.py"]
