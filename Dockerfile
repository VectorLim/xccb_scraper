FROM python:3.13-slim

ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1
WORKDIR /app
COPY requirements.txt requirements-web.txt ./
RUN python -m pip install --no-cache-dir -r requirements-web.txt
COPY xccb_report ./xccb_report

USER 10001:10001
EXPOSE 8080
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "1", "--threads", "4", "xccb_report.web:create_app()"]
