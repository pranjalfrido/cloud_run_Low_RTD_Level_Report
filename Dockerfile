FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY generate_low_rtd_report.py .
COPY send_low_rtd_email.py .

CMD ["python", "send_low_rtd_email.py"]
