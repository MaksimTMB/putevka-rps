FROM python:3.11-slim

WORKDIR /opt/asu_waybills
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && pip install --no-cache-dir -r requirements.txt

COPY . .
RUN cp -n data/app_data.json data/putevka_db.json || true

EXPOSE 8080
CMD ["python", "server.py"]
