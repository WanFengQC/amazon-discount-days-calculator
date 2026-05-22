FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY backend.py /app/backend.py
COPY index.html /app/index.html

EXPOSE 8010

CMD ["uvicorn", "backend:app", "--host", "0.0.0.0", "--port", "8010"]

