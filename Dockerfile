FROM python:3.12-slim

WORKDIR /app

RUN if [ -f /etc/apt/sources.list.d/debian.sources ]; then \
      sed -i 's|http://deb.debian.org/debian|https://mirrors.aliyun.com/debian|g; s|http://security.debian.org/debian-security|https://mirrors.aliyun.com/debian-security|g' /etc/apt/sources.list.d/debian.sources; \
    fi

RUN apt-get update \
    && apt-get install -y --no-install-recommends chromium \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY backend.py /app/backend.py
COPY amazon_deal_extractor.py /app/amazon_deal_extractor.py
COPY index.html /app/index.html
COPY fixed_export_template.xlsx /app/fixed_export_template.xlsx

EXPOSE 8010

CMD ["uvicorn", "backend:app", "--host", "0.0.0.0", "--port", "8010"]

