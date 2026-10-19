FROM python:3.12-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends python3-tk tk fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src ./src
COPY templates ./templates

ENV HOME=/data
WORKDIR /app/src
CMD ["python", "app.py"]
