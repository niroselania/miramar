FROM python:3.12-slim

WORKDIR /app

RUN pip install --no-cache-dir openpyxl

COPY server.py /app/server.py

EXPOSE 8765

CMD ["python", "server.py"]
