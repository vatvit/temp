FROM python:3.12-slim

WORKDIR /app

# Install dependencies for Excel handling
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["python", "src/main.py"]
