FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all source files
COPY . .

# Start server — PORT is injected by Railway
CMD uvicorn api:app --host 0.0.0.0 --port ${PORT:-8000}
