# Build the React frontend
FROM node:24-bookworm-slim AS frontend-build
WORKDIR /app/frontend
COPY frontend/package*.json ./
RUN npm install
COPY frontend/ ./
RUN npm run build

# Run the FastAPI backend and serve the built frontend from the same service
FROM python:3.12-slim AS runtime
WORKDIR /app/backend
ENV PYTHONUNBUFFERED=1
ENV PORT=8000
COPY backend/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY backend/ ./
COPY --from=frontend-build /app/frontend/dist ./app/static
RUN mkdir -p results
EXPOSE 8000
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8000}"]
