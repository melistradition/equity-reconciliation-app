# Deploy the reconciliation app publicly on Render

This version is deployment-ready as one public web service. The React frontend is built into the FastAPI backend, so users open one public URL and file uploads call the API on the same domain.

## One-time setup

1. Create a free GitHub account if needed.
2. Create a new GitHub repository, for example `equity-reconciliation-app`.
3. Upload all files from this folder to that repository.

## Deploy on Render

1. Go to Render and sign in.
2. Click **New +**.
3. Choose **Web Service**.
4. Connect your GitHub repository.
5. For environment/runtime, choose **Docker** if Render asks.
6. Render will use the included `Dockerfile`.
7. Click **Create Web Service**.
8. When deploy finishes, open the Render URL, for example:
   `https://equity-reconciliation-app.onrender.com`

## Why this fixes Failed to fetch

Before, the frontend ran on `localhost:5173` and tried to call the backend on `127.0.0.1:8000`. On a public site, a visitor's browser cannot access your computer's local backend.

This deployment serves both frontend and backend from the same public URL, so API calls go to the same domain as the website.

## Local development still works

Backend:

```bash
cd backend
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m uvicorn app.main:app --reload --port 8000
```

Frontend:

```bash
cd frontend
npm install
npm run dev
```

Then open the Vite URL, usually `http://localhost:5173/`.
