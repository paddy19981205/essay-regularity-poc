# 英語正則 POC

## Structure

```text
product-poc/
  frontend/   React + Vite UI
  backend/    Python standard-library API server
```

The backend serves `frontend/dist` in POC mode and owns local run history under `backend/server_data/runs`.

## Commands

```bash
npm run build
npm run serve:backend
npm run dev:frontend
npm run serve:poc
```

- `npm run dev:frontend`: starts Vite on `127.0.0.1:5173` and proxies `/api` to `127.0.0.1:8787`.
- `npm run serve:backend`: starts the Python API/static server on `127.0.0.1:8787`.
- `npm run serve:poc`: builds the frontend, then starts the backend serving the built files.
