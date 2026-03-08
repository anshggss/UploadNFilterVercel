# Client (React/Vite)

This folder contains the frontend application. Build output is placed in `dist/`.

## Development

```bash
npm install
npm run dev
```

The app will run on `localhost:5173` by default and will proxy API requests to a backend running on the same origin during development.

## Environment Variables

- `VITE_API_URL` – full URL to the filtering backend. If omitted the browser will use a relative path (`/api/filter`), which is useful when the server is served from the same origin during development.

For example, create a `.env.local` with:

```dotenv
VITE_API_URL=https://my-api.example.com/api/filter
```

Re‑build the site after changing env variables:

```bash
npm run build
```

## Deployment

Deploy the contents of this folder (or the `dist/` directory after building) as a static site. Vercel is already configured by the root `vercel.json` to build the client and optionally run the serverless function under `api/` if you choose.
