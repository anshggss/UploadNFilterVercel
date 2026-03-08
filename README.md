# Updates

## Update 4
### Issues Fixed:
- ✅ Added a new function to normalise the address
- ✅ Fixed flag checking logic

## Update 3
### Issues Fixed:
- ✅ Added a new file "Flagged-Add" which holds the possible flats whose owners might have possibly changed flats

## 📋 Update 2
### Issues Fixed:
- ✅ Fixed payment status issue and ensured proper naming as per customer demand
- ✅ Fixed incorrectly calculated order totals
- ✅ Added README.md

## 📋 Update 1
### Issues Fixed:
- ✅ Fixed formatting issues as per customer demand
- ✅ Added a bat file to make it easier to run the server and website locally
- ✅ Added a new file "New-Num" which holds the data for all the numbers which weren't present in the template file
- ✅ Added Instructions.txt

---

## Deployment & Architecture

This repository contains two independent pieces:

1. **Client** – a React/Vite application under `client/` that can be deployed as a **static site** (for example on Vercel or Netlify).
2. **Server** – a Node/Express backend under `server/` (or alternatively the `api/filter.js` serverless function) that accepts two Excel files, filters them, and returns a result.

The client no longer assumes the server is colocated. It uses an environment variable to determine where to send uploads:

```js
// client/src/App.jsx
const apiUrl = import.meta.env.VITE_API_URL || '/api/filter';
```

Set `VITE_API_URL` to the fully‑qualified URL of your backend when you build/deploy the client (e.g. `https://my-api.example.com/api/filter`).

The server is configured to allow CORS by default and reads an optional `CLIENT_URL` variable to restrict origins. It *does not* serve any static files unless `SERVE_STATIC=true` is set, which keeps the two parts completely decoupled.

### Running Locally

```bash
# start backend
cd server && npm install && npm run dev
# in a separate shell serve the client
cd client && npm install && npm run dev
```

If you want to test the full stack from a single port you can set `SERVE_STATIC=true` and run the server; it will serve the built client from `client/dist`.

### Deploying

- **Client**: push the `client/` folder to Vercel (this repo can be used directly with a `vercel.json` file configured at root). Make sure to define `VITE_API_URL` in Vercel settings.
- **Server**: deploy anywhere you like (Heroku, DigitalOcean, AWS, an independent Vercel project using the `api/` directory, etc.). For a lightweight option you can keep the helper in `api/filter.js` and deploy the entire repo to Vercel; that file implements the same filtering logic as the express backend. Ensure your hosting environment sets `PORT` and optionally `CLIENT_URL`.

Each time a user selects files in the browser, the client `fetch`es the backend URL; no files are ever processed on the client itself.
