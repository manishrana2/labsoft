# Labsoft

Minimal Vite + TypeScript starter.

## Login

- Auth is served by local API endpoint `POST /api/login`
- Passwords are stored as `bcrypt` hashes in `server/data/users.json`
- Login returns a JWT token and session is validated via `GET /api/me`
- Disabled users cannot login

## Admin Panel

Admin role gets an `Admin Panel` module with:

- User management: add user, enable/disable user, reset password
- Alerts: due-soon and overdue report alerts
- Audit trail: recent action history (login, CRUD, admin actions)
- Backup/restore: create JSON backup and restore selected backup

## Productivity Enhancements

- Form auto-save drafts for both entry forms (stored in browser localStorage)
- Overdue due-date highlighting in both records tables
- Advanced filters in records pages (date + role/parameter/customer filters)
- PDF export along with CSV export in both records pages
- Staff permission hardening: only admins can delete records
- Activity cards on dashboard: total entries, overdue entries, today's new entries
- Soft delete with undo window (8 seconds) before final delete API commit
- Duplicate checks on records (Issue: `Sr.No.`/`Code No.`, Drawn: `Sr.No.`)

## Data Retention

- Issue Register and Receiving Register main data is stored in `server/data/registers.json`
- Permanent register history is stored in `server/data/register-history.json`
- Every create/update/delete is appended to register history for long-term retention
- Backup/restore now includes register history as well

### Password Policy

New passwords must include:

- Minimum 8 characters
- At least 1 letter
- At least 1 number
- At least 1 symbol

## Scripts

- `npm install` – install dependencies
- `npm run dev` – run development server
- `npm run server` – run auth API server on port `3001`
- `npm run dev:full` – run frontend + auth API together
- `npm run share:demo` – run app and generate public demo URL (client opens link, no install needed)
- `npm run build` – type-check and build for production
- `npm run preview` – preview production build

## Zero-Install Client Demo

If your client is non-technical and should not install anything:

1. On your machine, run `npm run share:demo`
2. Copy the `https://...loca.lt` link shown in terminal
3. Share that link with client
4. Client opens link directly in browser

Login:
- Use a valid user account created in your environment.

Important:
- Keep your terminal running while client tests
- Stop with `Ctrl + C` after demo

## Production Deploy (Recommended)

For a stable client link, deploy on Render (no local tunnel needed).

### One-time setup

1. Push this project to GitHub.
2. Open Render dashboard → **New +** → **Blueprint**.
3. Select your GitHub repo (Render will auto-detect `render.yaml`).
4. Click **Apply**.

### What this creates

- One web service running both frontend + API on the same domain
- Persistent disk mounted to `server/data` so users/registers/audit/history stay saved
- Health check on `/api/health`

### Build/Start used in production

- Build: `npm ci --production=false && npm run build`
- Start: `npm run start`

After deploy completes, share the Render URL with client directly.
