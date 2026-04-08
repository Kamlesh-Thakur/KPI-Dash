# RBAC + Ingestion backend rollout

## 1) Provision Neon Postgres

1. Create a free Neon project.
2. Copy the connection string into `DATABASE_URL`.
3. Run:
   - `npm --prefix server run db:migrate`
   - `npm --prefix server run db:seed-admin`

## 2) Deploy API to Render free tier

1. Connect this repo in Render.
2. Use `render.yaml` as blueprint.
3. Set env vars:
   - `DATABASE_URL`
   - `JWT_SECRET`
   - `JWT_REFRESH_SECRET`
   - `CORS_ORIGIN` (your GitHub Pages URL)

## 3) Frontend API mode

Set in frontend environment:

- `VITE_USE_API_DATA=true`
- `VITE_API_BASE_URL=https://<render-service-url>/api`

Then deploy frontend.

## 4) Validation checklist

- Login with seeded admin account succeeds.
- Admin can create users with `viewer` / `analyst`.
- Analyst can upload workbook; viewer gets `403`.
- Re-uploading same file returns duplicate manifest.
- Uploading overlapping workbook updates/skip counts in summary.
- Dashboard loads through API mode and renders totals.

