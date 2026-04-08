# KPI API Server

## Setup

1. Copy `.env.example` to `.env` and fill values.
2. Install dependencies:
   - `npm install`
3. Run migrations:
   - `npm run db:migrate`
4. Seed admin:
   - `npm run db:seed-admin`
5. Start server:
   - `npm run dev`

## Core API

- `POST /api/auth/login`
- `POST /api/auth/refresh`
- `POST /api/auth/logout`
- `GET /api/auth/me`
- `GET /api/users` (admin)
- `POST /api/users` (admin)
- `POST /api/uploads/workbook` (admin, analyst)
- `GET /api/kpi/filters`
- `GET /api/kpi/rows?dataset=raw_data|incident`
- `GET /api/kpi/summary`

