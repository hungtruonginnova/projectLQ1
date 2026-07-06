---
title: LQVH Overdue Parts Report API
emoji: 📊
colorFrom: blue
colorTo: green
sdk: docker
pinned: false
---

# Overdue Parts Report API

FastAPI backend for the weekly Seawind overdue-parts report.

## Endpoints

- `GET /health` — health check
- `POST /api/v1/report` — multipart upload of 4 Excel exports; returns `.xlsx` report

## Environment

- `ALLOWED_ORIGINS` — comma-separated CORS origins (e.g. `https://your-app.vercel.app,https://*.vercel.app`)
