# Overdue Parts Report — Frontend

Next.js UI for the weekly Seawind overdue-parts report. Calls the HF Space API directly from the browser.

## Local development

```bash
cp .env.example .env.local
npm install
npm run dev
```

## Vercel deployment

1. Connect this GitHub repo to Vercel
2. Set **Root Directory** to `frontend`
3. Add environment variable:
   - `NEXT_PUBLIC_API_URL` = `https://truonghung239810-lqvhspace.hf.space`

## HF Space (API)

See [`../hf-space/README.md`](../hf-space/README.md). Set on the Space:

- `ALLOWED_ORIGINS` = `https://your-app.vercel.app,https://*.vercel.app`

(Default `*` allows all origins during initial setup.)
