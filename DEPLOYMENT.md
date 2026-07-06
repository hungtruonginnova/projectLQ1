# Deployment Guide

Monorepo layout for **HF Space API** + **Vercel Frontend**.

## URLs

| Service | URL |
|---------|-----|
| API (HF Space) | https://truonghung239810-lqvhspace.hf.space |
| Health check | https://truonghung239810-lqvhspace.hf.space/health |
| GitHub repo | https://github.com/hungtruonginnova/projectLQ1 |

## 1. Push to GitHub

```bash
git push origin main
```

Commit `81279c1` (or later) includes `frontend/` and `hf-space/`.

## 2. Vercel (frontend)

1. Import repo `hungtruonginnova/projectLQ1` in [Vercel Dashboard](https://vercel.com/dashboard)
2. **Root Directory:** `frontend`
3. **Framework:** Next.js (auto-detected)
4. **Environment variables** (Production + Preview):

| Name | Value |
|------|-------|
| `NEXT_PUBLIC_API_URL` | `https://truonghung239810-lqvhspace.hf.space` |

5. Deploy

## 3. Hugging Face Space (API)

Space: `truonghung239810/LQVHSpace`

**Option A — Link GitHub monorepo (recommended)**

1. Space Settings → Repository → Connect GitHub `hungtruonginnova/projectLQ1`
2. Set **Root directory** = `hf-space`
3. Auto-rebuild on push

**Option B — Push subdirectory manually**

```bash
cd /tmp && git clone https://huggingface.co/spaces/truonghung239810/LQVHSpace
rsync -av --delete --exclude '.git' ./projectLQ1/hf-space/ LQVHSpace/
cd LQVHSpace && git add -A && git commit -m "update" && git push
```

### HF environment variables

Space Settings → Variables:

| Name | Value |
|------|-------|
| `ALLOWED_ORIGINS` | `https://your-app.vercel.app,https://*.vercel.app` |

Default `*` (if unset) allows all origins during setup.

## 4. Verify

```bash
# API health
curl https://truonghung239810-lqvhspace.hf.space/health

# CORS preflight
curl -X OPTIONS https://truonghung239810-lqvhspace.hf.space/api/v1/report \
  -H "Origin: https://your-app.vercel.app" \
  -H "Access-Control-Request-Method: POST" -D -
```

Open the Vercel URL, upload 4 weekly exports, and download the Excel report.
