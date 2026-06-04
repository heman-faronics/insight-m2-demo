# Insight M2 Demo — Web (Vercel)

> **NOT PRODUCTION CODE** — standalone web demo of the Milestone 2 SSO + ClassLink Rostering integration.

Live demo: deployed on Vercel from `heman-faronics/insight-m2-demo`

## Architecture

| Layer | Technology |
|---|---|
| Frontend | Static HTML + Bootstrap 5 + MSAL Browser (CDN) |
| Auth | `@azure/msal-browser` — popup flow against `faronicsna.onmicrosoft.com` |
| ClassLink API | Vercel serverless function (`/api/classlink.js`) — proxies calls to avoid CORS |

## Test accounts

| Account | Maps to ClassLink |
|---|---|
| `teacher1@faronicsna.onmicrosoft.com` | Marian Lucas (5033_T5033-0005) |
| `student1@faronicsna.onmicrosoft.com` | Frank Hoffman (5033_S5033-0002) |
| `student2@faronicsna.onmicrosoft.com` | Gwendolyn Price (5033_S5033-0003) |

## ⚠️ Required: Entra redirect URI

Before the sign-in button works, add the deployment URL as a redirect URI in the Entra app registration:

1. Go to [portal.azure.com](https://portal.azure.com)
2. **Azure Active Directory → App registrations** → find `7eb737a9-32f2-4298-b617-87d35051d95d`
3. **Authentication → Add a platform → Single-page application**
4. Add redirect URI: `https://<your-vercel-url>.vercel.app`
5. Also add `http://localhost:3000` for local development
6. Save

The MSAL popup uses `window.location.origin` as the redirect URI — it must match exactly.

## Local development

```bash
npm install -g vercel
vercel dev   # runs on http://localhost:3000
```

## Deploy to Vercel

```bash
vercel --prod
```

Or connect this GitHub repo in the [Vercel dashboard](https://vercel.com/new) and it deploys automatically on every push.
