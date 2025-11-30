# Stable Proxy URL Solution - Never Update Manifest Again!

## 🎯 The Problem

Cloudflare free tunnels get a new random URL every restart. Every time = manifest update = pain 😫

## 💡 The Solution: Permanent Proxy Layer

```
Excel → PERMANENT PROXY URL → Current Tunnel → Backend
        (never changes)        (changes)
```

---

## 🔧 Best Solution: Cloudflare Workers (10 min setup)

**Permanent URL:** `https://netsuite-proxy.YOUR-NAME.workers.dev`

### Worker Code (Entire Implementation):

```javascript
export default {
  async fetch(request) {
    // ⚠️ UPDATE THIS LINE WHEN TUNNEL RESTARTS (30 seconds)
    const TUNNEL_URL = 'https://survivors-specialist-elegant-monthly.trycloudflare.com';
    
    // Handle CORS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type',
        }
      });
    }
    
    // Forward request to current tunnel
    const url = new URL(request.url);
    const targetUrl = TUNNEL_URL + url.pathname + url.search;
    
    const response = await fetch(targetUrl, {
      method: request.method,
      headers: request.headers,
      body: request.body
    });
    
    // Add CORS headers
    const newResponse = new Response(response.body, response);
    newResponse.headers.set('Access-Control-Allow-Origin', '*');
    newResponse.headers.set('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    newResponse.headers.set('Access-Control-Allow-Headers', 'Content-Type');
    
    return newResponse;
  }
};
```

### Setup (10 minutes, ONE TIME):

```bash
# 1. Install Cloudflare CLI
npm install -g wrangler

# 2. Login
wrangler login

# 3. Create worker
mkdir ~/netsuite-proxy
cd ~/netsuite-proxy
wrangler init netsuite-proxy

# 4. Create index.js with code above

# 5. Deploy
wrangler deploy
```

You'll get: `https://netsuite-proxy.YOUR-SUBDOMAIN.workers.dev`

### Update Manifest (ONE TIME ONLY):

Change `docs/functions.js` and `docs/taskpane.html`:
```javascript
const SERVER_URL = 'https://netsuite-proxy.YOUR-SUBDOMAIN.workers.dev';
```

Upload manifest v1.0.0.29 → **NEVER UPDATE AGAIN!** 🎉

### When Tunnel Restarts (30 seconds):

1. Open Cloudflare Workers dashboard
2. Edit worker
3. Change `TUNNEL_URL` line
4. Click "Deploy"
5. Done! Excel keeps working!

---

## 🎯 Alternative: Vercel Edge Function

Similar concept, slightly different setup:

```javascript
// api/proxy.js
export const config = { runtime: 'edge' };

export default async function handler(req) {
  const TUNNEL_URL = 'https://survivors-specialist-elegant-monthly.trycloudflare.com';
  
  const url = new URL(req.url);
  const targetUrl = TUNNEL_URL + url.pathname + url.search;
  
  const response = await fetch(targetUrl, {
    method: req.method,
    headers: req.headers,
    body: req.body
  });
  
  return new Response(response.body, {
    status: response.status,
    headers: {
      ...Object.fromEntries(response.headers),
      'Access-Control-Allow-Origin': '*'
    }
  });
}
```

Permanent URL: `https://netsuite-proxy.vercel.app`

---

## 📊 Comparison

| Solution | Setup | Update When Tunnel Restarts | Free Tier |
|----------|-------|----------------------------|-----------|
| **Current (No Proxy)** | 0 min | ⚠️ 10 min (manifest + cache clear) | Free |
| **Cloudflare Workers** | 10 min | ✅ 30 sec (edit worker) | 100k req/day |
| **Vercel Edge** | 15 min | ✅ 1 min (redeploy) | 100GB/mo |

---

## 🚀 Quick Start Guide

1. Go to https://workers.cloudflare.com/
2. Sign up (free)
3. Create new worker
4. Paste code above
5. Deploy
6. Get permanent URL
7. Update manifest ONE TIME
8. **Never update manifest again!**

When tunnel restarts: Edit worker's `TUNNEL_URL`, deploy (30 sec)

---

## 💰 Cost

**$0 forever!** Cloudflare Workers free tier: 100,000 requests/day

Your usage: ~1,000 requests/day → well within limits

---

## 🎯 Recommendation

**Do this!** It solves your exact problem:
- ✅ Permanent URL in manifest
- ✅ 30-second updates vs 10-minute manifest dance
- ✅ No Excel cache clearing
- ✅ 100% free
- ✅ Can automate updates with script

**Last Updated:** Nov 30, 2025
