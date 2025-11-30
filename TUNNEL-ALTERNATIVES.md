# Tunnel Alternatives & Architecture Options

## 🎯 Current Setup (Cloudflare Tunnel)

**Current URL:** `https://parish-regard-howto-intl.trycloudflare.com`  
**Status:** ✅ Working perfectly!

### ⚠️ Known Issue:
Cloudflare free tunnels generate **random URLs that change on every restart**. This requires:
1. Updating `docs/functions.js` and `docs/taskpane.html` with new URL
2. Bumping manifest version
3. Pushing to GitHub Pages
4. Re-uploading manifest to Admin Center

**RECOMMENDATION:** Keep the tunnel running 24/7 to avoid URL changes!

---

## 🔍 Free & Secure Alternatives

### 1. **Localtunnel** (Easiest Alternative)
- **Website:** https://localtunnel.github.io/www/
- **Installation:** `npm install -g localtunnel`
- **Usage:** `lt --port 5002`
- **Pros:**
  - ✅ Free, open-source
  - ✅ No account required
  - ✅ Simple setup
  - ✅ HTTPS by default
- **Cons:**
  - ⚠️ Random URLs (similar to Cloudflare)
  - ⚠️ Can be slower
  - ⚠️ Less reliable than Cloudflare

### 2. **Serveo** (SSH-Based)
- **Website:** https://serveo.net/
- **Usage:** `ssh -R 80:localhost:5002 serveo.net`
- **Pros:**
  - ✅ Free, no installation
  - ✅ No account required
  - ✅ SSH-based security
  - ✅ Can request custom subdomain
- **Cons:**
  - ⚠️ Service has had downtime issues historically
  - ⚠️ Limited to SSH connectivity

### 3. **ngrok** (Most Popular, Has Free Tier)
- **Website:** https://ngrok.com/
- **Installation:** `brew install ngrok/ngrok/ngrok`
- **Usage:** `ngrok http 5002`
- **Pros:**
  - ✅ Very reliable
  - ✅ Great documentation
  - ✅ Fixed URLs with free account
  - ✅ Web dashboard for monitoring
- **Cons:**
  - ⚠️ Requires free account signup
  - ⚠️ Free tier has session limits
  - ⚠️ URLs still change unless you pay

### 4. **PageKite** (Open Source)
- **Website:** https://pagekite.net/
- **Installation:** `pip install pagekite`
- **Usage:** `pagekite.py 5002 yourname.pagekite.me`
- **Pros:**
  - ✅ Open source
  - ✅ Custom subdomain
  - ✅ Can self-host the relay
- **Cons:**
  - ⚠️ Free tier very limited
  - ⚠️ Requires account

### 5. **Telebit** (Self-Hostable)
- **Website:** https://telebit.cloud/
- **Installation:** `curl https://get.telebit.io/ | bash`
- **Pros:**
  - ✅ Can self-host relay server
  - ✅ Custom domains
- **Cons:**
  - ⚠️ More complex setup
  - ⚠️ Less active development

---

## 🏗️ Architectural Alternatives (No Tunnels Needed)

### Option 1: **GitHub Pages + Serverless Backend**
**Pros:**
- ✅ No tunnels required
- ✅ 100% free
- ✅ No URLs changing
- ✅ Scales automatically

**Architecture:**
```
Excel Add-in (GitHub Pages)
    ↓
Serverless Functions (Vercel/Netlify/Cloudflare Workers)
    ↓
NetSuite API
```

**Services to consider:**
- **Vercel Functions:** https://vercel.com/docs/functions (Free tier: 100GB-hours/month)
- **Netlify Functions:** https://www.netlify.com/products/functions/ (Free tier: 125k requests/month)
- **Cloudflare Workers:** https://workers.cloudflare.com/ (Free tier: 100k requests/day)

**Migration Effort:** Medium
- Convert `backend/server.py` → JavaScript/TypeScript serverless functions
- Deploy to Vercel/Netlify
- Update `functions.js` with permanent serverless URL
- One-time manifest update

### Option 2: **Railway.app** (Free Tier for Hobby Projects)
**Website:** https://railway.app/
**Pros:**
- ✅ Deploy Flask app directly
- ✅ Fixed URL (doesn't change)
- ✅ Free tier: $5 credit/month
- ✅ Zero configuration

**Migration Effort:** Low
- Push `backend/` to GitHub
- Connect Railway to GitHub repo
- Railway auto-deploys Python app
- Get permanent URL
- Update `functions.js` once

### Option 3: **Fly.io** (Free Tier)
**Website:** https://fly.io/
**Pros:**
- ✅ Deploy Flask app directly
- ✅ Fixed URL
- ✅ Free tier: 3 VMs with 256MB RAM
- ✅ Good for Python apps

**Migration Effort:** Low
- Install `flyctl`
- Run `fly launch` in backend folder
- Deploy with `fly deploy`
- Get permanent URL

### Option 4: **Oracle Cloud Free Tier** (Most Generous)
**Website:** https://www.oracle.com/cloud/free/
**Pros:**
- ✅ Always free tier (not time-limited)
- ✅ Generous: 2 VMs with 1GB RAM each
- ✅ Fixed public IP
- ✅ Your own server, full control

**Migration Effort:** Medium-High
- Create Oracle Cloud account
- Set up VM
- Install Python, configure firewall
- Deploy Flask app manually

---

## 💡 Recommended Solution

### **For Immediate Use (Now):**
**Stick with Cloudflare Tunnel** - It's working perfectly!
- Just **keep it running 24/7** to avoid URL changes
- Use `tmux` or `screen` to keep tunnel alive:
  ```bash
  tmux new -s tunnel
  cloudflared tunnel --url http://localhost:5002
  # Press Ctrl+B, then D to detach
  ```

### **For Production (Future):**
**Railway.app or Fly.io** - Best balance of effort vs. stability
- Fixed URL (no more manifest updates!)
- Free tier sufficient for your usage
- Deploy once, forget about it
- Migration is just 30 minutes of work

### **For Maximum Control:**
**Serverless (Vercel/Netlify)** - If you want 100% uptime guarantee
- Convert Python → JavaScript (weekend project)
- Never worry about servers again
- Infinite scalability

---

## 🚀 Quick Comparison

| Solution | Cost | URL Stability | Setup Time | Best For |
|----------|------|---------------|------------|----------|
| **Cloudflare Tunnel** | Free | ⚠️ Changes on restart | 5 min | Testing |
| **Localtunnel** | Free | ⚠️ Changes always | 5 min | Quick demos |
| **ngrok (free)** | Free | ⚠️ Changes on restart | 10 min | Development |
| **Railway.app** | Free* | ✅ Fixed | 20 min | **RECOMMENDED** |
| **Fly.io** | Free* | ✅ Fixed | 20 min | **RECOMMENDED** |
| **Vercel Functions** | Free | ✅ Fixed | 2-4 hours | Production |
| **Oracle Cloud** | Free | ✅ Fixed | 1-2 hours | Full control |

*Free tier with usage limits, but more than sufficient for your needs.

---

## 📝 Next Steps

1. **Short term:** Use existing Cloudflare tunnel, keep it running
2. **This week:** Research Railway.app or Fly.io
3. **Next weekend:** Migrate to Railway/Fly for permanent URL
4. **Result:** Never update manifest again! 🎉

---

## 🔧 Current Tunnel Management

**To check tunnel status:**
```bash
ps aux | grep cloudflared
```

**To restart tunnel (if needed):**
```bash
pkill -f cloudflared
cloudflared tunnel --url http://localhost:5002 > /tmp/cloudflare-new.log 2>&1 &
```

**To keep tunnel running permanently:**
```bash
# Install tmux if not already installed
brew install tmux

# Start persistent session
tmux new -s netsuite-tunnel

# Inside tmux, start servers
cd "/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised/backend"
python3 server.py &
cloudflared tunnel --url http://localhost:5002

# Detach: Ctrl+B, then D
# Reattach later: tmux attach -t netsuite-tunnel
```

---

**Last Updated:** Nov 30, 2025  
**Current Working Tunnel:** `https://parish-regard-howto-intl.trycloudflare.com`

