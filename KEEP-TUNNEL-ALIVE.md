# Keep Tunnel Alive 24/7

## 🎯 The Solution: tmux

Your tunnel is now running in a **persistent tmux session** that will keep running even if you close the terminal!

---

## 📊 Current Status

Check anytime with:
```bash
tmux ls
```

Should show:
```
netsuite: 1 windows (created ...)
```

---

## 🔍 View Your Servers

Attach to the session to see both servers running:
```bash
tmux attach -t netsuite
```

You'll see:
- **Left pane:** Backend server (Flask)
- **Right pane:** Cloudflare tunnel

### To Detach (Leave Running):
Press: **Ctrl+B** then **D**

---

## 🔄 Common Commands

### Check if running:
```bash
tmux ls
```

### View servers:
```bash
tmux attach -t netsuite
```

### Stop servers:
```bash
tmux kill-session -t netsuite
```

### Restart servers:
```bash
# Kill old session if exists
tmux kill-session -t netsuite 2>/dev/null

# Create new session
tmux new-session -d -s netsuite -n "servers"

# Start backend
tmux send-keys -t netsuite:0 "cd '/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised/backend'" C-m
tmux send-keys -t netsuite:0 "python3 server.py" C-m

# Split and start tunnel
tmux split-window -h -t netsuite:0
tmux send-keys -t netsuite:0.1 "cd '/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised/backend'" C-m
tmux send-keys -t netsuite:0.1 "cloudflared tunnel --url http://localhost:5002 2>&1 | tee /tmp/cloudflare-new.log" C-m
```

### Get current tunnel URL:
```bash
cat /tmp/cloudflare-new.log | grep -oE "https://[a-z0-9-]+\.trycloudflare\.com" | head -1
```

---

## ✅ What Survives

The tunnel will keep running through:
- ✅ Closing terminal
- ✅ Closing Cursor
- ✅ Closing your laptop (sleep mode)
- ✅ Days/weeks of inactivity

---

## ⚠️ What Stops It

The tunnel will stop if:
- ❌ You restart your Mac
- ❌ You manually kill the tmux session
- ❌ Your Mac crashes/loses power
- ❌ Cloudflare tunnel service has issues (rare)

---

## 🔄 After Mac Restart

If you restart your Mac, just run:
```bash
tmux attach -t netsuite
```

If that fails (session doesn't exist), restart it:
```bash
cd "/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised"
./restart-servers.sh
```

---

## 🤖 Auto-Start on Mac Startup (Optional)

If you want the tunnel to start automatically when your Mac boots, we can create a LaunchAgent.

Let me know if you want this!

---

## 💡 Best Practices

1. **Check status daily:**
   ```bash
   tmux ls
   ```

2. **View logs occasionally:**
   ```bash
   tmux attach -t netsuite
   # Check for errors
   # Ctrl+B then D to exit
   ```

3. **If tunnel URL changes:**
   - Get new URL: `cat /tmp/cloudflare-new.log | grep trycloudflare`
   - Update Cloudflare Worker (see PROXY-SOLUTION.md)
   - Takes 30 seconds!

---

## 🎯 Current Setup

**Proxy URL (permanent):**
https://netsuite-proxy.chris-corcoran.workers.dev

**Tunnel URL (may change on restart):**
Check with: `cat /tmp/cloudflare-new.log | grep trycloudflare`

**Excel add-in uses:** Proxy URL (never needs updating!)

---

## 📝 Quick Reference Card

```bash
# View servers
tmux attach -t netsuite

# Leave running (while attached)
Ctrl+B then D

# Check status
tmux ls

# Stop servers
tmux kill-session -t netsuite

# Get tunnel URL
cat /tmp/cloudflare-new.log | grep trycloudflare | tail -1
```

---

**Last Updated:** Nov 30, 2025  
**Status:** Tunnel running in tmux session 'netsuite'

