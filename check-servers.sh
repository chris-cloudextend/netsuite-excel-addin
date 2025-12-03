#!/bin/bash
# Quick check if backend and tunnel are running

echo "🔍 Checking server status..."
echo ""

# Check backend
if ps aux | grep -q "[p]ython3 server.py"; then
    echo "✅ Backend server: RUNNING"
    ps aux | grep "[p]ython3 server.py" | awk '{print "   PID: " $2}'
else
    echo "❌ Backend server: NOT RUNNING"
    echo "   Run: cd backend && python3 server.py &"
fi

echo ""

# Check tunnel
if ps aux | grep -q "[c]loudflared tunnel"; then
    echo "✅ Cloudflare tunnel: RUNNING"
    ps aux | grep "[c]loudflared tunnel" | awk '{print "   PID: " $2}'
else
    echo "❌ Cloudflare tunnel: NOT RUNNING"
    echo "   Run: cloudflared tunnel --url http://localhost:5002 &"
fi

echo ""
echo "💡 Tip: Run ./restart-servers.sh to restart both"

