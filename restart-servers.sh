#!/bin/bash
# Quick script to restart servers in tmux

echo "================================================"
echo "🔄 Restarting NetSuite Servers"
echo "================================================"
echo ""

# Kill existing session
echo "Stopping old session..."
tmux kill-session -t netsuite 2>/dev/null
sleep 2

# Create new session
echo "Creating new session..."
tmux new-session -d -s netsuite -n "servers"

# Start backend
tmux send-keys -t netsuite:0 "cd '/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised/backend'" C-m
tmux send-keys -t netsuite:0 "python3 server.py" C-m

# Split and start tunnel
sleep 2
tmux split-window -h -t netsuite:0
tmux send-keys -t netsuite:0.1 "cd '/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised/backend'" C-m
tmux send-keys -t netsuite:0.1 "cloudflared tunnel --url http://localhost:5002 2>&1 | tee /tmp/cloudflare-new.log" C-m

echo ""
echo "✅ Servers started in tmux session 'netsuite'"
echo ""
echo "View with: tmux attach -t netsuite"
echo "Detach with: Ctrl+B then D"
echo ""

# Wait for startup
echo "Waiting for servers to start..."
sleep 10

# Show tunnel URL
TUNNEL_URL=$(cat /tmp/cloudflare-new.log | grep -oE "https://[a-z0-9-]+\.trycloudflare\.com" | head -1)
echo ""
echo "Tunnel URL: $TUNNEL_URL"
echo ""
echo "If this is different from before, update the Cloudflare Worker!"
echo "(See PROXY-SOLUTION.md for instructions)"
echo ""
echo "================================================"

