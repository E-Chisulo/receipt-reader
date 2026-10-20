#!/bin/zsh
# Receipt Reader — HTTP (Mac browser / localhost)

cd "$(dirname "$0")"

# Start Ollama if not running
if ! pgrep -x ollama > /dev/null; then
  echo "Starting Ollama..."
  ollama serve &>/dev/null &
  sleep 2
fi

# Activate venv
source .venv/bin/activate

IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "unavailable")
echo ""
echo "  Receipt Reader starting..."
echo "  On this Mac:  http://localhost:8501"
echo "  On phone:     http://$IP:8501"
echo ""
echo "  Note: camera upload works over HTTP. For live camera preview on iPhone use start_https.sh"
echo ""

streamlit run app.py \
  --server.port 8501 \
  --server.headless true \
  --browser.gatherUsageStats false
