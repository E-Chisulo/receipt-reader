#!/bin/zsh
# Receipt Reader — HTTPS (iPhone camera via browser)

cd "$(dirname "$0")"

# Check certs exist
if [[ ! -f cert.pem || ! -f key.pem ]]; then
  echo "SSL certs not found. Generating self-signed cert..."
  IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "127.0.0.1")
  openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 -nodes \
    -subj "/CN=$IP" \
    -addext "subjectAltName=IP:$IP,IP:127.0.0.1" 2>/dev/null
  echo "Cert created for IP: $IP"
fi

# Start Ollama if not running
if ! pgrep -x ollama > /dev/null; then
  echo "Starting Ollama..."
  ollama serve &>/dev/null &
  sleep 2
fi

# Activate venv
source .venv/bin/activate

# Show iPhone URL
IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "127.0.0.1")
echo ""
echo "  Receipt Reader starting (HTTPS)..."
echo "  On this Mac:  https://localhost:8501"
echo "  On iPhone:    https://$IP:8501"
echo ""
echo "  (Accept the security warning in Safari — your own cert)"
echo ""

streamlit run app.py \
  --server.port 8501 \
  --server.headless true \
  --browser.gatherUsageStats false \
  --server.sslCertFile cert.pem \
  --server.sslKeyFile key.pem
