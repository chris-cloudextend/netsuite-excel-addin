FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn python-dateutil

COPY server.py .
COPY constants.py .

# Create startup script that generates config from env vars
RUN echo '#!/bin/bash\n\
cat > netsuite_config.json << EOF\n\
{\n\
    "account_id": "${NETSUITE_ACCOUNT_ID:-configure-me}",\n\
    "consumer_key": "${NETSUITE_CONSUMER_KEY:-configure-me}",\n\
    "consumer_secret": "${NETSUITE_CONSUMER_SECRET:-configure-me}",\n\
    "token_id": "${NETSUITE_TOKEN_ID:-configure-me}",\n\
    "token_secret": "${NETSUITE_TOKEN_SECRET:-configure-me}"\n\
}\n\
EOF\n\
exec gunicorn --bind 0.0.0.0:5002 --workers 2 --timeout 300 server:app\n\
' > /app/start.sh && chmod +x /app/start.sh

EXPOSE 5002

CMD ["/app/start.sh"]

