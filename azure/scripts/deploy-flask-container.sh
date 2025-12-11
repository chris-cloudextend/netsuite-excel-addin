#!/bin/bash
#
# Deploy Flask Backend to Azure Container Apps
# =============================================
#
# Cost-optimized for testing:
# - Scale-to-zero when idle ($0)
# - Pay only when processing requests
# - ~2-5 second cold start
#
# Estimated cost: $0-5/month for occasional testing
#

set -e

# ============================================================================
# CONFIGURATION
# ============================================================================

RESOURCE_GROUP="netsuite-excel-func-rg"
LOCATION="eastus"
CONTAINER_APP_NAME="netsuite-flask"
CONTAINER_ENV_NAME="netsuite-env"
CONTAINER_REGISTRY="netsuiteacr"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}Flask Backend - Azure Container Apps${NC}"
echo -e "${GREEN}(Scale-to-Zero for Testing)${NC}"
echo -e "${GREEN}========================================${NC}"
echo ""

# ============================================================================
# CHECK PREREQUISITES
# ============================================================================

if ! command -v az &> /dev/null; then
    echo -e "${RED}Error: Azure CLI not installed${NC}"
    exit 1
fi

if ! az account show &> /dev/null; then
    echo -e "${YELLOW}Logging into Azure...${NC}"
    az login
fi

SUBSCRIPTION=$(az account show --query name -o tsv)
echo -e "${GREEN}✓ Azure: ${SUBSCRIPTION}${NC}"

# ============================================================================
# GET NETSUITE CREDENTIALS
# ============================================================================

echo ""
echo -e "${YELLOW}NetSuite Credentials${NC}"
echo "These will be stored securely as environment variables."
echo ""

# Check if credentials exist in local config file
LOCAL_CONFIG="$SCRIPT_DIR/../../backend/netsuite_config.json"

if [ -f "$LOCAL_CONFIG" ]; then
    echo -e "${GREEN}Found local netsuite_config.json. Reading credentials...${NC}"
    
    NETSUITE_ACCOUNT_ID=$(python3 -c "import json; print(json.load(open('$LOCAL_CONFIG')).get('account_id', ''))" 2>/dev/null || echo "")
    NETSUITE_CONSUMER_KEY=$(python3 -c "import json; print(json.load(open('$LOCAL_CONFIG')).get('consumer_key', ''))" 2>/dev/null || echo "")
    NETSUITE_CONSUMER_SECRET=$(python3 -c "import json; print(json.load(open('$LOCAL_CONFIG')).get('consumer_secret', ''))" 2>/dev/null || echo "")
    NETSUITE_TOKEN_ID=$(python3 -c "import json; print(json.load(open('$LOCAL_CONFIG')).get('token_id', ''))" 2>/dev/null || echo "")
    NETSUITE_TOKEN_SECRET=$(python3 -c "import json; print(json.load(open('$LOCAL_CONFIG')).get('token_secret', ''))" 2>/dev/null || echo "")
    
    if [ -n "$NETSUITE_ACCOUNT_ID" ]; then
        echo -e "${GREEN}  ✓ Loaded credentials from local config${NC}"
    fi
fi

# If credentials not found, use placeholders (configure in Azure Portal later)
if [ -z "$NETSUITE_ACCOUNT_ID" ] || [ "$NETSUITE_ACCOUNT_ID" == "" ] || [ "$NETSUITE_ACCOUNT_ID" == "YOUR_ACCOUNT_ID" ]; then
    echo -e "${YELLOW}Credentials not found locally.${NC}"
    echo "Deploying with placeholder values - configure in Azure Portal after deployment."
    NETSUITE_ACCOUNT_ID="configure-in-azure-portal"
    NETSUITE_CONSUMER_KEY="configure-in-azure-portal"
    NETSUITE_CONSUMER_SECRET="configure-in-azure-portal"
    NETSUITE_TOKEN_ID="configure-in-azure-portal"
    NETSUITE_TOKEN_SECRET="configure-in-azure-portal"
fi

# Use placeholders if not provided
NETSUITE_ACCOUNT_ID=${NETSUITE_ACCOUNT_ID:-"configure-later"}
NETSUITE_CONSUMER_KEY=${NETSUITE_CONSUMER_KEY:-"configure-later"}
NETSUITE_CONSUMER_SECRET=${NETSUITE_CONSUMER_SECRET:-"configure-later"}
NETSUITE_TOKEN_ID=${NETSUITE_TOKEN_ID:-"configure-later"}
NETSUITE_TOKEN_SECRET=${NETSUITE_TOKEN_SECRET:-"configure-later"}

# ============================================================================
# CREATE CONTAINER ENVIRONMENT
# ============================================================================

echo ""
echo -e "${YELLOW}Creating Container Apps environment...${NC}"

# Check if environment exists
if ! az containerapp env show --name $CONTAINER_ENV_NAME --resource-group $RESOURCE_GROUP &>/dev/null; then
    az containerapp env create \
        --name $CONTAINER_ENV_NAME \
        --resource-group $RESOURCE_GROUP \
        --location $LOCATION \
        --output none
fi
echo -e "${GREEN}✓ Environment: ${CONTAINER_ENV_NAME}${NC}"

# ============================================================================
# BUILD AND DEPLOY CONTAINER
# ============================================================================

echo -e "${YELLOW}Deploying Flask backend...${NC}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BACKEND_DIR="$SCRIPT_DIR/../../backend"

# Create a temporary Dockerfile
cat > "$BACKEND_DIR/Dockerfile" << 'EOF'
FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install gunicorn

# Copy application
COPY server.py .
COPY constants.py .

# Create a default config (will be overridden by env vars)
RUN echo '{}' > netsuite_config.json

# Expose port
EXPOSE 5002

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5002/health || exit 1

# Run with gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:5002", "--workers", "2", "--timeout", "120", "server:app"]
EOF

echo -e "${GREEN}✓ Dockerfile created${NC}"

# Deploy using az containerapp up (handles build + deploy)
az containerapp up \
    --name $CONTAINER_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --environment $CONTAINER_ENV_NAME \
    --source "$BACKEND_DIR" \
    --ingress external \
    --target-port 5002 \
    --env-vars \
        NETSUITE_ACCOUNT_ID="$NETSUITE_ACCOUNT_ID" \
        NETSUITE_CONSUMER_KEY="$NETSUITE_CONSUMER_KEY" \
        NETSUITE_CONSUMER_SECRET="$NETSUITE_CONSUMER_SECRET" \
        NETSUITE_TOKEN_ID="$NETSUITE_TOKEN_ID" \
        NETSUITE_TOKEN_SECRET="$NETSUITE_TOKEN_SECRET"

echo -e "${GREEN}✓ Container deployed${NC}"

# ============================================================================
# CONFIGURE SCALE-TO-ZERO
# ============================================================================

echo -e "${YELLOW}Configuring scale-to-zero (cost optimization)...${NC}"

az containerapp update \
    --name $CONTAINER_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --min-replicas 0 \
    --max-replicas 3 \
    --output none

echo -e "${GREEN}✓ Scale-to-zero enabled${NC}"

# ============================================================================
# GET CONTAINER URL
# ============================================================================

CONTAINER_URL=$(az containerapp show \
    --name $CONTAINER_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --query "properties.configuration.ingress.fqdn" -o tsv)

echo ""
echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}Deployment Complete!${NC}"
echo -e "${GREEN}========================================${NC}"
echo ""
echo -e "Flask Backend URL: ${YELLOW}https://${CONTAINER_URL}${NC}"
echo ""

# Save the URL for frontend deployment
echo "https://${CONTAINER_URL}" > "$SCRIPT_DIR/../.flask_backend_url"
echo -e "${GREEN}✓ Backend URL saved for frontend deployment${NC}"

echo ""
echo -e "${GREEN}Architecture (Simplified):${NC}"
echo "  Excel Add-in → Flask Container Apps → NetSuite"
echo "  (No Azure Functions layer)"
echo ""
echo -e "${GREEN}Cost Optimization:${NC}"
echo "  • Min replicas: 0 (scales to zero when idle)"
echo "  • Max replicas: 3 (handles bursts)"
echo "  • Estimated cost: \$0-5/month for testing"
echo ""
echo -e "${YELLOW}Commands:${NC}"
echo "  Stop (save costs):  az containerapp stop --name $CONTAINER_APP_NAME --resource-group $RESOURCE_GROUP"
echo "  Start:              az containerapp start --name $CONTAINER_APP_NAME --resource-group $RESOURCE_GROUP"
echo ""
echo -e "${YELLOW}Next Step:${NC}"
echo "  Run ./deploy-frontend-container.sh to update frontend"
echo ""
echo -e "Test: ${GREEN}curl https://${CONTAINER_URL}/health${NC}"

# Cleanup Dockerfile
rm -f "$BACKEND_DIR/Dockerfile"

