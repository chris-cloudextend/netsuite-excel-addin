#!/bin/bash
#
# Deploy Frontend pointing to Flask Container Apps
# =================================================
#
# This script:
# 1. Reads the Flask Container Apps URL
# 2. Patches frontend files to use it
# 3. Uploads to Azure Storage
#

set -e

# ============================================================================
# CONFIGURATION
# ============================================================================

RESOURCE_GROUP="netsuite-excel-func-rg"
STORAGE_ACCOUNT="netsuiteexcelweb"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}XAVI Frontend - Container Apps Backend${NC}"
echo -e "${GREEN}========================================${NC}"
echo ""

# ============================================================================
# GET BACKEND URL
# ============================================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Try to read from saved file
if [ -f "$SCRIPT_DIR/../.flask_backend_url" ]; then
    FLASK_BACKEND_URL=$(cat "$SCRIPT_DIR/../.flask_backend_url")
    echo -e "${GREEN}✓ Found backend URL: ${FLASK_BACKEND_URL}${NC}"
else
    # Try to get from Azure
    FLASK_BACKEND_URL=$(az containerapp show \
        --name netsuite-flask \
        --resource-group $RESOURCE_GROUP \
        --query "properties.configuration.ingress.fqdn" -o tsv 2>/dev/null || echo "")
    
    if [ -n "$FLASK_BACKEND_URL" ]; then
        FLASK_BACKEND_URL="https://${FLASK_BACKEND_URL}"
        echo -e "${GREEN}✓ Found backend URL from Azure: ${FLASK_BACKEND_URL}${NC}"
    else
        echo -e "${RED}Error: Flask backend URL not found.${NC}"
        echo "Run ./deploy-flask-container.sh first."
        exit 1
    fi
fi

# Original Cloudflare URL (to replace)
CLOUDFLARE_URL="https://netsuite-proxy.chris-corcoran.workers.dev"

# ============================================================================
# CHECK PREREQUISITES
# ============================================================================

if ! az account show &> /dev/null; then
    echo -e "${YELLOW}Logging into Azure...${NC}"
    az login
fi

echo -e "${GREEN}✓ Logged into Azure${NC}"

# ============================================================================
# CREATE TEMP DIRECTORY AND COPY FILES
# ============================================================================

DOCS_DIR="$SCRIPT_DIR/../../docs"
TEMP_DIR=$(mktemp -d)

echo -e "${YELLOW}Creating patched frontend files...${NC}"

# Copy all files
cp -r "$DOCS_DIR"/* "$TEMP_DIR/"

# ============================================================================
# PATCH FILES FOR CONTAINER APPS
# ============================================================================

echo -e "${YELLOW}Patching URLs to use Flask Container Apps...${NC}"
echo "  Replacing: $CLOUDFLARE_URL"
echo "  With: $FLASK_BACKEND_URL"

# Patch functions.js
if [ -f "$TEMP_DIR/functions.js" ]; then
    sed -i.bak "s|$CLOUDFLARE_URL|$FLASK_BACKEND_URL|g" "$TEMP_DIR/functions.js"
    rm -f "$TEMP_DIR/functions.js.bak"
    echo -e "${GREEN}  ✓ Patched functions.js${NC}"
fi

# Patch commands.js
if [ -f "$TEMP_DIR/commands.js" ]; then
    sed -i.bak "s|$CLOUDFLARE_URL|$FLASK_BACKEND_URL|g" "$TEMP_DIR/commands.js"
    rm -f "$TEMP_DIR/commands.js.bak"
    echo -e "${GREEN}  ✓ Patched commands.js${NC}"
fi

# Patch taskpane.html
if [ -f "$TEMP_DIR/taskpane.html" ]; then
    # Replace SERVER_URL
    sed -i.bak "s|$CLOUDFLARE_URL|$FLASK_BACKEND_URL|g" "$TEMP_DIR/taskpane.html"
    
    # Replace display text
    sed -i.bak "s|netsuite-proxy.chris-corcoran.workers.dev|$(echo $FLASK_BACKEND_URL | sed 's|https://||')|g" "$TEMP_DIR/taskpane.html"
    
    # Update labels
    sed -i.bak "s|Proxy URL:|Backend URL:|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|To restart tunnel, run in Terminal:|Azure Container Apps Backend:|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|pkill cloudflared; cloudflared tunnel --url http://localhost:5002|Flask running on Azure Container Apps (scale-to-zero)|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|Update Worker with new tunnel URL|Azure Container Apps Info|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|After restarting tunnel, copy the new URL|Container Apps deployment is configured|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|Workers & Pages|Container Apps|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|netsuite-proxy</strong>|netsuite-flask</strong>|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|Paste new tunnel URL here|Container Apps URL|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|backend/tunnel is reachable|Flask backend is reachable|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|dash.cloudflare.com|portal.azure.com|g" "$TEMP_DIR/taskpane.html"
    sed -i.bak "s|Cloudflare Dashboard|Azure Portal|g" "$TEMP_DIR/taskpane.html"
    
    rm -f "$TEMP_DIR/taskpane.html.bak"
    echo -e "${GREEN}  ✓ Patched taskpane.html${NC}"
fi

# Verify
echo ""
echo -e "${YELLOW}Verifying patches...${NC}"
grep "SERVER_URL" "$TEMP_DIR/functions.js" | head -1

# ============================================================================
# UPLOAD FILES
# ============================================================================

echo ""
echo -e "${YELLOW}Uploading to Azure Storage...${NC}"

az storage blob upload-batch \
    --account-name $STORAGE_ACCOUNT \
    --destination '$web' \
    --source "$TEMP_DIR" \
    --overwrite \
    --output none

echo -e "${GREEN}✓ Files uploaded${NC}"

# ============================================================================
# CLEANUP
# ============================================================================

rm -rf "$TEMP_DIR"

# ============================================================================
# DONE
# ============================================================================

WEBSITE_URL="https://${STORAGE_ACCOUNT}.z13.web.core.windows.net"

echo ""
echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}Frontend Deployment Complete!${NC}"
echo -e "${GREEN}========================================${NC}"
echo ""
echo -e "${GREEN}Simplified Architecture:${NC}"
echo ""
echo "  ┌─────────────────┐"
echo "  │  Excel Add-in   │"
echo "  │  (Azure Storage)│"
echo "  └────────┬────────┘"
echo "           │"
echo "           ▼"
echo "  ┌─────────────────┐"
echo "  │ Flask Container │"
echo "  │ Apps (29 APIs)  │"
echo "  └────────┬────────┘"
echo "           │"
echo "           ▼"
echo "  ┌─────────────────┐"
echo "  │    NetSuite     │"
echo "  └─────────────────┘"
echo ""
echo -e "Frontend: ${YELLOW}${WEBSITE_URL}${NC}"
echo -e "Backend:  ${YELLOW}${FLASK_BACKEND_URL}${NC}"
echo ""
echo -e "${YELLOW}Refresh Excel to load the updated add-in${NC}"

