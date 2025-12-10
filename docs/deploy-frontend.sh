#!/bin/bash
#===============================================================================
# Azure Storage Static Website Deployment for Excel Add-in Frontend
#===============================================================================
#
# This script is IDEMPOTENT:
#   - Only creates storage account if it doesn't exist
#   - Overwrites existing files with latest versions
#   - Safe to run multiple times
#
# Cost: ~$0.02/GB/month
#
# Prerequisites:
#   - Azure CLI: brew install azure-cli
#   - Logged in: az login
#
#===============================================================================

set -e

#-------------------------------------------------------------------------------
# Configuration
#-------------------------------------------------------------------------------
RESOURCE_GROUP="netsuite-excel-func-rg"
LOCATION="eastus"
STORAGE_ACCOUNT="netsuiteexcelweb"
API_URL="https://netsuite-excel-func.azurewebsites.net"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'

print_header() {
    echo ""
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${BLUE}  $1${NC}"
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
}

print_step() { echo -e "${GREEN}▶ $1${NC}"; }
print_success() { echo -e "${GREEN}✔ $1${NC}"; }
print_skip() { echo -e "${CYAN}↷ $1${NC}"; }
print_warning() { echo -e "${YELLOW}⚠ $1${NC}"; }

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEPLOY_DIR="${SCRIPT_DIR}/.deploy-frontend"

#-------------------------------------------------------------------------------
# Cleanup on exit
#-------------------------------------------------------------------------------
cleanup() {
    rm -rf "${DEPLOY_DIR}" 2>/dev/null || true
}
trap cleanup EXIT

#-------------------------------------------------------------------------------
# Pre-flight Checks
#-------------------------------------------------------------------------------
print_header "Pre-flight Checks"

if ! command -v az &> /dev/null; then
    echo "Azure CLI not installed. Run: brew install azure-cli"
    exit 1
fi
print_success "Azure CLI installed"

if ! az account show &> /dev/null 2>&1; then
    echo "Not logged in to Azure. Run: az login"
    exit 1
fi
print_success "Logged in to Azure"

#-------------------------------------------------------------------------------
# Prepare Frontend Files
#-------------------------------------------------------------------------------
print_header "Preparing Files"

rm -rf "${DEPLOY_DIR}"
mkdir -p "${DEPLOY_DIR}"

FRONTEND_FILES=(
    "index.html"
    "config.js"
    "taskpane.html"
    "functions.html"
    "functions.js"
    "functions.json"
    "commands.html"
    "commands.js"
    "icon-16.png"
    "icon-32.png"
    "icon-64.png"
    "icon-80.png"
)

print_step "Copying files..."
for file in "${FRONTEND_FILES[@]}"; do
    if [[ -f "${SCRIPT_DIR}/${file}" ]]; then
        cp "${SCRIPT_DIR}/${file}" "${DEPLOY_DIR}/"
    fi
done

# Update API URLs
print_step "Updating API URL to: ${API_URL}"
if [[ -f "${DEPLOY_DIR}/functions.js" ]]; then
    sed -i.bak "s|http://localhost:5002|${API_URL}|g" "${DEPLOY_DIR}/functions.js"
    sed -i.bak "s|https://localhost:5002|${API_URL}|g" "${DEPLOY_DIR}/functions.js"
    rm -f "${DEPLOY_DIR}/functions.js.bak"
fi
if [[ -f "${DEPLOY_DIR}/commands.js" ]]; then
    sed -i.bak "s|http://localhost:5002|${API_URL}|g" "${DEPLOY_DIR}/commands.js"
    sed -i.bak "s|https://localhost:5002|${API_URL}|g" "${DEPLOY_DIR}/commands.js"
    rm -f "${DEPLOY_DIR}/commands.js.bak"
fi
print_success "Files prepared"

#-------------------------------------------------------------------------------
# Check/Create Azure Resources (Only if not exists)
#-------------------------------------------------------------------------------
print_header "Checking Azure Resources"

# Resource Group
print_step "Resource Group: ${RESOURCE_GROUP}"
if az group show --name "${RESOURCE_GROUP}" &>/dev/null; then
    print_skip "Already exists"
else
    print_step "Creating..."
    az group create --name "${RESOURCE_GROUP}" --location "${LOCATION}" --output none
    print_success "Created"
fi

# Storage Account
print_step "Storage Account: ${STORAGE_ACCOUNT}"
if az storage account show --name "${STORAGE_ACCOUNT}" --resource-group "${RESOURCE_GROUP}" &>/dev/null; then
    print_skip "Already exists"
else
    print_step "Creating..."
    az storage account create \
        --name "${STORAGE_ACCOUNT}" \
        --resource-group "${RESOURCE_GROUP}" \
        --location "${LOCATION}" \
        --sku Standard_LRS \
        --kind StorageV2 \
        --allow-blob-public-access true \
        --output none
    print_success "Created"
fi

# Enable static website (idempotent)
print_step "Enabling static website hosting..."
az storage blob service-properties update \
    --account-name "${STORAGE_ACCOUNT}" \
    --static-website \
    --index-document index.html \
    --404-document index.html \
    --output none 2>/dev/null
print_success "Static website enabled"

#-------------------------------------------------------------------------------
# Upload Files (Overwrite existing)
#-------------------------------------------------------------------------------
print_header "Uploading Files"

# Get storage key
STORAGE_KEY=$(az storage account keys list \
    --account-name "${STORAGE_ACCOUNT}" \
    --resource-group "${RESOURCE_GROUP}" \
    --query "[0].value" -o tsv)

# Upload all files (--overwrite replaces existing)
print_step "Uploading to \$web container..."
az storage blob upload-batch \
    --account-name "${STORAGE_ACCOUNT}" \
    --account-key "${STORAGE_KEY}" \
    --destination '$web' \
    --source "${DEPLOY_DIR}" \
    --overwrite \
    --output none

print_success "All files uploaded"

# Set content types
print_step "Setting content types..."
for ext in html js json png; do
    case $ext in
        html) content_type="text/html" ;;
        js) content_type="application/javascript" ;;
        json) content_type="application/json" ;;
        png) content_type="image/png" ;;
    esac
    
    for file in "${DEPLOY_DIR}"/*."${ext}"; do
        if [[ -f "$file" ]]; then
            filename=$(basename "$file")
            az storage blob update \
                --account-name "${STORAGE_ACCOUNT}" \
                --account-key "${STORAGE_KEY}" \
                --container-name '$web' \
                --name "${filename}" \
                --content-type "${content_type}" \
                --output none 2>/dev/null || true
        fi
    done
done
print_success "Content types set"

#-------------------------------------------------------------------------------
# Summary
#-------------------------------------------------------------------------------
print_header "Deployment Complete! 🎉"

WEB_URL="https://${STORAGE_ACCOUNT}.z13.web.core.windows.net"

echo ""
echo -e "${GREEN}Frontend URL:${NC} ${WEB_URL}"
echo -e "${GREEN}Functions:${NC}   ${WEB_URL}/functions.html"
echo -e "${GREEN}Taskpane:${NC}    ${WEB_URL}/taskpane.html"
echo ""
echo -e "${CYAN}Commands:${NC}"
echo "  Redeploy:  ./deploy-frontend.sh"
echo "  List files: az storage blob list --account-name ${STORAGE_ACCOUNT} --container-name '\$web' -o table"
echo ""
