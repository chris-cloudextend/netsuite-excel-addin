#!/bin/bash
#===============================================================================
# Azure Functions Deployment Script for NetSuite Excel Add-in Backend
#===============================================================================
#
# This script is IDEMPOTENT:
#   - Only creates resources if they don't exist
#   - Updates code in-place without recreating resources
#   - Safe to run multiple times
#
# Cost: ~$0/month (Consumption Plan - 1M free executions)
#
# Prerequisites:
#   - Azure CLI: brew install azure-cli
#   - Azure Functions Core Tools: brew install azure-functions-core-tools@4
#   - Logged in: az login
#
#===============================================================================

set -e

#-------------------------------------------------------------------------------
# Configuration
#-------------------------------------------------------------------------------
RESOURCE_GROUP="netsuite-excel-func-rg"
LOCATION="eastus"
STORAGE_ACCOUNT="netsuiteexcelstor"
FUNCTION_APP="netsuite-excel-func"
PYTHON_VERSION="3.11"

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
print_error() { echo -e "${RED}✖ $1${NC}"; }

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEPLOY_DIR="${SCRIPT_DIR}/.deploy"

#-------------------------------------------------------------------------------
# Cleanup on exit
#-------------------------------------------------------------------------------
cleanup() {
    rm -rf "${DEPLOY_DIR}" 2>/dev/null || true
    find "${SCRIPT_DIR}" -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true
}
trap cleanup EXIT

#-------------------------------------------------------------------------------
# Pre-flight Checks
#-------------------------------------------------------------------------------
print_header "Pre-flight Checks"

# Check Azure CLI
if ! command -v az &> /dev/null; then
    print_error "Azure CLI not installed. Run: brew install azure-cli"
    exit 1
fi
print_success "Azure CLI installed"

# Check Azure Functions Core Tools
if ! command -v func &> /dev/null; then
    print_error "Azure Functions Core Tools not installed. Run: brew install azure-functions-core-tools@4"
    exit 1
fi
print_success "Azure Functions Core Tools installed"

# Check if logged in
if ! az account show &> /dev/null 2>&1; then
    print_error "Not logged in to Azure. Run: az login"
    exit 1
fi
print_success "Logged in to Azure"

SUBSCRIPTION=$(az account show --query name -o tsv)
echo -e "  Subscription: ${YELLOW}${SUBSCRIPTION}${NC}"

#-------------------------------------------------------------------------------
# Check/Create Azure Resources (Only if not exists)
#-------------------------------------------------------------------------------
print_header "Checking Azure Resources"

# Resource Group
print_step "Resource Group: ${RESOURCE_GROUP}"
if az group show --name "${RESOURCE_GROUP}" &>/dev/null; then
    print_skip "Already exists, skipping creation"
else
    print_step "Creating Resource Group..."
    az group create --name "${RESOURCE_GROUP}" --location "${LOCATION}" --output none
    print_success "Created"
fi

# Storage Account
print_step "Storage Account: ${STORAGE_ACCOUNT}"
if az storage account show --name "${STORAGE_ACCOUNT}" --resource-group "${RESOURCE_GROUP}" &>/dev/null; then
    print_skip "Already exists, skipping creation"
else
    print_step "Creating Storage Account..."
    az storage account create \
        --name "${STORAGE_ACCOUNT}" \
        --resource-group "${RESOURCE_GROUP}" \
        --location "${LOCATION}" \
        --sku Standard_LRS \
        --kind StorageV2 \
        --output none
    print_success "Created"
fi

# Function App
print_step "Function App: ${FUNCTION_APP}"
if az functionapp show --name "${FUNCTION_APP}" --resource-group "${RESOURCE_GROUP}" &>/dev/null; then
    print_skip "Already exists, skipping creation"
else
    print_step "Creating Function App (Consumption Plan)..."
    az functionapp create \
        --name "${FUNCTION_APP}" \
        --resource-group "${RESOURCE_GROUP}" \
        --storage-account "${STORAGE_ACCOUNT}" \
        --consumption-plan-location "${LOCATION}" \
        --runtime python \
        --runtime-version "${PYTHON_VERSION}" \
        --functions-version 4 \
        --os-type Linux \
        --output none
    print_success "Created"
    
    # Wait for new Function App to be ready
    echo "  Waiting for Function App to initialize..."
    sleep 10
fi

#-------------------------------------------------------------------------------
# Configure Settings (Only if needed)
#-------------------------------------------------------------------------------
print_header "Checking Configuration"

# Check if credentials are configured
CURRENT_ACCOUNT_ID=$(az functionapp config appsettings list \
    --name "${FUNCTION_APP}" \
    --resource-group "${RESOURCE_GROUP}" \
    --query "[?name=='NETSUITE_ACCOUNT_ID'].value" -o tsv 2>/dev/null || echo "")

if [[ -z "$CURRENT_ACCOUNT_ID" ]]; then
    print_step "Setting placeholder environment variables..."
    az functionapp config appsettings set \
        --name "${FUNCTION_APP}" \
        --resource-group "${RESOURCE_GROUP}" \
        --settings \
            NETSUITE_ACCOUNT_ID="CONFIGURE_IN_PORTAL" \
            NETSUITE_CONSUMER_KEY="CONFIGURE_IN_PORTAL" \
            NETSUITE_CONSUMER_SECRET="CONFIGURE_IN_PORTAL" \
            NETSUITE_TOKEN_ID="CONFIGURE_IN_PORTAL" \
            NETSUITE_TOKEN_SECRET="CONFIGURE_IN_PORTAL" \
        --output none
    print_warning "Credentials need to be configured in Azure Portal"
elif [[ "$CURRENT_ACCOUNT_ID" == "CONFIGURE_IN_PORTAL" ]]; then
    print_warning "Credentials still need to be configured in Azure Portal"
else
    print_success "NetSuite credentials already configured"
fi

# Ensure CORS is configured
print_step "Checking CORS configuration..."
az functionapp cors add \
    --name "${FUNCTION_APP}" \
    --resource-group "${RESOURCE_GROUP}" \
    --allowed-origins "*" \
    --output none 2>/dev/null || true
print_success "CORS configured"

# Ensure public access
print_step "Ensuring public web access..."
az functionapp update \
    --name "${FUNCTION_APP}" \
    --resource-group "${RESOURCE_GROUP}" \
    --set publicNetworkAccess=Enabled \
    --output none 2>/dev/null || true
print_success "Public access enabled"

#-------------------------------------------------------------------------------
# Prepare & Deploy Code
#-------------------------------------------------------------------------------
print_header "Deploying Code"

# Prepare clean deployment package
print_step "Preparing deployment package..."
rm -rf "${DEPLOY_DIR}"
mkdir -p "${DEPLOY_DIR}"

PRODUCTION_FILES=("function_app.py" "host.json" "requirements.txt" "constants.py")
for file in "${PRODUCTION_FILES[@]}"; do
    if [[ -f "${SCRIPT_DIR}/${file}" ]]; then
        cp "${SCRIPT_DIR}/${file}" "${DEPLOY_DIR}/"
    fi
done
print_success "Package prepared (${#PRODUCTION_FILES[@]} files)"

# Deploy code
cd "${DEPLOY_DIR}"
print_step "Publishing to Azure Functions..."
echo ""
func azure functionapp publish "${FUNCTION_APP}" --python --build remote --force
print_success "Code deployed!"

#-------------------------------------------------------------------------------
# Verify Deployment
#-------------------------------------------------------------------------------
print_header "Verifying Deployment"

FUNC_URL="https://${FUNCTION_APP}.azurewebsites.net"

print_step "Testing health endpoint..."
sleep 3
HTTP_STATUS=$(curl -s -o /dev/null -w "%{http_code}" "${FUNC_URL}/health" --max-time 30 2>/dev/null || echo "000")

if [[ "$HTTP_STATUS" == "200" ]]; then
    print_success "Health check passed! (HTTP ${HTTP_STATUS})"
elif [[ "$HTTP_STATUS" == "500" ]]; then
    print_warning "Function running but needs NetSuite credentials"
else
    print_warning "Health check returned HTTP ${HTTP_STATUS} (may still be warming up)"
fi

#-------------------------------------------------------------------------------
# Summary
#-------------------------------------------------------------------------------
print_header "Deployment Complete! 🎉"

echo ""
echo -e "${GREEN}API URL:${NC} ${FUNC_URL}"
echo -e "${GREEN}Health:${NC}  ${FUNC_URL}/health"
echo ""

if [[ -z "$CURRENT_ACCOUNT_ID" || "$CURRENT_ACCOUNT_ID" == "CONFIGURE_IN_PORTAL" ]]; then
    echo -e "${YELLOW}⚠️  Configure NetSuite credentials in Azure Portal:${NC}"
    echo "   Function App → ${FUNCTION_APP} → Settings → Environment variables"
    echo ""
fi

echo -e "${CYAN}Commands:${NC}"
echo "  Redeploy:  ./deploy-functions.sh"
echo "  Logs:      func azure functionapp logstream ${FUNCTION_APP}"
echo "  Delete:    az group delete --name ${RESOURCE_GROUP} --yes"
echo ""
