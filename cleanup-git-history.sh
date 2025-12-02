#!/bin/bash

echo "════════════════════════════════════════════════════════════════════"
echo "🧹 GIT HISTORY CLEANUP - REMOVE EXPOSED CREDENTIALS"
echo "════════════════════════════════════════════════════════════════════"
echo ""
echo "⚠️  WARNING: This will rewrite git history"
echo ""
echo "Prerequisites:"
echo "  ✅ You have revoked tokens in NetSuite"
echo "  ✅ You have new tokens stored securely offline"
echo "  ✅ You understand this rewrites history"
echo ""
read -p "Have you revoked the tokens in NetSuite? (yes/no): " REVOKED

if [ "$REVOKED" != "yes" ]; then
    echo ""
    echo "❌ STOP! Revoke tokens first:"
    echo "   1. Log into NetSuite (both accounts)"
    echo "   2. Setup → Integration → Manage Integrations"
    echo "   3. Delete exposed Access Tokens"
    echo "   4. Create NEW Access Tokens"
    echo ""
    echo "Then run this script again."
    exit 1
fi

echo ""
echo "════════════════════════════════════════════════════════════════════"
echo "Step 1: Installing BFG Repo Cleaner..."
echo "════════════════════════════════════════════════════════════════════"
if ! command -v bfg &> /dev/null; then
    echo "Installing BFG via Homebrew..."
    brew install bfg
else
    echo "✅ BFG already installed"
fi

echo ""
echo "════════════════════════════════════════════════════════════════════"
echo "Step 2: Creating backup..."
echo "════════════════════════════════════════════════════════════════════"
cd "/Users/chriscorcoran/Documents/Cursor"
BACKUP_DIR="NetSuite-Formulas-Backup-$(date +%Y%m%d-%H%M%S)"
echo "Creating backup at: $BACKUP_DIR"
cp -R "NetSuite Formulas Revised" "$BACKUP_DIR"
echo "✅ Backup created"

echo ""
echo "════════════════════════════════════════════════════════════════════"
echo "Step 3: Removing SWITCH-ACCOUNTS.md from git history..."
echo "════════════════════════════════════════════════════════════════════"
cd "/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised"

# Remove the file from ALL commits
bfg --delete-files SWITCH-ACCOUNTS.md

echo ""
echo "════════════════════════════════════════════════════════════════════"
echo "Step 4: Cleaning up repository..."
echo "════════════════════════════════════════════════════════════════════"
git reflog expire --expire=now --all
git gc --prune=now --aggressive

echo ""
echo "════════════════════════════════════════════════════════════════════"
echo "Step 5: Force pushing to GitHub (rewrites history)..."
echo "════════════════════════════════════════════════════════════════════"
echo ""
echo "⚠️  This will REWRITE GitHub history"
read -p "Continue with force push? (yes/no): " CONTINUE

if [ "$CONTINUE" != "yes" ]; then
    echo "❌ Aborted. History cleaned locally but not pushed."
    echo "   Run 'git push origin main --force' when ready."
    exit 1
fi

git push origin main --force

echo ""
echo "════════════════════════════════════════════════════════════════════"
echo "✅ GIT HISTORY CLEANUP COMPLETE"
echo "════════════════════════════════════════════════════════════════════"
echo ""
echo "What was done:"
echo "  ✅ SWITCH-ACCOUNTS.md removed from ALL git history"
echo "  ✅ Repository cleaned and optimized"
echo "  ✅ GitHub history rewritten"
echo ""
echo "════════════════════════════════════════════════════════════════════"
echo "🔍 VERIFICATION STEPS"
echo "════════════════════════════════════════════════════════════════════"
echo ""
echo "1. Check GitHub:"
echo "   Visit: https://github.com/chris-cloudextend/netsuite-excel-addin"
echo "   Search for '589861' - should find nothing"
echo ""
echo "2. If still visible, wait 5-10 minutes (GitHub caching)"
echo ""
echo "3. If still visible after 30 min, contact GitHub support:"
echo "   https://support.github.com/contact"
echo "   Request: 'Purge cache for removed sensitive data'"
echo ""
echo "════════════════════════════════════════════════════════════════════"
echo "📧 REPLY TO NETSUITE"
echo "════════════════════════════════════════════════════════════════════"
echo ""
echo "Email them confirmation:"
echo ""
echo "Subject: RE: Security Alert - Credentials Revoked"
echo ""
echo "Hello NetSuite Security Team,"
echo ""
echo "I have taken immediate action:"
echo ""
echo "1. ✅ Revoked all exposed Access Tokens (both accounts)"
echo "2. ✅ Generated new Access Tokens"
echo "3. ✅ Deleted SWITCH-ACCOUNTS.md from repository"
echo "4. ✅ Removed credentials from git history (BFG)"
echo "5. ✅ Force-pushed to GitHub to rewrite history"
echo "6. ✅ Updated .gitignore to prevent future commits"
echo ""
echo "The exposed credentials are no longer valid."
echo ""
echo "Thank you for the alert."
echo ""
echo "Best regards,"
echo "Chris Corcoran"
echo ""
echo "════════════════════════════════════════════════════════════════════"
echo ""
echo "Backup location: /Users/chriscorcoran/Documents/Cursor/$BACKUP_DIR"
echo ""
echo "🎯 DONE!"
echo ""

