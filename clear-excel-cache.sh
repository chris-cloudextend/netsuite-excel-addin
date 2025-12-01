#!/bin/bash

echo "================================================"
echo "🧹 CLEARING EXCEL ADD-IN CACHE"
echo "================================================"
echo ""

CACHE_DIR=~/Library/Containers/com.microsoft.Excel/Data/Library/Application\ Support/Microsoft/Office/16.0/Wef

if [ -d "$CACHE_DIR" ]; then
    echo "Found Excel add-in cache directory:"
    echo "$CACHE_DIR"
    echo ""
    echo "Removing cache..."
    rm -rf "$CACHE_DIR"
    echo "✅ Cache cleared!"
else
    echo "⚠️  Cache directory not found (maybe already clean)"
fi

echo ""
echo "================================================"
echo "📋 NEXT STEPS:"
echo "================================================"
echo ""
echo "1. Close Excel if still open (Cmd+Q)"
echo "2. Wait 10 seconds"
echo "3. Reopen Excel"
echo "4. Insert → My Add-ins → Shared Folder"
echo "5. Upload manifest-claude.xml (v1.0.0.60)"
echo "6. Test formulas!"
echo ""
echo "================================================"

