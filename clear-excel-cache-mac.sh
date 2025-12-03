#!/bin/bash
# Clear Excel Add-in Cache on Mac

echo "🧹 Clearing Excel Add-in Cache..."

# Close Excel if running
echo "Closing Excel..."
osascript -e 'quit app "Microsoft Excel"' 2>/dev/null

# Wait for Excel to close
sleep 2

# Clear the Office Web Extension Framework (WEF) cache
echo "Clearing WEF cache..."
WEF_PATH="$HOME/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
if [ -d "$WEF_PATH" ]; then
    rm -rf "$WEF_PATH"/*
    echo "✅ WEF cache cleared"
else
    echo "⚠️  WEF cache directory not found at: $WEF_PATH"
fi

# Clear the general Office cache
echo "Clearing Office cache..."
OFFICE_CACHE="$HOME/Library/Containers/com.microsoft.Excel/Data/Library/Caches"
if [ -d "$OFFICE_CACHE" ]; then
    rm -rf "$OFFICE_CACHE"/*
    echo "✅ Office cache cleared"
else
    echo "⚠️  Office cache directory not found"
fi

echo ""
echo "✅ Cache cleared successfully!"
echo "Now:"
echo "1. Restart Excel"
echo "2. Go to Insert > Add-ins > My Add-ins"
echo "3. Remove the NetSuite Formulas add-in"
echo "4. Re-add it by uploading the manifest-claude.xml file"
echo ""

