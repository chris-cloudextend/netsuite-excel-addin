#!/bin/bash
# Bump manifest version numbers to force Excel to reload files

MANIFEST="excel-addin/manifest-claude.xml"

if [ ! -f "$MANIFEST" ]; then
    echo "❌ Error: manifest-claude.xml not found"
    exit 1
fi

echo "🔄 Bumping manifest version numbers..."
echo ""

# Get current version from first occurrence
CURRENT_VERSION=$(grep -o 'v=[0-9]*' "$MANIFEST" | head -1 | cut -d'=' -f2)

if [ -z "$CURRENT_VERSION" ]; then
    echo "❌ Error: Could not find version number in manifest"
    exit 1
fi

NEW_VERSION=$((CURRENT_VERSION + 1))

echo "Current cache-bust version: v=$CURRENT_VERSION"
echo "New cache-bust version: v=$NEW_VERSION"
echo ""

# Replace all v=XXXX occurrences
sed -i '' "s/v=$CURRENT_VERSION/v=$NEW_VERSION/g" "$MANIFEST"

# Also bump the main Version tag
CURRENT_MAIN_VERSION=$(grep '<Version>' "$MANIFEST" | sed 's/.*<Version>\(.*\)<\/Version>.*/\1/')
if [ ! -z "$CURRENT_MAIN_VERSION" ]; then
    # Parse version (e.g., 1.0.0.94)
    MAJOR=$(echo $CURRENT_MAIN_VERSION | cut -d. -f1)
    MINOR=$(echo $CURRENT_MAIN_VERSION | cut -d. -f2)
    PATCH=$(echo $CURRENT_MAIN_VERSION | cut -d. -f3)
    BUILD=$(echo $CURRENT_MAIN_VERSION | cut -d. -f4)
    
    NEW_BUILD=$((BUILD + 1))
    NEW_MAIN_VERSION="$MAJOR.$MINOR.$PATCH.$NEW_BUILD"
    
    echo "Current main version: $CURRENT_MAIN_VERSION"
    echo "New main version: $NEW_MAIN_VERSION"
    echo ""
    
    sed -i '' "s/<Version>$CURRENT_MAIN_VERSION<\/Version>/<Version>$NEW_MAIN_VERSION<\/Version>/" "$MANIFEST"
fi

echo "✅ Manifest updated successfully!"
echo ""
echo "Next steps:"
echo "1. In Excel, go to Insert > Add-ins > My Add-ins"
echo "2. Remove the 'NetSuite Formulas' add-in"
echo "3. Upload the manifest again: $MANIFEST"
echo "4. The new version should force Excel to reload all files"
echo ""

