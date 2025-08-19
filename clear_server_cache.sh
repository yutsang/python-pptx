#!/bin/bash
# Script to clear all caches on the server

echo "🧹 Clearing server caches..."

# Clear Python cache
echo "Clearing Python cache..."
find . -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null || true
find . -name "*.pyc" -delete 2>/dev/null || true

# Clear application cache
echo "Clearing application cache..."
rm -rf cache/ 2>/dev/null || true

# Clear Streamlit cache
echo "Clearing Streamlit cache..."
rm -rf .streamlit/ 2>/dev/null || true

# Clear any temporary files
echo "Clearing temporary files..."
rm -f *.tmp 2>/dev/null || true
rm -f temp_*.xlsx 2>/dev/null || true

echo "✅ Server cache cleared!"
echo "🔄 Please restart your Streamlit application now."
