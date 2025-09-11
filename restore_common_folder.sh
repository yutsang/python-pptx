#!/bin/bash
# 🚨 Emergency Restore Script for Common Folder
# Run this if you accidentally pulled changes that removed the common folder

echo "🔄 Restoring repository to protected state with common folder..."

# Reset to the protected commit
git reset --hard fee36c4

# Switch to main branch if not already there
git checkout main

echo "✅ Repository restored to protected state!"
echo "📁 Common folder should now be present in root directory"
echo ""
echo "If issues persist, you can also restore from backup branch:"
echo "  git checkout backup-with-common-folder"
echo ""
echo "Remember: Remote is now 'upstream', not 'origin'"
