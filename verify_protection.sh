#!/bin/bash
echo "🔍 VERIFYING REPOSITORY PROTECTION STATUS"
echo "========================================"

echo ""
echo "1️⃣ Git Remote Status:"
git remote -v
if [ $? -eq 0 ] && [ -z "$(git remote -v)" ]; then
    echo "✅ PASS: No remote connections found"
else
    echo "❌ FAIL: Remote connections still exist"
fi

echo ""
echo "2️⃣ Common Folder Check:"
if [ -d "common" ] && [ -f "common/assistant.py" ]; then
    echo "✅ PASS: Common folder exists with assistant.py"
else
    echo "❌ FAIL: Common folder missing or incomplete"
fi

echo ""
echo "3️⃣ GitHub Desktop Configuration:"
if [ -f ".github-desktop-ignore" ]; then
    echo "✅ PASS: GitHub Desktop ignore marker present"
else
    echo "ℹ️  INFO: No ignore marker (this is OK)"
fi

echo ""
echo "4️⃣ Repository Integrity:"
if git status > /dev/null 2>&1; then
    echo "✅ PASS: Git repository is healthy"
else
    echo "❌ FAIL: Git repository corrupted"
fi

echo ""
echo "🎯 RESULT: If all checks show PASS, your repository is fully protected!"
echo "🚀 You can safely run: streamlit run fdd_app.py --server.headless true --server.port 8501"
