#!/usr/bin/env python3
"""
Quick Status Check
Run this to diagnose AI provider issues
"""

import sys
from pathlib import Path
import json

# Add parent directory to path
current_dir = Path(__file__).resolve().parent
parent_dir = current_dir.parent
if str(parent_dir) not in sys.path:
    sys.path.insert(0, str(parent_dir))

print("="*80)
print(" AI MODULE STATUS CHECK")
print("="*80)

# Check 1: Configuration file
print("\n📋 Step 1: Checking configuration...")
try:
    config_path = current_dir / "config.json"
    with open(config_path, 'r') as f:
        config = json.load(f)
    print(f"✅ Config loaded from: {config_path}")
    
    print(f"\n   DeepSeek API Key: {'sk-****' + config.get('DEEPSEEK_API_KEY', '')[-4:] if config.get('DEEPSEEK_API_KEY') else '❌ NOT SET'}")
    print(f"   DeepSeek API Base: {config.get('DEEPSEEK_API_BASE', '❌ NOT SET')}")
    print(f"   OpenAI API Key: {'✅ SET' if config.get('OPENAI_API_KEY') and 'placeholder' not in config.get('OPENAI_API_KEY', '').lower() else '❌ NOT SET'}")
    print(f"   Local AI Enabled: {'✅ YES' if config.get('LOCAL_AI_ENABLED') else '❌ NO'}")
    print(f"   Local AI Base: {config.get('LOCAL_AI_API_BASE', '❌ NOT SET')}")
    print(f"   Default Provider: {config.get('DEFAULT_AI_PROVIDER', '❌ NOT SET')}")
except Exception as e:
    print(f"❌ Error loading config: {e}")
    sys.exit(1)

# Check 2: Import AI module
print("\n📦 Step 2: Importing AI module...")
try:
    from ai_module import AIModule
    print("✅ AI module imported successfully")
except Exception as e:
    print(f"❌ Error importing: {e}")
    sys.exit(1)

# Check 3: Initialize module
print("\n🔧 Step 3: Initializing AI module...")
try:
    ai = AIModule()
    print(f"✅ AI module initialized")
    print(f"   Available providers: {list(ai.clients.keys())}")
except Exception as e:
    print(f"❌ Error initializing: {e}")
    sys.exit(1)

# Check 4: Test each provider
print("\n🧪 Step 4: Testing each provider...")
test_results = {}

for provider in ai.clients.keys():
    print(f"\n   Testing {provider}...")
    try:
        result = ai.generate_content(
            system_prompt="You are a helpful assistant.",
            user_prompt="Say 'test' in 1 word.",
            provider=provider,
            max_tokens=10
        )
        
        if result.get('content'):
            test_results[provider] = "✅ WORKING"
            print(f"   ✅ {provider}: WORKING")
            print(f"      Response: {result['content']}")
            print(f"      Tokens: {result['tokens']['total_tokens']}")
        else:
            test_results[provider] = f"❌ FAILED"
            error_msg = str(result.get('error', 'Unknown error'))
            print(f"   ❌ {provider}: FAILED")
            print(f"      Error: {error_msg[:100]}")
            
            # Provide specific guidance
            if '401' in error_msg or 'Authentication' in error_msg:
                print(f"      💡 Fix: API key is invalid. Get a new key.")
            elif 'Connection refused' in error_msg or 'Connection error' in error_msg:
                print(f"      💡 Fix: Local server not running. Start it on port 1234.")
            elif 'Rate limit' in error_msg:
                print(f"      💡 Fix: Rate limit exceeded. Wait or upgrade plan.")
    except Exception as e:
        test_results[provider] = f"❌ ERROR"
        print(f"   ❌ {provider}: ERROR - {str(e)[:100]}")

# Final Summary
print("\n" + "="*80)
print(" FINAL STATUS SUMMARY")
print("="*80)

working_providers = [p for p, status in test_results.items() if 'WORKING' in status]
failed_providers = [p for p, status in test_results.items() if 'WORKING' not in status]

print(f"\n✅ Working Providers: {len(working_providers)}")
for p in working_providers:
    print(f"   • {p}")

print(f"\n❌ Failed Providers: {len(failed_providers)}")
for p in failed_providers:
    print(f"   • {p}: {test_results[p]}")

print("\n" + "="*80)

if working_providers:
    print(f"\n🎉 SUCCESS! You can use: {', '.join(working_providers)}")
    print(f"\n💡 To test with {working_providers[0]}:")
    print(f"   python test_ai.py")
    print(f"   # Then choose option 1-4 to test")
else:
    print("\n⚠️ No working providers found!")
    print("\n💡 Next steps:")
    print("   1. Get a valid DeepSeek API key from https://platform.deepseek.com/")
    print("   2. OR start your local AI server on port 1234")
    print("   3. OR add a valid OpenAI API key")
    print("   4. Update ai_test/config.json with the new credentials")
    print("   5. Run this script again: python check_status.py")

print("="*80 + "\n")

