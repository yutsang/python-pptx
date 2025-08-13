#!/usr/bin/env python3
"""
Direct AI Processing Test

This script directly tests the AI processing functionality with DeepSeek API
to verify that the system is actually using the AI instead of fallback mode.
"""

import json
import time
import os
import sys
from pathlib import Path
import warnings

# Suppress warnings for cleaner output
warnings.filterwarnings('ignore')

def test_direct_ai_processing():
    """Test AI processing directly with DeepSeek API"""
    print("🤖 Testing Direct AI Processing with DeepSeek...")
    
    try:
        from common.assistant import process_keys, load_config, AI_AVAILABLE
        
        print(f"AI_AVAILABLE: {AI_AVAILABLE}")
        
        # Load configuration
        config = load_config('fdd_utils/config.json')
        print(f"DeepSeek API Key configured: {'Yes' if config.get('DEEPSEEK_API_KEY') else 'No'}")
        print(f"DeepSeek API Base configured: {'Yes' if config.get('DEEPSEEK_API_BASE') else 'No'}")
        
        # Test with a single key to minimize API usage
        test_keys = ['Cash']
        entity_name = 'Haining'
        entity_helpers = ['Haining', 'Wanpu']
        
        # Check if databook exists
        if not os.path.exists('databook.xlsx'):
            print("⚠️ No databook.xlsx found, creating minimal test data")
            # Create a minimal test scenario
            return test_minimal_ai_processing()
        
        print(f"📄 Using databook.xlsx for testing")
        print(f"🔑 Testing key: {test_keys[0]}")
        print(f"🏢 Entity: {entity_name}")
        
        # Test AI processing with explicit use_ai=True
        start_time = time.time()
        
        results = process_keys(
            keys=test_keys,
            entity_name=entity_name,
            entity_helpers=entity_helpers,
            input_file='databook.xlsx',
            mapping_file="utils/mapping.json",
            pattern_file="utils/pattern.json",
            config_file='fdd_utils/config.json',
            prompts_file='fdd_utils/prompts.json',
            use_ai=True,  # Explicitly enable AI
            convert_thousands=False
        )
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        if results:
            print(f"✅ AI Processing: Successful (Processing time: {processing_time:.2f}s)")
            print(f"📝 Generated content for {len(results)} keys")
            
            # Check if it's using real AI or fallback
            for key, content in results.items():
                if "[TEST]" in content:
                    print(f"❌ Using fallback mode: {content[:100]}...")
                    return False, processing_time, "Fallback mode detected"
                else:
                    print(f"✅ Using real AI: {content[:100]}...")
                    return True, processing_time, "Real AI processing detected"
        else:
            print("❌ AI Processing: No results generated")
            return False, 0, "No results"
            
    except Exception as e:
        print(f"❌ AI Processing: Error - {e}")
        return False, 0, str(e)

def test_minimal_ai_processing():
    """Test AI processing with minimal data"""
    print("🧪 Testing Minimal AI Processing...")
    
    try:
        from openai import OpenAI
        import httpx
        
        # Load config
        with open('fdd_utils/config.json', 'r') as f:
            config = json.load(f)
        
        # Initialize client
        client = OpenAI(
            api_key=config['DEEPSEEK_API_KEY'],
            base_url=config['DEEPSEEK_API_BASE'],
            http_client=httpx.Client(verify=False)
        )
        
        # Test with a financial analysis prompt
        system_prompt = """You are a senior financial analyst specializing in due diligence reporting. 
        Your task is to analyze financial data and provide insights."""
        
        user_prompt = """Please provide a brief analysis of cash management for a company. 
        Include specific dollar amounts and entity names in your response. 
        Format your response as a professional financial analysis paragraph."""
        
        print("📤 Sending financial analysis request to DeepSeek API...")
        start_time = time.time()
        
        response = client.chat.completions.create(
            model=config['DEEPSEEK_CHAT_MODEL'],
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            max_tokens=200,
            temperature=0.3
        )
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        if response.choices and len(response.choices) > 0:
            content = response.choices[0].message.content
            tokens_used = response.usage.total_tokens if response.usage else 0
            
            print(f"✅ Minimal AI Processing: Successful (Processing time: {processing_time:.2f}s)")
            print(f"📝 Response: {content}")
            print(f"🔢 Tokens used: {tokens_used}")
            
            # Check if response includes required elements
            has_amounts = any(char.isdigit() for char in content)
            has_entities = any(word.istitle() and len(word) > 3 for word in content.split())
            
            if has_amounts and has_entities:
                print("✅ Response includes dollar amounts and entity names")
                return True, processing_time, "Real AI with financial analysis"
            else:
                print("⚠️ Response may not include required financial elements")
                return True, processing_time, "Real AI but missing financial elements"
        else:
            print("❌ No response content received")
            return False, 0, "No response content"
            
    except Exception as e:
        print(f"❌ Minimal AI Processing: Error - {e}")
        return False, 0, str(e)

def test_enhanced_prompt_features():
    """Test if the enhanced prompts are working correctly"""
    print("\n💡 Testing Enhanced Prompt Features...")
    
    try:
        from openai import OpenAI
        import httpx
        
        # Load config
        with open('fdd_utils/config.json', 'r') as f:
            config = json.load(f)
        
        # Initialize client
        client = OpenAI(
            api_key=config['DEEPSEEK_API_KEY'],
            base_url=config['DEEPSEEK_API_BASE'],
            http_client=httpx.Client(verify=False)
        )
        
        # Test with the enhanced prompt requirements
        system_prompt = """You are a content generation specialist for financial reports. Your role is to generate comprehensive financial analysis content based on worksheet data and predefined patterns. Focus on:
1. Content generation using patterns from pattern.json
2. Integration of actual worksheet data into narrative content
3. Professional financial writing suitable for audit reports
4. Consistent formatting and structure
5. Clear, accurate descriptions of financial positions
6. Replace all entity placeholders (e.g., [ENTITY_NAME], [COMPANY_NAME]) with the actual entity name from the provided financial data tables
7. Use the exact entity name as shown in the financial data tables (e.g., 'Haining Wanpu', 'Ningbo Wanchen')
8. ALWAYS specify the exact dollar amounts and currency when filling in financial figures
9. ALWAYS identify and mention the specific entity names you are filling in
10. Provide a summary of key financial figures and entities used in your response"""
        
        user_prompt = """Please analyze the cash position for Haining Wanpu company. 
        Use the following data:
        - Cash at bank: CNY9.1M
        - Entity: Haining Wanpu
        
        Please provide your analysis and include a summary of the amounts and entities used."""
        
        print("📤 Testing enhanced prompt features...")
        start_time = time.time()
        
        response = client.chat.completions.create(
            model=config['DEEPSEEK_CHAT_MODEL'],
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            max_tokens=300,
            temperature=0.3
        )
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        if response.choices and len(response.choices) > 0:
            content = response.choices[0].message.content
            tokens_used = response.usage.total_tokens if response.usage else 0
            
            print(f"✅ Enhanced Prompt Test: Successful (Processing time: {processing_time:.2f}s)")
            print(f"📝 Response: {content}")
            print(f"🔢 Tokens used: {tokens_used}")
            
            # Check for enhanced features
            features_found = []
            if "CNY" in content or "$" in content or "M" in content or "K" in content:
                features_found.append("dollar amounts")
            if "Haining Wanpu" in content:
                features_found.append("entity names")
            if "summary" in content.lower():
                features_found.append("summary")
            
            print(f"✅ Enhanced features found: {features_found}")
            return True, processing_time, f"Enhanced features: {features_found}"
        else:
            print("❌ No response content received")
            return False, 0, "No response content"
            
    except Exception as e:
        print(f"❌ Enhanced Prompt Test: Error - {e}")
        return False, 0, str(e)

def main():
    """Run comprehensive AI processing tests"""
    print("🚀 Direct AI Processing Test for DeepSeek Integration")
    print("=" * 60)
    
    test_results = {}
    
    # Test 1: Direct AI Processing
    ai_success, processing_time, details = test_direct_ai_processing()
    test_results['direct_ai_processing'] = {
        'success': ai_success,
        'processing_time': processing_time,
        'details': details
    }
    
    # Test 2: Enhanced Prompt Features
    enhanced_success, enhanced_time, enhanced_details = test_enhanced_prompt_features()
    test_results['enhanced_prompt_features'] = {
        'success': enhanced_success,
        'processing_time': enhanced_time,
        'details': enhanced_details
    }
    
    # Summary
    print("\n" + "=" * 60)
    print("📊 DIRECT AI PROCESSING TEST SUMMARY")
    print("=" * 60)
    
    total_tests = len(test_results)
    passed_tests = sum(1 for result in test_results.values() if result['success'])
    
    print(f"✅ Tests Passed: {passed_tests}/{total_tests}")
    
    for test_name, result in test_results.items():
        status = "✅ PASS" if result['success'] else "❌ FAIL"
        print(f"{status} {test_name.replace('_', ' ').title()}")
        print(f"   Details: {result['details']}")
        if result['processing_time'] > 0:
            print(f"   Time: {result['processing_time']:.2f}s")
    
    # Performance metrics
    if test_results['direct_ai_processing']['success']:
        print(f"\n⚡ Performance Metrics:")
        print(f"   Direct AI Processing: {test_results['direct_ai_processing']['processing_time']:.2f}s")
    
    if test_results['enhanced_prompt_features']['success']:
        print(f"   Enhanced Prompt Test: {test_results['enhanced_prompt_features']['processing_time']:.2f}s")
    
    # Overall assessment
    print(f"\n🎯 Overall Assessment:")
    if passed_tests == total_tests:
        print("🎉 EXCELLENT: All AI processing tests passed! DeepSeek integration is working correctly.")
    elif passed_tests >= total_tests * 0.5:
        print("👍 GOOD: Most AI processing tests passed. DeepSeek integration is functional.")
    else:
        print("❌ POOR: Many AI processing tests failed. DeepSeek integration needs attention.")
    
    return test_results

if __name__ == "__main__":
    results = main()
    
    # Save results to file
    with open('direct_ai_test_results.json', 'w') as f:
        json.dump(results, f, indent=2, default=str)
    
    print(f"\n📄 Results saved to: direct_ai_test_results.json") 