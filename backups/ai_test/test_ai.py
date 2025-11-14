#!/usr/bin/env python3
"""
AI Testing Script
Run various tests on the AI module with different configurations
"""

import sys
import json
from pathlib import Path
from ai_module import AIModule

# Sample financial data for testing
SAMPLE_CASH_DATA = """
**Cash - Dongguan Lianyang**

*Sheet: Cash and Cash Equivalents*
*Date: 31-Dec-2021*
*Currency: CNY'000*

| Category | Subcategory | Value |
|----------|-------------|-------|
| 活期存款_专用账户_资本金户 | 东莞农商行 | 12,567,955.7 |
| 银行存款_活期存款基本户 | 工行东莞麻涌支行 | 0.01 |
| 活期存款_一般账户 | 东莞农商行常平支行 | 3,057,807.5 |
| 活期存款_一般账户 | 招行车公庙支行 | 163,535.6 |
| 活期存款_专用账户_贷款户 | 东莞农商行 | 29,427.3 |
| 合计 | | 15,818,726.1 |
"""

SAMPLE_AR_DATA = """
**Accounts Receivable - Dongguan Lianyang**

*Sheet: Accounts Receivable*
*Date: 31-Dec-2021*
*Currency: CNY'000*

| Category | Subcategory | Value |
|----------|-------------|-------|
| 应收账款 | 客户A | 5,230,145.2 |
| 应收账款 | 客户B | 2,145,678.3 |
| 应收账款 | 客户C | 1,234,567.8 |
| 坏账准备 | | -125,456.7 |
| 合计 | | 8,484,934.6 |
"""


def run_single_test(ai_module, test_name, financial_data, key, entity_name, mode='english', provider='deepseek'):
    """Run a single test case"""
    print(f"\n\n{'#'*100}")
    print(f"# TEST: {test_name}")
    print(f"{'#'*100}\n")
    
    results = ai_module.test_multi_agent(
        financial_data=financial_data,
        key=key,
        entity_name=entity_name,
        mode=mode,
        provider=provider
    )
    
    return results


def run_all_tests():
    """Run all test cases"""
    print("="*100)
    print(" AI MODULE - COMPREHENSIVE TEST SUITE")
    print("="*100)
    
    # Initialize AI Module
    ai = AIModule()
    ai.list_available_providers()
    
    all_results = {}
    
    # Test 1: Cash Analysis - English Mode
    all_results['cash_english'] = run_single_test(
        ai_module=ai,
        test_name="Cash Analysis (English Mode)",
        financial_data=SAMPLE_CASH_DATA,
        key="Cash",
        entity_name="Dongguan Lianyang",
        mode='english',
        provider='deepseek'
    )
    
    # Test 2: Cash Analysis - Chinese Mode
    all_results['cash_chinese'] = run_single_test(
        ai_module=ai,
        test_name="Cash Analysis (Chinese Mode)",
        financial_data=SAMPLE_CASH_DATA,
        key="Cash",
        entity_name="东莞联洋",
        mode='chinese',
        provider='deepseek'
    )
    
    # Test 3: AR Analysis - English Mode
    all_results['ar_english'] = run_single_test(
        ai_module=ai,
        test_name="Accounts Receivable Analysis (English Mode)",
        financial_data=SAMPLE_AR_DATA,
        key="AR",
        entity_name="Dongguan Lianyang",
        mode='english',
        provider='deepseek'
    )
    
    # Test 4: AR Analysis - Chinese Mode
    all_results['ar_chinese'] = run_single_test(
        ai_module=ai,
        test_name="Accounts Receivable Analysis (Chinese Mode)",
        financial_data=SAMPLE_AR_DATA,
        key="AR",
        entity_name="东莞联洋",
        mode='chinese',
        provider='deepseek'
    )
    
    # Generate summary report
    print(f"\n\n{'='*100}")
    print(" FINAL TEST SUMMARY")
    print(f"{'='*100}\n")
    
    total_tests = len(all_results)
    passed_tests = sum(1 for r in all_results.values() if r.get('agent1', {}).get('content'))
    
    print(f"Total Tests Run: {total_tests}")
    print(f"Tests Passed: {passed_tests}")
    print(f"Tests Failed: {total_tests - passed_tests}")
    print(f"Success Rate: {(passed_tests/total_tests*100):.1f}%\n")
    
    # Detailed results
    print("Detailed Results:")
    print("-" * 100)
    for test_name, results in all_results.items():
        agent1_status = "✅" if results.get('agent1', {}).get('content') else "❌"
        agent2_status = "✅" if results.get('agent2', {}).get('content') else "❌"
        
        agent1_tokens = results.get('agent1', {}).get('tokens', {}).get('total_tokens', 0)
        agent2_tokens = results.get('agent2', {}).get('tokens', {}).get('total_tokens', 0)
        total_tokens = agent1_tokens + agent2_tokens
        
        print(f"{test_name:30s} | Agent1: {agent1_status} | Agent2: {agent2_status} | Tokens: {total_tokens:,}")
    
    print("-" * 100)
    print(f"\n{'='*100}\n")
    
    return all_results


def interactive_test():
    """Interactive testing mode"""
    print("="*100)
    print(" AI MODULE - INTERACTIVE TEST MODE")
    print("="*100)
    
    ai = AIModule()
    ai.list_available_providers()
    
    while True:
        print("\n\n" + "="*80)
        print("🧪 Interactive Test Options:")
        print("="*80)
        print("\n📋 Pre-configured Tests:")
        print("  1. Test Cash Analysis (English)")
        print("  2. Test Cash Analysis (Chinese)")
        print("  3. Test AR Analysis (English)")
        print("  4. Test AR Analysis (Chinese)")
        print("\n🤖 Individual Agent Tests:")
        print("  5. Test Agent 1 Only (Content Generation)")
        print("  6. Test Agent 2 Only (Proofreader)")
        print("  7. Test Agent 3 Only (Pattern Compliance)")
        print("\n🔧 Advanced Options:")
        print("  8. Custom Test (enter your own data)")
        print("  9. Multi-Agent with Specific Agents (choose which to run)")
        print("  A. Run All Automated Tests")
        print("\n  0. Exit")
        print("="*80)
        
        choice = input("\nEnter your choice (0-9, A): ").strip().upper()
        
        if choice == '0':
            print("\n👋 Exiting...")
            break
        elif choice == '1':
            run_single_test(ai, "Cash English", SAMPLE_CASH_DATA, "Cash", "Dongguan Lianyang", 'english')
        elif choice == '2':
            run_single_test(ai, "Cash Chinese", SAMPLE_CASH_DATA, "Cash", "东莞联洋", 'chinese')
        elif choice == '3':
            run_single_test(ai, "AR English", SAMPLE_AR_DATA, "AR", "Dongguan Lianyang", 'english')
        elif choice == '4':
            run_single_test(ai, "AR Chinese", SAMPLE_AR_DATA, "AR", "东莞联洋", 'chinese')
        elif choice == '5':
            # Test Agent 1 Only
            print("\n🤖 Testing Agent 1 Only (Content Generation)")
            print("Using Cash data as example...")
            result = ai.test_agent1(
                financial_data=SAMPLE_CASH_DATA,
                key="Cash",
                entity_name="Dongguan Lianyang",
                mode='english',
                provider='deepseek'
            )
            input("\nPress Enter to continue...")
        elif choice == '6':
            # Test Agent 2 Only
            print("\n🔍 Testing Agent 2 Only (Proofreader)")
            print("First, generate content with Agent 1...")
            agent1_result = ai.test_agent1(
                financial_data=SAMPLE_CASH_DATA,
                key="Cash",
                entity_name="Dongguan Lianyang",
                mode='english',
                provider='deepseek'
            )
            if agent1_result.get('content'):
                print("\nNow testing Agent 2...")
                result = ai.test_agent2(
                    financial_data=SAMPLE_CASH_DATA,
                    agent1_content=agent1_result['content'],
                    mode='english',
                    provider='deepseek'
                )
            input("\nPress Enter to continue...")
        elif choice == '7':
            # Test Agent 3 Only
            print("\n✨ Testing Agent 3 Only (Pattern Compliance)")
            print("First, generate content with Agent 1...")
            agent1_result = ai.test_agent1(
                financial_data=SAMPLE_CASH_DATA,
                key="Cash",
                entity_name="Dongguan Lianyang",
                mode='english',
                provider='deepseek'
            )
            if agent1_result.get('content'):
                print("\nNow testing Agent 3...")
                result = ai.test_agent3(
                    agent1_content=agent1_result['content'],
                    mode='english',
                    provider='deepseek'
                )
            input("\nPress Enter to continue...")
        elif choice == '8':
            print("\n📝 Custom Test")
            key = input("Enter financial key (e.g., Cash, AR): ").strip()
            entity = input("Enter entity name: ").strip()
            mode = input("Enter mode (english/chinese): ").strip()
            print("Enter financial data (end with empty line):")
            lines = []
            while True:
                line = input()
                if not line:
                    break
                lines.append(line)
            data = "\n".join(lines)
            run_single_test(ai, "Custom Test", data, key, entity, mode)
        elif choice == '9':
            # Multi-Agent with specific agents
            print("\n🔧 Multi-Agent with Specific Agents")
            print("\nAvailable combinations:")
            print("  all   - Run all agents (1 → 2 → 3)")
            print("  1     - Run Agent 1 only")
            print("  1+2   - Run Agent 1 → Agent 2")
            print("  1+3   - Run Agent 1 → Agent 3")
            agents = input("\nEnter agents to run: ").strip()
            
            ai.test_multi_agent(
                financial_data=SAMPLE_CASH_DATA,
                key="Cash",
                entity_name="Dongguan Lianyang",
                mode='english',
                provider='deepseek',
                agents=agents
            )
            input("\nPress Enter to continue...")
        elif choice == 'A':
            run_all_tests()
        else:
            print("❌ Invalid choice. Please try again.")


if __name__ == "__main__":
    print("""
    ╔═══════════════════════════════════════════════════════════════════╗
    ║                   AI MODULE TEST SUITE                            ║
    ║                                                                   ║
    ║  Tests the AI generation module with various configurations      ║
    ║  and financial data samples                                      ║
    ╚═══════════════════════════════════════════════════════════════════╝
    """)
    
    # Check command line arguments
    if len(sys.argv) > 1:
        if sys.argv[1] == '--all':
            run_all_tests()
        elif sys.argv[1] == '--interactive' or sys.argv[1] == '-i':
            interactive_test()
        else:
            print(f"Usage: {sys.argv[0]} [--all | --interactive | -i]")
            print("  --all         : Run all automated tests")
            print("  --interactive : Run in interactive mode")
            print("  (no args)     : Run interactive mode by default")
    else:
        # Default to interactive mode
        interactive_test()

