#!/usr/bin/env python3
"""
Quick Example - How to Use the AI Module
Run this for a quick demonstration
"""

from ai_module import AIModule

# Initialize AI Module (automatically loads config.json)
ai = AIModule()

# Example 1: Simple Content Generation
print("="*80)
print("EXAMPLE 1: Simple Content Generation")
print("="*80)

result = ai.generate_content(
    system_prompt="You are a financial analyst.",
    user_prompt="Explain what cash and cash equivalents are in 50 words.",
    provider='deepseek',
    temperature=0.7
)

print(f"\n✅ Generated Content:")
print(result['content'])
print(f"\n📊 Tokens Used: {result['tokens']['total_tokens']}")

# Example 2: Test with Financial Data
print("\n\n" + "="*80)
print("EXAMPLE 2: Financial Data Analysis")
print("="*80)

sample_data = """
| Category | Subcategory | Value |
|----------|-------------|-------|
| Cash | Bank of China | 12,567,955.7 |
| Cash | ICBC | 3,057,807.5 |
| Cash | CMB | 163,535.6 |
| Total | | 15,789,298.8 |
"""

result2 = ai.generate_content(
    system_prompt="You are a financial report writer. Analyze the data and write a brief commentary.",
    user_prompt=f"Analyze this cash data and write a 2-sentence summary:\n\n{sample_data}",
    provider='deepseek',
    temperature=0.7
)

print(f"\n✅ Analysis:")
print(result2['content'])

# Example 3: Multi-Agent Test
print("\n\n" + "="*80)
print("EXAMPLE 3: Multi-Agent Workflow")
print("="*80)

results = ai.test_multi_agent(
    financial_data=sample_data,
    key="Cash",
    entity_name="Example Company",
    mode='english',
    provider='deepseek'
)

print("\n✅ Multi-Agent Test Complete!")
print(f"Agent 1 Generated: {len(results['agent1']['content'])} characters")
if results.get('agent2'):
    print(f"Agent 2 Reviewed: {len(results['agent2']['content'])} characters")

# List available providers
print("\n\n" + "="*80)
print("AVAILABLE AI PROVIDERS")
print("="*80)
providers = ai.list_available_providers()

