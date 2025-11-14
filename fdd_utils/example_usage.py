"""
Example Usage of AI Pipeline
Demonstrates different ways to use the multi-agent AI pipeline
"""

from ai_pipeline import AIPipelineOrchestrator, run_quick_pipeline


def example_1_quick_pipeline():
    """Example 1: Quick pipeline execution with defaults."""
    print("\n" + "="*70)
    print("EXAMPLE 1: Quick Pipeline Execution")
    print("="*70 + "\n")
    
    results = run_quick_pipeline(
        databook_path='../databook.xlsx',
        entity_name='Sample Company',
        mode='All',
        model_type='deepseek',
        language='Eng'
    )
    
    print(f"\nProcessed {len(results)} items successfully!")
    

def example_2_full_pipeline():
    """Example 2: Full pipeline with custom settings."""
    print("\n" + "="*70)
    print("EXAMPLE 2: Full Pipeline with Custom Settings")
    print("="*70 + "\n")
    
    # Initialize orchestrator
    orchestrator = AIPipelineOrchestrator(config_path='config.yml')
    
    # Load data
    data_info = orchestrator.load_data(
        databook_path='../databook.xlsx',
        entity_name='Sample Company',
        mode='All'
    )
    
    print(f"\nLoaded {len(data_info['mapping_keys'])} worksheets")
    print(f"Detected language: {data_info['report_language']}")
    
    # Run full pipeline
    results = orchestrator.run_full_pipeline(
        model_type='deepseek',
        language='Eng',
        use_heuristic=False,
        use_multithreading=True,
        max_workers=4
    )
    
    # Save results
    orchestrator.save_results('output/full_pipeline_results.yml')
    
    # Print summary
    orchestrator.print_results_summary()


def example_3_individual_agents():
    """Example 3: Running individual agents."""
    print("\n" + "="*70)
    print("EXAMPLE 3: Running Individual Agents")
    print("="*70 + "\n")
    
    orchestrator = AIPipelineOrchestrator()
    orchestrator.load_data('../databook.xlsx', 'Sample Company')
    
    # Run Agent 1 only
    print("\nRunning Agent 1 (Content Generator)...")
    results_1 = orchestrator.run_agent_only('agent_1')
    
    # Run Agent 2 with Agent 1 results
    print("\nRunning Agent 2 (Value Checker)...")
    results_2 = orchestrator.run_agent_only('agent_2', previous_results=results_1)
    
    # Run Agent 4 to finalize (skip Agent 3 if desired)
    print("\nRunning Agent 4 (Format Checker)...")
    final_results = orchestrator.run_agent_only('agent_4', previous_results=results_2)
    
    orchestrator.save_results('output/individual_agents_results.yml')


def example_4_sequential_agents():
    """Example 4: Running specific agents in sequence."""
    print("\n" + "="*70)
    print("EXAMPLE 4: Sequential Agents")
    print("="*70 + "\n")
    
    orchestrator = AIPipelineOrchestrator()
    orchestrator.load_data('../databook.xlsx', 'Sample Company')
    
    # Run only Agent 1, 2, and 4 (skip Agent 3)
    results = orchestrator.run_sequential_agents(
        agents=['agent_1', 'agent_2', 'agent_4'],
        model_type='deepseek',
        language='Eng'
    )
    
    orchestrator.save_results('output/sequential_results.yml')


def example_5_heuristic_mode():
    """Example 5: Using heuristic mode (no AI)."""
    print("\n" + "="*70)
    print("EXAMPLE 5: Heuristic Mode (No AI)")
    print("="*70 + "\n")
    
    orchestrator = AIPipelineOrchestrator()
    orchestrator.load_data('../databook.xlsx', 'Sample Company')
    
    # Run with heuristic mode (rule-based, no AI calls)
    results = orchestrator.run_full_pipeline(
        use_heuristic=True,
        use_multithreading=True,
        max_workers=8
    )
    
    print("\nHeuristic mode completed (no AI tokens used)")
    orchestrator.save_results('output/heuristic_results.yml')


def example_6_chinese_language():
    """Example 6: Processing with Chinese language."""
    print("\n" + "="*70)
    print("EXAMPLE 6: Chinese Language Processing")
    print("="*70 + "\n")
    
    orchestrator = AIPipelineOrchestrator()
    
    # Load data (will auto-detect Chinese if present)
    orchestrator.load_data('../databook.xlsx', 'Sample Company')
    
    # Run with Chinese language
    results = orchestrator.run_full_pipeline(
        model_type='deepseek',
        language='Chi',  # Force Chinese
        use_multithreading=True
    )
    
    orchestrator.save_results('output/chinese_results.yml')


def example_7_custom_aihelper():
    """Example 7: Using AIHelper directly."""
    print("\n" + "="*70)
    print("EXAMPLE 7: Custom AIHelper Usage")
    print("="*70 + "\n")
    
    from ai_helper import AIHelper
    import pandas as pd
    
    # Create sample data
    df = pd.DataFrame({
        'Account': ['Cash', 'Bank Deposits'],
        'Amount': [1000000, 2500000]
    })
    
    # Initialize AIHelper
    ai_helper = AIHelper(
        model_type='deepseek',
        agent_name='agent_1',
        language='Eng',
        use_heuristic=False
    )
    
    # Load prompts
    system_prompt, user_prompt_template = ai_helper.load_prompts()
    
    # Format user prompt
    user_prompt = user_prompt_template.format(
        patterns="Pattern 1: Cash comprises {amount} as at {date}.",
        key="Cash",
        financial_figure=df.to_markdown(index=False)
    )
    
    # Get response
    response = ai_helper.get_response(
        user_prompt=user_prompt,
        system_prompt=system_prompt,
        temperature=0.7
    )
    
    print("\nAI Response:")
    print(response['content'])
    print(f"\nTokens used: {response['tokens_used']}")
    print(f"Duration: {response['duration']:.2f}s")


def example_8_performance_comparison():
    """Example 8: Compare performance with/without multi-threading."""
    print("\n" + "="*70)
    print("EXAMPLE 8: Performance Comparison")
    print("="*70 + "\n")
    
    import time
    
    orchestrator = AIPipelineOrchestrator()
    orchestrator.load_data('../databook.xlsx', 'Sample Company')
    
    # Test without multi-threading
    print("\nTest 1: Without Multi-threading")
    start = time.time()
    results_1 = orchestrator.run_full_pipeline(
        use_multithreading=False
    )
    time_single = time.time() - start
    print(f"Time: {time_single:.2f}s")
    
    # Test with multi-threading
    print("\nTest 2: With Multi-threading (4 workers)")
    start = time.time()
    results_2 = orchestrator.run_full_pipeline(
        use_multithreading=True,
        max_workers=4
    )
    time_multi = time.time() - start
    print(f"Time: {time_multi:.2f}s")
    
    # Compare
    speedup = time_single / time_multi
    print(f"\nSpeedup: {speedup:.2f}x faster with multi-threading")


if __name__ == '__main__':
    # Run examples (comment out the ones you don't want to run)
    
    print("\n" + "="*70)
    print("AI PIPELINE EXAMPLES")
    print("="*70)
    
    # Uncomment the examples you want to run:
    
    # example_1_quick_pipeline()
    # example_2_full_pipeline()
    # example_3_individual_agents()
    # example_4_sequential_agents()
    # example_5_heuristic_mode()
    # example_6_chinese_language()
    # example_7_custom_aihelper()
    # example_8_performance_comparison()
    
    print("\n" + "="*70)
    print("EXAMPLES COMPLETED")
    print("="*70 + "\n")
    
    print("To run an example, uncomment it in the main section of this file.")

