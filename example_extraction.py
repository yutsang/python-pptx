#!/usr/bin/env python3
"""
Simple example showing how to use extract_data_from_excel correctly
"""

from fdd_utils.process_databook import extract_data_from_excel
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)


def example_basic_usage():
    """Example 1: Basic usage"""
    print("=" * 60)
    print("Example 1: Basic Usage")
    print("=" * 60)
    
    # IMPORTANT: Change these values to match your file
    databook_path = "databook.xlsx"  # ← Your Excel file path
    entity_name = ""                  # ← Entity name (or "" for single entity)
    mode = "All"                      # ← "All", "BS", or "IS"
    
    print(f"\nExtracting from: {databook_path}")
    print(f"Entity: {entity_name if entity_name else '(single entity)'}")
    print(f"Mode: {mode}\n")
    
    # Extract
    dfs, workbook_list, result_type, language = extract_data_from_excel(
        databook_path=databook_path,
        entity_name=entity_name,
        mode=mode
    )
    
    # Check results
    if dfs and len(dfs) > 0:
        print(f"✅ SUCCESS! Extracted {len(dfs)} sheets")
        print(f"   Language detected: {language}")
        print(f"   Sheets extracted: {workbook_list}")
        
        # Show first sheet data
        first_key = list(dfs.keys())[0]
        print(f"\n📊 Sample data from '{first_key}':")
        print(dfs[first_key].head(5))
        
        return dfs, workbook_list, language
    else:
        print("❌ EXTRACTION FAILED!")
        print("\nPlease run the diagnostic tool:")
        print("   python test_extraction.py")
        print("\nOr read the guide:")
        print("   EXTRACTION_GUIDE.md")
        return None, None, None


def example_chinese_databook():
    """Example 2: Chinese databook"""
    print("\n" + "=" * 60)
    print("Example 2: Chinese Databook")
    print("=" * 60)
    
    dfs, workbook_list, _, language = extract_data_from_excel(
        databook_path="240624.联洋-databook.xlsx",
        entity_name="联洋",
        mode="All"
    )
    
    if dfs:
        print(f"✅ Extracted {len(dfs)} sheets")
        print(f"   Language: {language}")
        
        # Show formatted values
        if '货币资金' in dfs:
            print(f"\n📊 货币资金 (Cash) data:")
            print(dfs['货币资金'])
            print("\nNote: Values are auto-formatted:")
            print("  - 万元 = 1 decimal place (e.g., 7.8万)")
            print("  - 亿元 = 2 decimal places (e.g., 1.23亿)")


def example_english_databook():
    """Example 3: English databook"""
    print("\n" + "=" * 60)
    print("Example 3: English Databook")
    print("=" * 60)
    
    dfs, workbook_list, _, language = extract_data_from_excel(
        databook_path="inputs/221128.Project TK.Databook.JW.xlsx",
        entity_name="Haining Wanpu",
        mode="BS"
    )
    
    if dfs:
        print(f"✅ Extracted {len(dfs)} sheets")
        print(f"   Language: {language}")
        
        # Show formatted values
        if 'Cash' in dfs:
            print(f"\n📊 Cash data:")
            print(dfs['Cash'])
            print("\nNote: Values are auto-formatted:")
            print("  - K = 1 decimal place (e.g., 78.2K)")
            print("  - million = 2 decimal places (e.g., 12.35 million)")


def example_with_ai_pipeline():
    """Example 4: Full pipeline with AI"""
    print("\n" + "=" * 60)
    print("Example 4: Extract + AI Pipeline")
    print("=" * 60)
    
    # Step 1: Extract data
    print("\n[1/3] Extracting data...")
    dfs, workbook_list, _, language = extract_data_from_excel(
        databook_path="databook.xlsx",
        entity_name="",
        mode="All"
    )
    
    if not dfs or len(dfs) == 0:
        print("❌ Extraction failed! Cannot proceed with AI pipeline.")
        return
    
    print(f"✅ Extracted {len(dfs)} sheets")
    
    # Step 2: Run AI pipeline
    print("\n[2/3] Running AI pipeline...")
    from fdd_utils.content_generation import run_ai_pipeline
    
    results = run_ai_pipeline(
        mapping_keys=workbook_list,
        dfs=dfs,
        model_type='local',  # or 'deepseek', 'openai'
        language=language,
        use_multithreading=True
    )
    
    print(f"✅ AI pipeline completed for {len(results)} items")
    
    # Step 3: Get final contents
    print("\n[3/3] Extracting final contents...")
    from fdd_utils.content_generation import extract_final_contents
    
    final_contents = extract_final_contents(results)
    
    print(f"✅ Generated content for {len(final_contents)} accounts")
    
    # Show sample
    if final_contents:
        first_key = list(final_contents.keys())[0]
        print(f"\n📝 Sample content for '{first_key}':")
        print(final_contents[first_key][:200] + "...")
    
    return final_contents


if __name__ == "__main__":
    print("\n" + "=" * 80)
    print("EXTRACT_DATA_FROM_EXCEL - USAGE EXAMPLES")
    print("=" * 80)
    
    # Run basic example
    example_basic_usage()
    
    # Uncomment to run other examples:
    # example_chinese_databook()
    # example_english_databook()
    # example_with_ai_pipeline()
    
    print("\n" + "=" * 80)
    print("💡 Tips:")
    print("=" * 80)
    print("""
1. If extraction returns None or empty:
   → Run: python test_extraction.py (diagnostic tool)
   → Read: EXTRACTION_GUIDE.md (troubleshooting guide)

2. Common issues:
   → File path is wrong
   → Entity name doesn't match
   → Sheet names don't match mappings.yml
   → Missing financial indicators in sheets

3. The function returns formatted values:
   → Chinese: 万元 (1 d.p.), 亿元 (2 d.p.)
   → English: K (1 d.p.), million (2 d.p.)

4. For negative retained earnings:
   → 未分配利润 (negative) → 未弥补亏损 (positive display)
   → Retained Earnings (negative) → Accumulated Losses (positive display)
""")

