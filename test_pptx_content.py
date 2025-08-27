#!/usr/bin/env python3
"""
Test script to check what content is being placed in PowerPoint
"""

import os
from common.pptx_export import PowerPointGenerator

def test_pptx_content():
    """Test what content is being parsed and placed in PowerPoint"""
    
    print("🔍 Testing PowerPoint Content Generation")
    print("=" * 60)
    
    # Read the content file
    markdown_path = "fdd_utils/bs_content.md"
    
    if not os.path.exists(markdown_path):
        print(f"❌ Content file not found: {markdown_path}")
        return
    
    print(f"✅ Content file found: {markdown_path}")
    
    # Read the content
    with open(markdown_path, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    print(f"📊 Content file size: {len(md_content)} characters")
    print(f"📊 Content file lines: {len(md_content.split(chr(10)))}")
    
    # Create PowerPoint generator
    template_path = "fdd_utils/template.pptx"
    generator = PowerPointGenerator(template_path)
    
    # Parse the markdown content
    print(f"\n📝 Parsing markdown content...")
    items = generator.parse_markdown(md_content)
    
    print(f"📊 Parsed {len(items)} financial items:")
    for i, item in enumerate(items):
        print(f"   {i+1}. {item.accounting_type} - {item.account_title}")
        print(f"      Descriptions: {len(item.descriptions)}")
        for j, desc in enumerate(item.descriptions):
            print(f"         {j+1}: {desc[:100]}{'...' if len(desc) > 100 else ''}")
    
    # Plan content distribution
    print(f"\n📋 Planning content distribution...")
    distribution = generator._plan_content_distribution(items)
    
    print(f"📊 Distribution plan:")
    for slide_idx, section, section_items in distribution:
        print(f"   Slide {slide_idx + 1}, Section '{section}': {len(section_items)} items")
        for item in section_items:
            print(f"      - {item.accounting_type}: {item.account_title}")
    
    print("\n" + "=" * 60)
    print("🔍 Content analysis completed")

if __name__ == "__main__":
    test_pptx_content()
