#!/usr/bin/env python3
"""
Debug the _populate_section method
"""

import os
from common.pptx_export import PowerPointGenerator

def debug_populate_section():
    """Debug the _populate_section method"""
    
    print("🔍 Debugging _populate_section method")
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
    
    # Create PowerPoint generator
    template_path = "fdd_utils/template.pptx"
    generator = PowerPointGenerator(template_path)
    
    # Parse the markdown content
    print(f"\n📝 Parsing markdown content...")
    items = generator.parse_markdown(md_content)
    
    print(f"📊 Parsed {len(items)} financial items")
    
    # Take the first few items for testing
    test_items = items[:3]
    print(f"📋 Testing with {len(test_items)} items:")
    for i, item in enumerate(test_items):
        print(f"   {i+1}. {item.accounting_type} - {item.account_title}")
        print(f"      Descriptions: {len(item.descriptions)}")
        for j, desc in enumerate(item.descriptions):
            print(f"         {j+1}: {desc[:50]}{'...' if len(desc) > 50 else ''}")
    
    # Get a shape to test with
    slide = generator.prs.slides[1]  # Slide 2 has textMainBullets
    shape = None
    for s in slide.shapes:
        if s.name == "textMainBullets":
            shape = s
            break
    
    if not shape:
        print("❌ Could not find textMainBullets shape")
        return
    
    print(f"\n📄 Testing with shape: {shape.name}")
    print(f"   Shape type: {type(shape).__name__}")
    
    # Test _populate_section
    print(f"\n🔄 Testing _populate_section...")
    try:
        generator._populate_section(shape, test_items)
        print(f"✅ _populate_section completed successfully")
        
        # Check the content
        text_frame = shape.text_frame
        print(f"📊 Text frame has {len(text_frame.paragraphs)} paragraphs")
        
        for i, para in enumerate(text_frame.paragraphs):
            para_text = para.text.strip()
            if para_text:
                print(f"   Para {i+1}: {para_text[:100]}{'...' if len(para_text) > 100 else ''}")
            else:
                print(f"   Para {i+1}: (empty)")
        
    except Exception as e:
        print(f"❌ _populate_section failed: {str(e)}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 60)
    print("🔍 Debug completed")

if __name__ == "__main__":
    debug_populate_section()
