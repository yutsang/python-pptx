#!/usr/bin/env python3
"""
Debug script to test PowerPoint export exactly as Streamlit app does
"""

import os
import sys
from datetime import datetime
from common.pptx_export import export_pptx, merge_presentations

def debug_streamlit_export():
    """Debug PowerPoint export exactly as Streamlit app does"""
    
    print("🔍 Debugging Streamlit PowerPoint Export")
    print("=" * 60)
    
    # Simulate Streamlit app parameters
    project_name = "Haining Wanpu Limited"  # This is what the app uses
    statement_type = "BS"  # Balance Sheet
    
    print(f"📋 Parameters:")
    print(f"   Project Name: {project_name}")
    print(f"   Statement Type: {statement_type}")
    
    # Check for template file in common locations (exactly as Streamlit app does)
    possible_templates = [
        "fdd_utils/template.pptx",
        "template.pptx", 
        "old_ver/template.pptx",
        "common/template.pptx"
    ]
    
    template_path = None
    for template in possible_templates:
        if os.path.exists(template):
            template_path = template
            print(f"✅ Template found: {template_path}")
            break
    
    if not template_path:
        print("❌ No template found in any location")
        return
    
    # Define output path with timestamp (exactly as Streamlit app does)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"{project_name}_{statement_type.upper()}_{timestamp}.pptx"
    output_path = f"fdd_utils/output/{output_filename}"
    
    print(f"📁 Output path: {output_path}")
    
    # Ensure output directory exists
    os.makedirs("fdd_utils/output", exist_ok=True)
    
    # Check content files (exactly as Streamlit app does)
    if statement_type == "BS":
        markdown_path = "fdd_utils/bs_content.md"
    elif statement_type == "IS":
        markdown_path = "fdd_utils/is_content.md"
    else:
        print("❌ Unsupported statement type")
        return
    
    print(f"📄 Content file: {markdown_path}")
    
    if not os.path.exists(markdown_path):
        print(f"❌ Content file not found: {markdown_path}")
        return
    
    print(f"✅ Content file found: {markdown_path}")
    
    # Simulate Excel file path logic (exactly as Streamlit app does)
    excel_file_path = None
    if os.path.exists("databook.xlsx"):
        excel_file_path = "databook.xlsx"
        print(f"📊 Excel file found: {excel_file_path}")
    else:
        print(f"📊 No Excel file found, using None")
    
    # Read content file to check its structure
    try:
        with open(markdown_path, 'r', encoding='utf-8') as f:
            content = f.read()
        print(f"📊 Content file size: {len(content)} characters")
        print(f"📊 Content file lines: {len(content.split(chr(10)))}")
        
        # Show first few lines
        lines = content.split(chr(10))
        print(f"📄 First 5 lines:")
        for i, line in enumerate(lines[:5]):
            print(f"   {i+1}: {line[:100]}{'...' if len(line) > 100 else ''}")
            
    except Exception as e:
        print(f"❌ Error reading content file: {e}")
        return
    
    # Test export (exactly as Streamlit app does)
    try:
        print(f"\n🔄 Testing export...")
        print(f"   Template: {template_path}")
        print(f"   Markdown: {markdown_path}")
        print(f"   Output: {output_path}")
        print(f"   Project: {project_name}")
        print(f"   Excel: {excel_file_path}")
        
        export_pptx(
            template_path=template_path,
            markdown_path=markdown_path,
            output_path=output_path,
            project_name=project_name,
            excel_file_path=excel_file_path
        )
        
        if os.path.exists(output_path):
            print(f"✅ Export successful: {output_path}")
            print(f"📊 File size: {os.path.getsize(output_path)} bytes")
        else:
            print(f"❌ Export failed: file not created")
            
    except Exception as e:
        print(f"❌ Export error: {str(e)}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 60)
    print("🔍 Debug completed")

if __name__ == "__main__":
    debug_streamlit_export()
