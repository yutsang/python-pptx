#!/usr/bin/env python3
"""
Test script to debug PowerPoint generation with Chinese content
"""

import os
import sys
sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')

def test_pptx_generation():
    """Test PowerPoint generation with debug logging"""
    try:
        # Import required modules
        from common.pptx_export import export_pptx

        # Test data
        template_path = "fdd_utils/template.pptx"
        markdown_path = "fdd_utils/bs_content.md"
        output_path = "fdd_utils/output/debug_test.pptx"
        project_name = "TestProject"
        language = "chinese"  # Test with Chinese

        print("🧪 TESTING PPTX GENERATION WITH DEBUG LOGGING")
        print("=" * 80)
        print(f"Template: {template_path}")
        print(f"Markdown: {markdown_path}")
        print(f"Output: {output_path}")
        print(f"Language: {language}")
        print("=" * 80)

        # Check if files exist
        if not os.path.exists(template_path):
            print(f"❌ Template not found: {template_path}")
            return

        if not os.path.exists(markdown_path):
            print(f"❌ Markdown not found: {markdown_path}")
            return

        # Generate PowerPoint
        print("🚀 Starting PowerPoint generation...")
        result = export_pptx(
            template_path=template_path,
            markdown_path=markdown_path,
            output_path=output_path,
            project_name=project_name,
            language=language
        )

        if result and os.path.exists(result):
            print(f"✅ PowerPoint generated successfully: {result}")
            print(f"   File size: {os.path.getsize(result)} bytes")
        else:
            print("❌ PowerPoint generation failed")

    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_pptx_generation()
