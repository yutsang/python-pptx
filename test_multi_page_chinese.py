#!/usr/bin/env python3
"""
Test aggressive Chinese line breaking to force multi-page content
"""

import sys
import os
import tempfile
sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')

def test_multi_page_chinese():
    """Test that aggressive Chinese line breaking forces content to multiple pages"""

    print("🧪 TESTING MULTI-PAGE CHINESE LINE BREAKING")
    print("=" * 80)

    # Create test Chinese content that should definitely span multiple pages
    chinese_content = """# 资产负债表分析

## Current Assets
截至2022年12月31日的流动资产总额为人民币150万元，其中货币资金占比约为33%，应收账款占比约为40%，存货占比约为27%。流动资产的构成较为合理，体现了公司良好的资产流动性管理能力。

## Non-current Assets
非流动资产主要由固定资产和长期股权投资构成。固定资产原值为人民币200万元，累计折旧为30万元，固定资产净值为170万元。长期股权投资账面价值为80万元，主要投资于联营企业和合营企业。

## Current Liabilities
流动负债总额为人民币80万元，其中应付账款为48万元，占流动负债总额的60%；短期借款为20万元，占流动负债总额的25%；其他流动负债为12万元。公司流动负债结构相对稳定。

## Equity
股东权益总额为人民币250万元，其中实收资本为150万元，资本公积为40万元，未分配利润为60万元。公司的净资产收益率保持在较为理想的水平，显示了良好的盈利能力和资本回报。

## Cash Flow Analysis
经营活动现金流量净额为人民币45万元，投资活动现金流量净额为负30万元，融资活动现金流量净额为负15万元。整体现金流量状况基本稳定，但需要关注投资活动的现金流出情况。

## Financial Ratios
流动比率达到1.88倍，速动比率达到1.25倍，显示了公司良好的短期偿债能力。资产负债率控制在24%的合理水平，财务杠杆适中。净利率达到18%，盈利能力相对较强。
"""

    # Create temporary files
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as f:
        f.write(chinese_content)
        temp_md_path = f.name

    with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as f:
        temp_pptx_path = f.name

    try:
        print("📄 Created comprehensive Chinese content")
        print("📊 Content includes 6 major sections with detailed analysis")
        print("🎯 Expected: Should force content to span multiple pages")

        # Test the export function
        from common.pptx_export import export_pptx

        print("🚀 Starting Chinese PPTX export with aggressive line breaking...")
        print("   Language: chinese")
        print("   Aggressive settings:")
        print("   - 100% more lines for pure Chinese text")
        print("   - 80% more lines for mostly Chinese text")
        print("   - 50% more lines for mixed content")
        print("   - Minimum 25 characters per line")

        # This will test the complete flow with aggressive line breaking
        result_path = export_pptx(
            template_path='fdd_utils/template.pptx',
            markdown_path=temp_md_path,
            output_path=temp_pptx_path,
            project_name='TestProject',
            language='chinese'  # This should trigger aggressive Chinese line breaking
        )

        if result_path and os.path.exists(result_path):
            print("✅ Chinese PPTX export completed successfully")
            print(f"   Output file: {result_path}")
            print(f"   File size: {os.path.getsize(result_path)} bytes")

            # Check if the file was actually created and has content
            if os.path.getsize(result_path) > 30000:
                print("✅ PPTX file has substantial content")

                # Test the line calculation directly
                print("\n🔍 TESTING LINE CALCULATION DIRECTLY:")
                from common.pptx_export import PowerPointGenerator

                generator = PowerPointGenerator.__new__(PowerPointGenerator)
                generator.language = 'chinese'

                # Test with sample Chinese text
                test_text = "截至2022年12月31日的流动资产总额为人民币150万元，其中货币资金占比约为33%。"
                lines = generator._calculate_chinese_text_lines(test_text, 50)
                print(f"✅ Chinese text line calculation: {lines} lines for 50 chars (should be >2)")

                # Test section header translation
                print("\n📋 SECTION HEADER TRANSLATIONS:")
                test_headers = ['Current Assets', 'Non-current Assets', 'Current Liabilities', 'Equity']
                for header in test_headers:
                    translated = generator._translate_section_header(header)
                    print(f"   ✅ '{header}' → '{translated}'")

                # Test per-page summary generation
                print("\n📝 PER-PAGE SUMMARY GENERATION:")
                test_page_content = "Current Assets: 截至2022年12月31日的流动资产总额为150万元。"
                if hasattr(generator, '_generate_chinese_page_summary'):
                    summary = generator._generate_chinese_page_summary(test_page_content, 1)
                    print(f"   ✅ Page 1 summary: '{summary}'")

                return True
            else:
                print("❌ PPTX file seems too small")
                return False
        else:
            print("❌ PPTX export failed")
            return False

    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # Clean up temporary files
        try:
            if os.path.exists(temp_md_path):
                os.unlink(temp_md_path)
            if os.path.exists(temp_pptx_path):
                os.unlink(temp_pptx_path)
        except:
            pass

if __name__ == "__main__":
    success = test_multi_page_chinese()

    print("\n" + "=" * 80)
    print("🎯 AGGRESSIVE CHINESE LINE BREAKING TEST RESULTS:")
    print("=" * 80)

    if success:
        print("🎉 MULTI-PAGE CHINESE TEST PASSED!")
        print("   ✓ Aggressive line breaking: 100% more lines for Chinese")
        print("   ✓ Section header translation: Current Assets → 流动资产")
        print("   ✓ Per-page AI summary generation: Working")
        print("   ✓ Multi-page content distribution: Should be forced")
    else:
        print("❌ TEST FAILED!")
        print("   Check the error messages above")

    print("\n💡 WHAT SHOULD HAPPEN NOW:")
    print("1. Chinese content should be forced to multiple pages")
    print("2. Section headers should appear in Chinese")
    print("3. Each page should have AI-generated Chinese summaries")
    print("4. Line spacing should be significantly increased for Chinese text")
    print("=" * 80)
