# ai_prompt_manager.py - Enhanced AI Prompt Management System
import json
import logging
from typing import Dict, List, Optional
from fs_processor import FSType

class AIPromptManager:
    """Enhanced AI prompt management with role-specific templates"""
    
    def __init__(self, config: Dict, fs_type: FSType):
        self.config = config
        self.fs_type = fs_type
        self.logger = logging.getLogger(__name__)
        
    def get_system_prompt(self) -> str:
        """Get role-specific system prompt based on financial statement type"""
        if self.fs_type == FSType.BS:
            return self._get_bs_system_prompt()
        elif self.fs_type == FSType.IS:
            return self._get_is_system_prompt()
        else:
            raise ValueError(f"Unsupported financial statement type: {self.fs_type}")
    
    def _get_bs_system_prompt(self) -> str:
        """System prompt for Balance Sheet analysis"""
        return """
Role: Senior Financial Analyst - Balance Sheet Specialist

You are a senior financial analyst with 15+ years of experience in due diligence, asset valuation, and balance sheet analysis. You specialize in:
- Asset quality assessment and liquidity analysis
- Liability structure evaluation and solvency metrics
- Capital adequacy and equity composition analysis
- Working capital management and operational efficiency
- Risk assessment and financial stability indicators

CORE EXPERTISE:
- Deep understanding of accounting standards (IFRS/GAAP)
- Expert in asset valuation methodologies
- Comprehensive knowledge of liability management
- Proven experience in due diligence reporting
- Strong analytical skills in financial ratio analysis

TASK REQUIREMENTS:
1. SELECT exactly one appropriate pattern from provided options
2. Replace ALL placeholder values with corresponding actual data
3. Output ONLY the completed financial narrative text
4. NEVER show template structure or pattern names
5. Use ONLY provided data - never estimate or extrapolate
6. Express figures appropriately (K for thousands, M for millions)
7. Translate any foreign content to clear business English
8. Focus on material figures and key insights
9. Ensure audit-grade accuracy and professional presentation
10. Maintain consistency with industry standards

VALIDATION CHECKLIST:
✓ Pattern selected based on data completeness
✓ All placeholders replaced with actual values
✓ Currency conversions accurate (K/M notation)
✓ Foreign content translated to English
✓ Professional business language used
✓ No template artifacts or JSON structure
✓ Material figures highlighted appropriately
✓ Consistent with balance sheet principles

OUTPUT REQUIREMENTS:
- Professional audit-grade financial narrative
- Clear, concise business English
- Appropriate figure scaling and formatting
- Focus on balance sheet analysis principles
- NO pattern names, templates, or explanations
"""
    
    def _get_is_system_prompt(self) -> str:
        """System prompt for Income Statement analysis"""
        return """
Role: Senior Financial Analyst - Income Statement Specialist

You are a senior financial analyst with 15+ years of experience in performance analysis, revenue recognition, and earnings quality assessment. You specialize in:
- Revenue analysis and recognition patterns
- Cost structure evaluation and margin analysis
- Operating efficiency and performance metrics
- Earnings quality and sustainability assessment
- Profitability drivers and trend analysis

CORE EXPERTISE:
- Expert in revenue recognition standards
- Deep understanding of cost accounting principles
- Comprehensive knowledge of operating leverage
- Proven experience in performance analysis
- Strong analytical skills in profitability assessment

TASK REQUIREMENTS:
1. SELECT exactly one appropriate pattern from provided options
2. Replace ALL placeholder values with corresponding actual data
3. Output ONLY the completed financial narrative text
4. NEVER show template structure or pattern names
5. Use ONLY provided data - never estimate or extrapolate
6. Express figures appropriately (K for thousands, M for millions)
7. Translate any foreign content to clear business English
8. Focus on performance indicators and key insights
9. Ensure audit-grade accuracy and professional presentation
10. Maintain consistency with income statement principles

VALIDATION CHECKLIST:
✓ Pattern selected based on data completeness
✓ All placeholders replaced with actual values
✓ Currency conversions accurate (K/M notation)
✓ Foreign content translated to English
✓ Professional business language used
✓ No template artifacts or JSON structure
✓ Performance metrics highlighted appropriately
✓ Consistent with income statement principles

OUTPUT REQUIREMENTS:
- Professional audit-grade financial narrative
- Clear, concise business English
- Appropriate figure scaling and formatting
- Focus on income statement analysis principles
- NO pattern names, templates, or explanations
"""
    
    def get_user_prompt(
        self,
        key: str,
        pattern: Dict,
        financial_figure: str,
        excel_tables: str,
        entity_name: str,
        detect_zeros: str = ""
    ) -> str:
        """Generate comprehensive user prompt with validation requirements"""
        
        base_prompt = f"""
TASK: Select ONE pattern and complete it with actual data

AVAILABLE PATTERNS: {json.dumps(pattern, indent=2)}

FINANCIAL FIGURE: {key}: {financial_figure}

DATA SOURCE: {excel_tables}

ENTITY CONTEXT: {entity_name}

SELECTION CRITERIA:
- Choose the pattern with the most complete data coverage
- Prioritize patterns that match the primary account category
- Use most recent data available
- Ensure all placeholders can be filled with actual data
{detect_zeros}

VALIDATION REQUIREMENTS:
- Verify all numbers match the FINANCIAL FIGURE context
- Cross-reference data source for accuracy
- Ensure entity names in template are from DATA SOURCE (not {entity_name} itself)
- Check totals and sub-components for mathematical consistency
- Confirm appropriate currency scaling (K/M notation)

REQUIRED OUTPUT FORMAT:
- Only the completed pattern text
- No pattern names, labels, or JSON formatting
- No template structure or placeholders
- Replace ALL 'xxx' or brackets with actual data values
- No bullet points for narrative descriptions
- Professional business English only
- No foreign language content
- No extra explanations or comments

QUALITY STANDARDS:
- Audit-grade accuracy and presentation
- Clear, professional financial narrative
- Appropriate materiality focus
- Industry-standard terminology
- Consistent formatting and style

Example CORRECT output format:
"Cash at bank comprises deposits of $2.3M held with major financial institutions as at 30/09/2022. The deposits are unrestricted and available for operations."

Example INCORRECT output format:
"Pattern 1: Cash at bank comprises deposits of [amount] held with [institution] as at [date]."
"""
        
        # Add FS-specific requirements
        if self.fs_type == FSType.BS:
            fs_specific = """
BALANCE SHEET FOCUS:
- Emphasize asset quality and liquidity characteristics
- Highlight liability terms and covenant compliance
- Address capital structure and equity composition
- Include relevant ratios and financial metrics
- Focus on balance sheet strength indicators
"""
        else:  # Income Statement
            fs_specific = """
INCOME STATEMENT FOCUS:
- Emphasize revenue quality and recognition patterns
- Highlight margin trends and cost structure
- Address operating efficiency indicators
- Include relevant performance metrics
- Focus on earnings sustainability factors
"""
        
        return base_prompt + fs_specific
    
    def validate_output(self, output: str, key: str) -> Dict[str, any]:
        """Validate AI output against quality standards"""
        validation_result = {
            'is_valid': True,
            'issues': [],
            'score': 100
        }
        
        # Check for template artifacts
        template_artifacts = [
            'Pattern 1:', 'Pattern 2:', 'Pattern 3:',
            '[', ']', 'xxx', '{', '}',
            'template', 'placeholder'
        ]
        
        for artifact in template_artifacts:
            if artifact.lower() in output.lower():
                validation_result['issues'].append(f"Template artifact found: {artifact}")
                validation_result['score'] -= 20
        
        # Check for professional language
        if len(output.split()) < 10:
            validation_result['issues'].append("Output too brief for professional narrative")
            validation_result['score'] -= 15
        
        # Check for key presence
        if key.lower() not in output.lower():
            validation_result['issues'].append(f"Key '{key}' not clearly referenced in output")
            validation_result['score'] -= 10
        
        # Set validity based on score
        validation_result['is_valid'] = validation_result['score'] >= 70
        
        return validation_result