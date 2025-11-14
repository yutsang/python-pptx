"""
UI configuration constants for FDD application
Centralized location for UI-related constants and messages
"""

# Pattern analysis required elements
PATTERN_REQUIRED_ELEMENTS = ['balance', 'CNY', 'represented']

# Balance sheet categories for consistency check
BALANCE_SHEET_CATEGORIES = {
    'current_assets': ['Cash', 'AR', 'Prepayments', 'Other CA'],
    'non_current_assets': ['IP', 'Other NCA'],
    'liabilities': ['AP', 'Taxes payable', 'OP'],
    'equity': ['Capital', 'Reserve']
}

# Balance sheet consistency messages
BALANCE_SHEET_MESSAGES = {
    'current_assets': "✅ Current Asset - Data structure appears consistent",
    'non_current_assets': "✅ Non-Current Asset - Data structure appears consistent",
    'liabilities': "✅ Liability - Data structure appears consistent",
    'equity': "✅ Equity - Data structure appears consistent"
}

# Conversation flow description
CONVERSATION_FLOW_DESCRIPTION = """
**Message Sequence:**
1. **System Message**: Sets the AI's role and expertise
2. **Assistant Message**: Provides context data from financial statements
3. **User Message**: Specific analysis request for the financial key
"""

# Generic prompt fallback template
GENERIC_PROMPT_TEMPLATE = """
**System Prompt:**
{system_prompt}

**User Prompt:**
Analyze the {key_display_name} position:

1. **Current Balance**: Review the current balance and composition
2. **Trend Analysis**: Assess historical trends and changes
3. **Risk Assessment**: Evaluate any associated risks
4. **Business Impact**: Consider the impact on business operations
5. **Future Outlook**: Assess future expectations and plans

**Key Questions to Address:**
- What is the current balance and its composition?
- How has this changed over time?
- What are the key drivers of this position?
- Are there any unusual items or concentrations?
- How does this compare to industry norms?

**Analysis Requirements:**
- Focus on material items and key insights
- Highlight any significant changes or trends
- Identify potential risk areas or concerns
- Provide clear, actionable insights for management

**Expected Output:**
- Clear and concise analysis
- Identification of key trends and drivers
- Assessment of any potential issues or opportunities
- Professional financial analysis language
"""

def get_pattern_required_elements():
    """Get required elements for pattern analysis"""
    return PATTERN_REQUIRED_ELEMENTS

def get_balance_sheet_categories():
    """Get balance sheet categories for consistency check"""
    return BALANCE_SHEET_CATEGORIES

def get_balance_sheet_message(category_type):
    """Get consistency message for balance sheet category"""
    return BALANCE_SHEET_MESSAGES.get(category_type, "✅ Data structure appears consistent")

def get_conversation_flow_description():
    """Get conversation flow description"""
    return CONVERSATION_FLOW_DESCRIPTION

def get_generic_prompt_template():
    """Get generic prompt template for fallback"""
    return GENERIC_PROMPT_TEMPLATE

def get_category_for_key(key):
    """Determine which category a key belongs to"""
    categories = get_balance_sheet_categories()
    for category, keys in categories.items():
        if key in keys:
            return category
    return None
