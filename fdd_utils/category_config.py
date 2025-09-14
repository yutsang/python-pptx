"""
Category configuration for FDD application
Centralized category mappings for different statement types and entities
"""

# Display name mappings - Chinese versions
DISPLAY_NAME_MAPPING_DEFAULT_CHINESE = {
    'Cash': '现金', 'AR': '应收账款', 'Prepayments': '预付款项',
    'OR': '其他应收款', 'Other CA': '其他流动资产', 'Other NCA': '其他非流动资产',
    'IP': '投资性房地产', 'NCA': '无形资产', 'AP': '应付账款',
    'Taxes payable': '应交税费', 'OP': '其他应付款', 'Capital': '股本',
    'Reserve': '资本公积', 'OI': '营业收入', 'OC': '营业成本',
    'Tax and Surcharges': '税金及附加', 'GA': '管理费用', 'Fin Exp': '财务费用',
    'Cr Loss': '信用损失', 'Other Income': '其他收益',
    'Non-operating Income': '营业外收入', 'Non-operating Exp': '营业外支出',
    'Income tax': '所得税', 'LT DTA': '递延所得税资产'
}

DISPLAY_NAME_MAPPING_NB_NJ_CHINESE = {
    'Cash': '现金', 'AR': '应收账款', 'Prepayments': '预付款项',
    'OR': '其他应收款', 'Other CA': '其他流动资产', 'Other NCA': '其他非流动资产',
    'IP': '投资性房地产', 'NCA': '无形资产', 'AP': '应付账款',
    'Taxes payable': '应交税费', 'OP': '其他应付款', 'Capital': '股本',
    'Reserve': '资本公积', 'OI': '营业收入', 'OC': '营业成本',
    'Tax and Surcharges': '税金及附加', 'GA': '管理费用', 'Fin Exp': '财务费用',
    'Cr Loss': '信用损失', 'Other Income': '其他收益',
    'Non-operating Income': '营业外收入', 'Non-operating Exp': '营业外支出',
    'Income tax': '所得税', 'LT DTA': '递延所得税资产'
}

# Display name mappings - English versions
DISPLAY_NAME_MAPPING_DEFAULT_ENGLISH = {
    'Cash': 'Cash at bank', 'AR': 'Accounts receivables', 'Prepayments': 'Prepayments',
    'OR': 'Other receivables', 'Other CA': 'Other current assets', 'Other NCA': 'Other non-current assets',
    'IP': 'Investment properties', 'NCA': 'Intangible assets', 'AP': 'Accounts payable',
    'Taxes payable': 'Taxes payables', 'OP': 'Other payables', 'Capital': 'Capital',
    'Reserve': 'Surplus reserve', 'OI': 'Operating Income', 'OC': 'Operating Cost',
    'Tax and Surcharges': 'Tax and Surcharges', 'GA': 'G&A expenses', 'Fin Exp': 'Finance Expenses',
    'Cr Loss': 'Credit Losses', 'Other Income': 'Other Income',
    'Non-operating Income': 'Non-operating Income', 'Non-operating Exp': 'Non-operating Expenses',
    'Income tax': 'Income tax', 'LT DTA': 'Long-term Deferred Tax Assets'
}

DISPLAY_NAME_MAPPING_NB_NJ_ENGLISH = {
    'Cash': 'Cash at bank', 'AR': 'Accounts receivables', 'Prepayments': 'Prepayments',
    'OR': 'Other receivables', 'Other CA': 'Other current assets', 'Other NCA': 'Other non-current assets',
    'IP': 'Investment properties', 'NCA': 'Intangible assets', 'AP': 'Accounts payable',
    'Taxes payable': 'Taxes payables', 'OP': 'Other payables', 'Capital': 'Capital',
    'Reserve': 'Surplus reserve', 'OI': 'Operating Income', 'OC': 'Operating Cost',
    'Tax and Surcharges': 'Tax and Surcharges', 'GA': 'G&A expenses', 'Fin Exp': 'Finance Expenses',
    'Cr Loss': 'Credit Losses', 'Other Income': 'Other Income',
    'Non-operating Income': 'Non-operating Income', 'Non-operating Exp': 'Non-operating Expenses',
    'Income tax': 'Income tax', 'LT DTA': 'Long-term Deferred Tax Assets'
}

# Backward compatibility - default to Chinese for now
DISPLAY_NAME_MAPPING_DEFAULT = DISPLAY_NAME_MAPPING_DEFAULT_CHINESE
DISPLAY_NAME_MAPPING_NB_NJ = DISPLAY_NAME_MAPPING_NB_NJ_CHINESE

def get_category_mapping(statement_type, entity_name, language='chinese'):
    """
    Get category mapping based on statement type, entity name, and language
    """
    # Select the appropriate display name mapping based on language and entity
    if language.lower() == 'english':
        if entity_name in ['Ningbo', 'Nanjing']:
            name_mapping = DISPLAY_NAME_MAPPING_NB_NJ_ENGLISH
        else:
            name_mapping = DISPLAY_NAME_MAPPING_DEFAULT_ENGLISH
    else:  # Chinese (default)
        if entity_name in ['Ningbo', 'Nanjing']:
            name_mapping = DISPLAY_NAME_MAPPING_NB_NJ_CHINESE
        else:
            name_mapping = DISPLAY_NAME_MAPPING_DEFAULT_CHINESE
    
    if statement_type == "IS":
        # Income Statement categories
        return {
            'Revenue': ['OI', 'Other Income', 'Non-operating Income'],
            'Expenses': ['OC', 'GA', 'Fin Exp', 'Cr Loss', 'Non-operating Exp'],
            'Taxes': ['Tax and Surcharges', 'Income tax'],
            'Other': ['LT DTA']
        }, name_mapping
    else:
        # Balance Sheet categories (default)
        if entity_name in ['Ningbo', 'Nanjing']:
            return {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital']
            }, name_mapping
        else:  # Haining and others
            return {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital', 'Reserve']
            }, name_mapping

def get_statement_keys(statement_type):
    """
    Get keys for a specific statement type
    """
    if statement_type == "IS":
        return ['OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss',
                'Other Income', 'Non-operating Income', 'Non-operating Exp',
                'Income tax', 'LT DTA']
    elif statement_type == "BS":
        return ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA', 'Other NCA',
                'IP', 'NCA', 'AP', 'Taxes payable', 'OP', 'Capital', 'Reserve']
    else:  # ALL
        return ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA', 'Other NCA',
                'IP', 'NCA', 'AP', 'Taxes payable', 'OP', 'Capital', 'Reserve',
                'OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss',
                'Other Income', 'Non-operating Income', 'Non-operating Exp',
                'Income tax', 'LT DTA']
