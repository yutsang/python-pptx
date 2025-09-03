"""
Category configuration for FDD application
Centralized category mappings for different statement types and entities
"""

# Display name mappings
DISPLAY_NAME_MAPPING_DEFAULT = {
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

DISPLAY_NAME_MAPPING_NB_NJ = {
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

def get_category_mapping(statement_type, entity_name):
    """
    Get category mapping based on statement type and entity name
    """
    if statement_type == "IS":
        # Income Statement categories
        if entity_name in ['Ningbo', 'Nanjing']:
            return {
                'Revenue': ['OI', 'Other Income', 'Non-operating Income'],
                'Expenses': ['OC', 'GA', 'Fin Exp', 'Cr Loss', 'Non-operating Exp'],
                'Taxes': ['Tax and Surcharges', 'Income tax'],
                'Other': ['LT DTA']
            }, DISPLAY_NAME_MAPPING_NB_NJ
        else:  # Haining and others
            return {
                'Revenue': ['OI', 'Other Income', 'Non-operating Income'],
                'Expenses': ['OC', 'GA', 'Fin Exp', 'Cr Loss', 'Non-operating Exp'],
                'Taxes': ['Tax and Surcharges', 'Income tax'],
                'Other': ['LT DTA']
            }, DISPLAY_NAME_MAPPING_DEFAULT
    else:
        # Balance Sheet categories (default)
        if entity_name in ['Ningbo', 'Nanjing']:
            return {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital']
            }, DISPLAY_NAME_MAPPING_NB_NJ
        else:  # Haining and others
            return {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital', 'Reserve']
            }, DISPLAY_NAME_MAPPING_DEFAULT

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
