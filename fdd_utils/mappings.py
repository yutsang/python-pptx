"""
Centralized mapping tables used across the app for consistency and reuse.
"""

# Mapping from financial keys to section headers used in offline markdown content
KEY_TO_SECTION_MAPPING = {
    'Cash': 'Cash at bank',
    'AR': 'Accounts receivables',
    'Prepayments': 'Prepayments',
    'OR': 'Other receivables',
    'Other CA': 'Other current assets',
    'IP': 'Investment properties',
    'Other NCA': 'Other non-Current assets',
    'AP': 'Accounts payable',
    'Advances': 'Advances',
    'Taxes payable': 'Taxes payables',
    'OP': 'Other payables',
    'Capital': 'Capital',
    'Reserve': 'Surplus reserve',
    'Capital reserve': 'Capital reserve',
    'OI': 'Other Income',
    'OC': 'Other Costs',
    'Tax and Surcharges': 'Tax and Surcharges',
    'GA': 'G&A expenses',
    'Fin Exp': 'Finance Expenses',
    'Cr Loss': 'Credit Losses',
    'Other Income': 'Other Income',
    'Non-operating Income': 'Non-operating Income',
    'Non-operating Exp': 'Non-operating Expenses',
    'Income tax': 'Income tax',
    'LT DTA': 'Long-term Deferred Tax Assets'
}

# Key terms used for row-highlighting during offline data validation
KEY_TERMS_BY_KEY = {
    # Balance Sheet keys
    'Cash': ['cash', 'bank', 'deposit'],
    'AR': ['receivable', 'receivables', 'ar'],
    'AP': ['payable', 'payables', 'ap'],
    'IP': ['investment', 'property', 'properties'],
    'Capital': ['capital', 'share', 'equity'],
    'Reserve': ['reserve', 'surplus'],
    'Taxes payable': ['tax', 'taxes', 'taxable'],
    'OP': ['other', 'payable', 'payables'],
    'Prepayments': ['prepayment', 'prepaid'],
    'OR': ['other', 'receivable', 'receivables'],
    'Other CA': ['other', 'current', 'asset'],
    'Other NCA': ['other', 'non-current', 'asset'],
    # Income Statement keys
    'OI': ['other income', 'income', 'revenue'],
    'OC': ['other cost', 'cost', 'expense'],
    'Tax and Surcharges': ['tax', 'surcharge', 'vat', 'gst'],
    'GA': ['g&a', 'general', 'administrative', 'expense'],
    'Fin Exp': ['finance', 'financial', 'expense', 'interest'],
    'Cr Loss': ['credit loss', 'bad debt', 'provision'],
    'Other Income': ['other income', 'income', 'gain'],
    'Non-operating Income': ['non-operating', 'income', 'gain'],
    'Non-operating Exp': ['non-operating', 'expense', 'loss'],
    'Income tax': ['income tax', 'tax expense', 'corporate tax'],
    'LT DTA': ['deferred tax', 'dta', 'long-term']
}

# Display name mapping used in multiple places (session content, export, etc.)
DISPLAY_NAME_MAPPING_DEFAULT = {
    # Balance Sheet items
    'Cash': 'Cash at bank',
    'AR': 'Accounts receivables',
    'Prepayments': 'Prepayments',
    'OR': 'Other receivables',
    'Other CA': 'Other current assets',
    'IP': 'Investment properties',
    'Other NCA': 'Other non-current assets',
    'AP': 'Accounts payable',
    'Taxes payable': 'Taxes payables',
    'OP': 'Other payables',
    'Capital': 'Capital',
    'Reserve': 'Surplus reserve',
    # Income Statement items
    'OI': 'Other Income',
    'OC': 'Other Costs',
    'Tax and Surcharges': 'Tax and Surcharges',
    'GA': 'G&A expenses',
    'Fin Exp': 'Finance Expenses',
    'Cr Loss': 'Credit Losses',
    'Other Income': 'Other Income',
    'Non-operating Income': 'Non-operating Income',
    'Non-operating Exp': 'Non-operating Expenses',
    'Income tax': 'Income tax',
    'LT DTA': 'Long-term Deferred Tax Assets'
}

DISPLAY_NAME_MAPPING_NB_NJ = {
    # Balance Sheet items
    'Cash': 'Cash at bank',
    'AR': 'Accounts receivables',
    'Prepayments': 'Prepayments',
    'OR': 'Other receivables',
    'Other CA': 'Other current assets',
    'IP': 'Investment properties',
    'Other NCA': 'Other non-current assets',
    'AP': 'Accounts payable',
    'Taxes payable': 'Taxes payables',
    'OP': 'Other payables',
    'Capital': 'Capital',
    # Income Statement items (same for all entities)
    'OI': 'Other Income',
    'OC': 'Other Costs',
    'Tax and Surcharges': 'Tax and Surcharges',
    'GA': 'G&A expenses',
    'Fin Exp': 'Finance Expenses',
    'Cr Loss': 'Credit Losses',
    'Other Income': 'Other Income',
    'Non-operating Income': 'Non-operating Income',
    'Non-operating Exp': 'Non-operating Expenses',
    'Income tax': 'Income tax',
    'LT DTA': 'Long-term Deferred Tax Assets'
}



