import argparse, warnings, re, os, time
from utils.utils import process_keys

warnings.filterwarnings(
    "ignore"
    message='Data Validation extension is not supported and will be removed',
    category=UserWarning,
    module='openpyxl'
)

def main(input_file_path, entity_name, entity_helpers):
    keys = [
        "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
        "AP", "Taxes payable", "OP", "Capital", "Reserve"
    ]
    if entity_name == 'Ningbo' or entity_name == 'Nanjing':
        keys = [key for key in keys if key != "Reserve"]
    
    mapping_file_path = "utils/mapping.json"
    pattern_fule_path = "utils/pattern.json"
    
    # Process the keys and get the response
    results = process_keys(
        keys=keys,
        entity_name=entity_name,
        entity_helpers=entity_helpers,
        input_file=input_file_path,
        mapping_file=mapping_file_path,
        pattern_file=pattern_fule_path
    )
    if entity_name == 'Ningbo' or entity_name == 'Nanjing':
        name_mapping = {
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
            'Capital': 'Capital'
        }
        category_mapping = {
            'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
            'Non-current Assets': ['IP', 'Other NCA'],
            'Liabilities': ['AP', 'Taxes payable', 'OP'],
            'Equity': ['Capital']
        }
    elif entity_name == 'Haining':
        name_mapping = {
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
            'Reserve': 'Surplus reserve'
        }
        category_mapping = {
            'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
            'Non-current Assets': ['IP', 'Other NCA'],
            'Liabilities': ['AP', 'Taxes payable', 'OP'],
            'Equity': ['Capital', 'Reserve']
        }
        
    def extract_first_pattern_value(content):
        pattern_match = re.match(r"\.{.*?'Pattern 1': '(.*?)'.*?\}", content)
        if pattern_match:
            return pattern_match.group(1)
        return content
    
    markdown_lines = []
    for category, items in category_mapping.items():
        markdown_lines.append(f"## {category}\n")
        for item in items:
            full_name = name_mapping[item]
            info = results.get(item, f"No information available for {item}")
            
            description = extract_first_pattern_value(info)
            
            markdown_lines.append(f"### {full_name}\n{description}\n")
            
    markdown_text = "\n".join(markdown_lines)
    
    file_path = 'utils/bs_content.md'
    
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(markdown_text)
        
    print(f"Markdown saved to {file_path}")
    
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process keys and generate markdown content.')
    parser.add_argument('-i', '--input', required=True, help='Path to the Excel input file.')
    parser.add_argument('-e', '--entity', required=True, help='Entity name.')
    parser.add_argument('--helpers', nargs='+', default=[], help='List of entity helpers.')
    
    args = parser.parse_args()
    main(args.input, args.entity, args.helpers)
    
# python 1.ai_fs_content.py -i "221128.Project TK.Databook.JW.xlsx" -e "Haining" --helpers "Wanpu" "Limited" ""
        