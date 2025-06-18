# data_processor.py - Robust Data Processing Utilities
import pandas as pd
import json
import re
import warnings
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from tqdm import tqdm
import urllib3

# Suppress warnings
urllib3.disable_warnings()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)

class DataProcessor:
    """Enhanced data processing with improved error handling and validation"""
    
    def __init__(self, input_file: str):
        self.input_file = input_file
        self.main_dir = Path(__file__).parent.parent if Path(__file__).parent.parent.exists() else Path('.')
        self.file_path = self.main_dir / input_file
        
        # Validate file exists
        if not self.file_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.file_path}")
    
    def process_keys(
        self,
        keys: List[str],
        entity_name: str,
        entity_helpers: List[str],
        pattern_file: str,
        prompt_manager
    ) -> Dict[str, str]:
        """Process financial statement keys with enhanced error handling"""
        
        # Load required data
        financial_figures = self._find_financial_figures_with_context_check(entity_name)
        mapping_file_path = "utils/mapping.json"
        
        # Initialize results dictionary
        results = {}
        
        # Initialize progress bar
        pbar = tqdm(keys, desc="Processing keys", unit="key")
        
        for key in pbar:
            try:
                # Load pattern and mapping
                pattern = self._load_pattern(pattern_file, key)
                mapping = {key: self._load_mapping(mapping_file_path, key)}
                
                # Process Excel tables
                excel_tables = self._process_and_filter_excel(mapping, entity_name, entity_helpers)
                
                # Get financial figure
                financial_figure = self._get_financial_figure(financial_figures, key)
                
                # Detect thousands notation
                detect_zeros = self._detect_thousands_notation(excel_tables)
                
                # Generate AI response
                response = self._generate_ai_response(
                    key, pattern, financial_figure, excel_tables, 
                    entity_name, detect_zeros, prompt_manager
                )
                
                # Store result
                results[key] = response
                
                # Update progress
                pbar.set_postfix_str(f"key={key}")
                
            except Exception as e:
                self.logger.error(f"Error processing key {key}: {str(e)}")
                results[key] = f"Error processing {key}: {str(e)}"
        
        return results
    
    def _find_financial_figures_with_context_check(self, entity_name: str) -> Dict[str, float]:
        """Enhanced financial figure extraction with better error handling"""
        try:
            sheet_name = self._get_tab_name(entity_name)
            date_str = '30/09/2022'  # Default date, could be parameterized
            
            xl = pd.ExcelFile(self.file_path)
            
            if sheet_name not in xl.sheet_names:
                raise ValueError(f"Sheet '{sheet_name}' not found in file")
            
            # Parse the sheet
            df = xl.parse(sheet_name)
            
            # Standardize column names
            if len(df.columns) >= 4:
                df.columns = ['Description', 'Date_2020', 'Date_2021', 'Date_2022']
            else:
                raise ValueError(f"Insufficient columns in sheet {sheet_name}")
            
            # Date column mapping
            date_column_map = {
                '31/12/2020': 'Date_2020',
                '31/12/2021': 'Date_2021',
                '30/09/2022': 'Date_2022'
            }
            
            if date_str not in date_column_map:
                raise ValueError(f"Date '{date_str}' not recognized")
            
            date_column = date_column_map[date_str]
            
            # Scale factor for thousands notation
            scale_factor = 1000
            
            # Financial figure mapping
            financial_figure_map = {
                "Cash": "Cash at bank",
                "AR": "Accounts receivable",
                "Prepayments": "Prepayments",
                "OR": "Other receivables",
                "Other CA": "Other current assets",
                "IP": "Investment properties",
                "Other NCA": "Other non-current assets",
                "AP": "Accounts payable",
                "Taxes payable": "Taxes payable",
                "OP": "Other payables",
                "Capital": "Paid-in capital",
                "Reserve": "Surplus reserve"
            }
            
            financial_figures = {}
            
            # Extract figures
            for key, desc in financial_figure_map.items():
                try:
                    mask = df['Description'].str.contains(desc, case=False, na=False)
                    values = df.loc[mask, date_column].values
                    
                    if len(values) > 0 and pd.notna(values[0]):
                        financial_figures[key] = float(values[0]) / scale_factor
                    else:
                        financial_figures[key] = 0.0
                        
                except Exception as e:
                    print(f"Warning: Could not extract figure for {key}: {e}")
                    financial_figures[key] = 0.0
            
            return financial_figures
            
        except Exception as e:
            print(f"Error extracting financial figures: {e}")
            return {}
    
    def _get_tab_name(self, entity_name: str) -> str:
        """Get sheet tab name based on entity"""
        tab_mapping = {
            'Haining': 'BSHN',
            'Nanjing': 'BSNJ',
            'Ningbo': 'BSNB'
        }
        
        if entity_name not in tab_mapping:
            raise ValueError(f"Unknown entity: {entity_name}")
            
        return tab_mapping[entity_name]
    
    def _get_financial_figure(self, financial_figures: Dict[str, float], key: str) -> str:
        """Format financial figure with appropriate scaling"""
        figure = financial_figures.get(key, None)
        
        if figure is None:
            return f"{key} not found in the financial figures."
        
        if figure >= 1000000:
            return f"{figure / 1000000:.1f}M"
        elif figure >= 1000:
            return f"{figure / 1000:,.0f}K"
        else:
            return f"{figure:.1f}"
    
    def _load_pattern(self, pattern_file: str, key: str) -> Dict:
        """Load pattern from JSON file"""
        try:
            with open(pattern_file, 'r') as f:
                patterns = json.load(f)
            return patterns.get(key, {})
        except FileNotFoundError:
            print(f"Pattern file not found: {pattern_file}")
            return {}
        except json.JSONDecodeError:
            print(f"Error decoding JSON from pattern file: {pattern_file}")
            return {}
    
    def _load_mapping(self, mapping_file: str, key: str) -> List[str]:
        """Load mapping from JSON file"""
        try:
            with open(mapping_file, 'r') as f:
                mappings = json.load(f)
            return mappings.get(key, [])
        except FileNotFoundError:
            print(f"Mapping file not found: {mapping_file}")
            return []
        except json.JSONDecodeError:
            print(f"Error decoding JSON from mapping file: {mapping_file}")
            return []
    
    def _process_and_filter_excel(
        self,
        tab_name_mapping: Dict[str, List[str]],
        entity_name: str,
        entity_suffixes: List[str]
    ) -> str:
        """Enhanced Excel processing with better error handling"""
        try:
            xl = pd.ExcelFile(self.file_path)
            
            # Create reverse mapping
            reverse_mapping = {}
            for key, values in tab_name_mapping.items():
                for value in values:
                    reverse_mapping[value] = key
            
            markdown_content = ""
            
            # Process each sheet
            for sheet_name in xl.sheet_names:
                if sheet_name in reverse_mapping:
                    try:
                        df = xl.parse(sheet_name)
                        
                        # Split dataframes on empty rows
                        empty_rows = df.index[df.isnull().all(1)]
                        start_idx = 0
                        dataframes = []
                        
                        for end_idx in empty_rows:
                            if end_idx > start_idx:
                                split_df = df[start_idx:end_idx]
                                if not split_df.dropna(how='all').empty:
                                    dataframes.append(split_df)
                            start_idx = end_idx + 1
                        
                        if start_idx < len(df):
                            dataframes.append(df[start_idx:])
                        
                        # Filter by entity
                        entity_keywords = [f"{entity_name}{suffix}" for suffix in entity_suffixes]
                        combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                        
                        for data_frame in dataframes:
                            mask = data_frame.apply(
                                lambda row: row.astype(str).str.contains(
                                    combined_pattern, case=False, regex=True, na=False
                                ).any(),
                                axis=1
                            )
                            
                            if mask.any():
                                from tabulate import tabulate
                                markdown_content += tabulate(
                                    data_frame, headers='keys', tablefmt='pipe'
                                ) + '\n\n'
                    
                    except Exception as e:
                        print(f"Error processing sheet {sheet_name}: {e}")
                        continue
            
            return markdown_content
            
        except Exception as e:
            print(f"Error processing Excel file: {e}")
            return ""
    
    def _detect_thousands_notation(self, excel_tables: str) -> str:
        """Detect if data is in thousands notation"""
        return "3. The figures in this table is already expressed in k, express the number in M " \
               "(divide by 1000), rounded to 1 decimal place, if final figure less than 1M, express in K (no decimal places)." \
               if "'000" in excel_tables else ""
    
    def _generate_ai_response(
        self,
        key: str,
        pattern: Dict,
        financial_figure: str,
        excel_tables: str,
        entity_name: str,
        detect_zeros: str,
        prompt_manager
    ) -> str:
        """Generate AI response using prompt manager"""
        try:
            # Load AI services
            from utils.ai_helper import load_config, initialize_ai_services, generate_response
            
            config_details = load_config("utils/config.json")
            oai_client, search_client = initialize_ai_services(config_details)
            
            # Get prompts
            system_prompt = prompt_manager.get_system_prompt()
            user_prompt = prompt_manager.get_user_prompt(
                key, pattern, financial_figure, excel_tables, entity_name, detect_zeros
            )
            
            # Generate response
            response = generate_response(
                user_prompt,
                system_prompt,
                oai_client,
                "",  # No context content needed
                config_details['CHAT_MODEL']
            )
            
            # Validate response
            validation_result = prompt_manager.validate_output(response, key)
            
            if not validation_result['is_valid']:
                print(f"Warning: Quality issues detected for {key}: {validation_result['issues']}")
            
            return response
            
        except Exception as e:
            print(f"Error generating AI response for {key}: {e}")
            return f"Error generating content for {key}"