# fs_processor.py - Core Financial Statement Processor
import json
import logging
from enum import Enum
from typing import Dict, List, Optional, Tuple
from pathlib import Path
from dataclasses import dataclass

class FSType(Enum):
    """Financial Statement Types"""
    BS = "Balance Sheet"
    IS = "Income Statement"

@dataclass
class ProcessingResult:
    """Container for processing results"""
    content: str
    quality_score: float
    recommendations: List[str]
    entity_name: str
    fs_type: FSType
    date: str

class FinancialStatementProcessor:
    """Core financial statement processing orchestrator"""
    
    def __init__(self, config_path: str = "utils/config.json"):
        self.config_path = config_path
        self.config = self._load_config()
        self._setup_logging()
        
    def _load_config(self) -> Dict:
        """Load configuration from JSON file"""
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            logging.error(f"Config file not found: {self.config_path}")
            raise
        except json.JSONDecodeError as e:
            logging.error(f"Invalid JSON in config file: {e}")
            raise
    
    def _setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('fs_processor.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def get_fs_keys(self, fs_type: FSType, entity_name: str) -> List[str]:
        """Get appropriate keys based on financial statement type and entity"""
        if fs_type == FSType.BS:
            base_keys = [
                "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                "AP", "Taxes payable", "OP", "Capital", "Reserve"
            ]
            # Remove Reserve for Ningbo and Nanjing
            if entity_name in ['Ningbo', 'Nanjing']:
                base_keys = [key for key in base_keys if key != "Reserve"]
            return base_keys
        
        elif fs_type == FSType.IS:
            return [
                "Revenue", "Cost of Sales", "Gross Profit", "Operating Expenses",
                "Operating Income", "Finance Costs", "Other Income", "Net Income"
            ]
        
        else:
            raise ValueError(f"Unsupported financial statement type: {fs_type}")
    
    def get_name_mapping(self, fs_type: FSType, entity_name: str) -> Dict[str, str]:
        """Get name mapping based on financial statement type and entity"""
        if fs_type == FSType.BS:
            base_mapping = {
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
            # Add Reserve for Haining only
            if entity_name == 'Haining':
                base_mapping['Reserve'] = 'Surplus reserve'
            return base_mapping
        
        elif fs_type == FSType.IS:
            return {
                'Revenue': 'Revenue',
                'Cost of Sales': 'Cost of sales',
                'Gross Profit': 'Gross profit',
                'Operating Expenses': 'Operating expenses',
                'Operating Income': 'Operating income',
                'Finance Costs': 'Finance costs',
                'Other Income': 'Other income',
                'Net Income': 'Net income'
            }
        
        else:
            raise ValueError(f"Unsupported financial statement type: {fs_type}")
    
    def get_category_mapping(self, fs_type: FSType, entity_name: str) -> Dict[str, List[str]]:
        """Get category mapping based on financial statement type and entity"""
        if fs_type == FSType.BS:
            base_mapping = {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital']
            }
            # Add Reserve to Equity for Haining
            if entity_name == 'Haining':
                base_mapping['Equity'].append('Reserve')
            return base_mapping
        
        elif fs_type == FSType.IS:
            return {
                'Revenue and Costs': ['Revenue', 'Cost of Sales', 'Gross Profit'],
                'Operating Performance': ['Operating Expenses', 'Operating Income'],
                'Net Earnings': ['Finance Costs', 'Other Income', 'Net Income']
            }
        
        else:
            raise ValueError(f"Unsupported financial statement type: {fs_type}")
    
    def get_pattern_file_path(self, fs_type: FSType) -> str:
        """Get pattern file path based on financial statement type"""
        if fs_type == FSType.BS:
            return "utils/bs_patterns.json"
        elif fs_type == FSType.IS:
            return "utils/is_patterns.json"
        else:
            raise ValueError(f"Unsupported financial statement type: {fs_type}")
    
    def process_financial_statement(
        self,
        input_file: str,
        entity_name: str,
        date: str,
        fs_type: FSType,
        entity_helpers: Optional[List[str]] = None
    ) -> ProcessingResult:
        """Main processing method that orchestrates the entire workflow"""
        try:
            self.logger.info(f"Starting {fs_type.value} processing for {entity_name}")
            
            # Import required modules
            from ai_prompt_manager import AIPromptManager
            from data_processor import DataProcessor
            from summary_generator import SummaryGenerator
            
            # Initialize components
            prompt_manager = AIPromptManager(self.config, fs_type)
            data_processor = DataProcessor(input_file)
            summary_generator = SummaryGenerator(self.config)
            
            # Get appropriate keys and mappings
            keys = self.get_fs_keys(fs_type, entity_name)
            name_mapping = self.get_name_mapping(fs_type, entity_name)
            category_mapping = self.get_category_mapping(fs_type, entity_name)
            pattern_file = self.get_pattern_file_path(fs_type)
            
            # Process data
            results = data_processor.process_keys(
                keys=keys,
                entity_name=entity_name,
                entity_helpers=entity_helpers or [],
                pattern_file=pattern_file,
                prompt_manager=prompt_manager
            )
            
            # Generate content
            content = self._generate_markdown_content(
                results, category_mapping, name_mapping
            )
            
            # Generate summary
            summary = summary_generator.generate_summary(content, fs_type)
            
            # Quality assurance
            from quality_assurance import QualityAssuranceAgent
            qa_agent = QualityAssuranceAgent()
            quality_result = qa_agent.validate_content(content, fs_type)
            
            # Save content
            output_file = f"utils/{fs_type.name.lower()}_content.md"
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            self.logger.info(f"Content saved to {output_file}")
            
            return ProcessingResult(
                content=content,
                quality_score=quality_result.score,
                recommendations=quality_result.recommendations,
                entity_name=entity_name,
                fs_type=fs_type,
                date=date
            )
            
        except Exception as e:
            self.logger.error(f"Processing failed: {str(e)}")
            raise
    
    def _generate_markdown_content(
        self,
        results: Dict[str, str],
        category_mapping: Dict[str, List[str]],
        name_mapping: Dict[str, str]
    ) -> str:
        """Generate markdown content from processing results"""
        def extract_first_pattern_value(content: str) -> str:
            """Extract the first pattern value from AI response"""
            import re
            pattern_match = re.match(r".*?'Pattern 1': '(.*?)'.*?", content)
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
        
        return "\n".join(markdown_lines)