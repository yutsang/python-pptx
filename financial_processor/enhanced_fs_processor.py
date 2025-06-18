# enhanced_fs_processor.py - Main Execution Script
import argparse
import logging
import os
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from fs_processor import FinancialStatementProcessor, FSType

def setup_logging():
    """Setup logging configuration"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('fs_processor.log'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description='Enhanced Financial Statement Processor for Due Diligence',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Process Haining Balance Sheet:
    python enhanced_fs_processor.py -i "221128.Project TK.Databook.JW.xlsx" -e "Haining" -d "30/09/2022" -t "BS" --helpers "Wanpu" "Limited"
    
  Process Ningbo Income Statement:
    python enhanced_fs_processor.py -i "databook.xlsx" -e "Ningbo" -d "31/12/2022" -t "IS"
    
  Run with validation only mode:
    python enhanced_fs_processor.py -i "databook.xlsx" -e "Haining" -d "30/09/2022" -t "BS" --validate-only
"""
    )
    
    parser.add_argument('-i', '--input', required=True, 
                        help='Path to the Excel input file.')
    
    parser.add_argument('-e', '--entity', required=True,
                        help='Entity name (e.g., Haining, Ningbo, Nanjing).')
    
    parser.add_argument('-d', '--date', required=True,
                        help='Reporting date (e.g., 30/09/2022, 31/12/2022).')
    
    parser.add_argument('-t', '--fs-type', required=True, choices=['BS', 'IS'],
                        help='Financial statement type: BS (Balance Sheet) or IS (Income Statement).')
    
    parser.add_argument('--helpers', nargs='+', default=[],
                        help='List of entity helpers (e.g., "Wanpu" "Limited").')
    
    parser.add_argument('--validate-only', action='store_true',
                        help='Run validation on existing content without regenerating.')
    
    parser.add_argument('--output-dir', default='utils',
                        help='Directory for output files (default: utils).')
    
    parser.add_argument('--pptx-template', default=None,
                        help='Path to PowerPoint template for presentation generation.')
    
    parser.add_argument('--config', default='utils/config.json',
                        help='Path to configuration file (default: utils/config.json).')
    
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='Enable verbose logging.')
    
    return parser.parse_args()

def validate_arguments(args, logger):
    """Validate command line arguments"""
    # Check input file exists
    input_path = Path(args.input)
    if not input_path.exists():
        logger.error(f"Input file not found: {args.input}")
        return False
    
    # Check entity name is valid
    valid_entities = ['Haining', 'Ningbo', 'Nanjing']
    if args.entity not in valid_entities:
        logger.error(f"Invalid entity name: {args.entity}. Must be one of {valid_entities}")
        return False
    
    # Check date format
    import re
    if not re.match(r'^\d{2}/\d{2}/\d{4}$', args.date):
        logger.error(f"Invalid date format: {args.date}. Must be DD/MM/YYYY")
        return False
    
    # Check output directory exists or create it
    output_dir = Path(args.output_dir)
    if not output_dir.exists():
        logger.info(f"Creating output directory: {output_dir}")
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # Check PowerPoint template if specified
    if args.pptx_template and not Path(args.pptx_template).exists():
        logger.error(f"PowerPoint template not found: {args.pptx_template}")
        return False
    
    # Check config file exists
    if not Path(args.config).exists():
        logger.error(f"Configuration file not found: {args.config}")
        return False
    
    return True

def convert_fs_type(fs_type_str: str) -> FSType:
    """Convert string to FSType enum"""
    if fs_type_str == 'BS':
        return FSType.BS
    elif fs_type_str == 'IS':
        return FSType.IS
    else:
        raise ValueError(f"Invalid financial statement type: {fs_type_str}")

def validate_existing_content(args, logger):
    """Validate existing content without regenerating"""
    from quality_assurance import QualityAssuranceAgent
    
    fs_type = convert_fs_type(args.fs_type)
    content_path = Path(args.output_dir) / f"{fs_type.name.lower()}_content.md"
    
    if not content_path.exists():
        logger.error(f"Content file not found: {content_path}")
        return False
    
    try:
        with open(content_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        qa_agent = QualityAssuranceAgent()
        result = qa_agent.validate_content(content, fs_type)
        
        # Generate and save quality report
        report = qa_agent.generate_quality_report(result)
        report_path = content_path.with_name(f"{content_path.stem}_quality_report.md")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        logger.info(f"Content validation {'PASSED' if result.is_valid else 'FAILED'} with score {result.score:.1f}/100")
        logger.info(f"Quality report saved to {report_path}")
        
        return result.is_valid
        
    except Exception as e:
        logger.error(f"Error validating content: {str(e)}")
        return False

def generate_presentation(args, logger, result):
    """Generate PowerPoint presentation"""
    if not args.pptx_template:
        logger.info("PowerPoint template not specified, skipping presentation generation")
        return True
    
    try:
        from pptx_generator import PowerPointGenerator
        
        # Determine paths
        content_path = Path(args.output_dir) / f"{result.fs_type.name.lower()}_content.md"
        pptx_output = Path(args.output_dir) / f"{result.entity_name}_{result.fs_type.name}_report.pptx"
        
        # Generate presentation
        generator = PowerPointGenerator(args.pptx_template)
        generator.generate_full_report(
            content_path,
            result.content,
            str(pptx_output)
        )
        
        logger.info(f"PowerPoint presentation generated: {pptx_output}")
        
        # Update project title if needed
        from wrap_up import update_project_titles
        update_project_titles(str(pptx_output), result.entity_name)
        
        return True
        
    except Exception as e:
        logger.error(f"Error generating presentation: {str(e)}")
        return False

def main():
    """Main execution function"""
    logger = setup_logging()
    logger.info("Enhanced Financial Statement Processor")
    logger.info("--------------------------------------")
    
    # Parse and validate arguments
    args = parse_arguments()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    if not validate_arguments(args, logger):
        return 1
    
    # Validate existing content if requested
    if args.validate_only:
        logger.info("Running validation only mode")
        return 0 if validate_existing_content(args, logger) else 1
    
    try:
        # Process financial statement
        start_time = time.time()
        logger.info(f"Processing {args.fs_type} for {args.entity} as of {args.date}")
        
        processor = FinancialStatementProcessor(args.config)
        result = processor.process_financial_statement(
            input_file=args.input,
            entity_name=args.entity,
            date=args.date,
            fs_type=convert_fs_type(args.fs_type),
            entity_helpers=args.helpers
        )
        
        processing_time = time.time() - start_time
        logger.info(f"Processing completed in {processing_time:.2f} seconds")
        logger.info(f"Quality score: {result.quality_score:.1f}/100")
        
        # Generate presentation if template provided
        if args.pptx_template:
            generate_presentation(args, logger, result)
        
        return 0
        
    except Exception as e:
        logger.error(f"Processing failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())