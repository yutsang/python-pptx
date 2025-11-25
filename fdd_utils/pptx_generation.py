#!/usr/bin/env python3
"""
PowerPoint Generation Module for Financial Reports
Based on the backup methods but implemented fresh for the new system
"""

import os
import re
import logging
from typing import Dict, List, Optional, Tuple
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def get_tab_name(project_name: str) -> Optional[str]:
    """Get tab name from project name for Excel embedding"""
    if not project_name:
        return None

    # Extract key identifier from project name
    # e.g., "东莞联洋利润表" -> "东莞联洋"
    words = project_name.split()
    if words:
        return words[0]
    return None


def clean_content_quotes(content: str) -> str:
    """Clean and format content quotes"""
    if not content:
        return ""

    # Remove excessive quotes and clean formatting
    content = re.sub(r'^"*|"*$', '', content.strip())
    content = re.sub(r'""+', '"', content)

    return content


def detect_chinese_text(text: str, force_chinese_mode: bool = False) -> bool:
    """Detect if text contains significant Chinese characters"""
    if force_chinese_mode:
        return True

    if not text:
        return False

    chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
    total_chars = len(text)

    if total_chars == 0:
        return False

    # Consider it Chinese if > 30% Chinese characters
    return (chinese_chars / total_chars) > 0.3


def get_font_size_for_text(text: str, base_size: int = 9, force_chinese_mode: bool = False) -> Pt:
    """Get appropriate font size for text content"""
    is_chinese = detect_chinese_text(text, force_chinese_mode)

    if is_chinese:
        # Slightly larger for Chinese readability
        return Pt(base_size + 1)
    else:
        return Pt(base_size)


def get_font_name_for_text(text: str, default_font: str = 'Arial') -> str:
    """Get appropriate font name for text"""
    is_chinese = detect_chinese_text(text)

    if is_chinese:
        return 'Microsoft YaHei'  # or 'SimSun' as fallback
    else:
        return default_font


def get_line_spacing_for_text(text: str, force_chinese_mode: bool = False) -> float:
    """Get line spacing for text"""
    is_chinese = detect_chinese_text(text, force_chinese_mode)
    return 0.9 if is_chinese else 1.0


def get_space_after_for_text(text: str, force_chinese_mode: bool = False) -> Pt:
    """Get space after for text paragraphs"""
    is_chinese = detect_chinese_text(text, force_chinese_mode)
    return Pt(6) if is_chinese else Pt(4)


def get_space_before_for_text(text: str, force_chinese_mode: bool = False) -> Pt:
    """Get space before for text paragraphs"""
    is_chinese = detect_chinese_text(text, force_chinese_mode)
    return Pt(3) if is_chinese else Pt(2)


def replace_entity_placeholders(content: str, project_name: str) -> str:
    """Replace entity placeholders in content"""
    if not content or not project_name:
        return content

    # Replace common placeholders
    replacements = {
        '[PROJECT]': project_name,
        '[Entity]': project_name,
        '[Company]': project_name,
    }

    for placeholder, replacement in replacements.items():
        content = content.replace(placeholder, replacement)

    return content


class PowerPointGenerator:
    """Main PowerPoint generator class"""

    def __init__(self, template_path: str, language: str = 'english', row_limit: int = 20):
        self.template_path = template_path
        self.language = language.lower()
        self.row_limit = row_limit
        self.presentation = None

    def load_template(self):
        """Load the PowerPoint template"""
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found: {self.template_path}")

        self.presentation = Presentation(self.template_path)
        logger.info(f"Loaded template: {self.template_path}")

    def find_shape_by_name(self, shapes, name: str):
        """Find shape by name in slide"""
        for shape in shapes:
            if hasattr(shape, 'name') and shape.name == name:
                return shape
        return None
    
    def find_content_shape(self, shapes):
        """Find content shape by trying multiple possible names"""
        # Try different possible names for content shapes
        possible_names = [
            'Content',
            'Text-commentary',
            'textMainBullets',
            'Text',
            'Commentary',
            'MainContent',
            'Body'
        ]
        
        for name in possible_names:
            shape = self.find_shape_by_name(shapes, name)
            if shape and shape.has_text_frame:
                return shape
        
        # If no named shape found, try to find any text frame shape that's not a title
        for shape in shapes:
            if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
                shape_name = getattr(shape, 'name', '')
                # Skip title shapes and other non-content shapes
                if shape_name and 'title' not in shape_name.lower() and 'proj' not in shape_name.lower():
                    return shape
        
        return None

    def replace_text_preserve_formatting(self, shape, replacements: Dict[str, str]):
        """Replace text while preserving formatting"""
        if not shape.has_text_frame:
            return

        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                for old_text, new_text in replacements.items():
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)

    def update_project_titles(self, project_name: str, statement_type: str = 'BS'):
        """Update project titles in presentation"""
        if not self.presentation:
            return

        # Extract first two words for professional display
        if project_name:
            words = project_name.split()
            display_entity = ' '.join(words[:2]) if len(words) >= 2 else words[0] if words else project_name
        else:
            display_entity = project_name

        # Define title templates based on language and statement type
        if statement_type.upper() == 'BS':
            if self.language == 'chinese':
                title_template = f"资产负债表概览 - {display_entity}"
            else:
                title_template = f"Entity Overview - {display_entity}"
        elif statement_type.upper() == 'IS':
            if self.language == 'chinese':
                title_template = f"利润表概览 - {display_entity}"
            else:
                title_template = f"Income Statement - {display_entity}"
        else:
            if self.language == 'chinese':
                title_template = f"财务报表概览 - {display_entity}"
            else:
                title_template = f"Financial Report - {display_entity}"

        # Update titles in all slides
        for slide_index, slide in enumerate(self.presentation.slides):
            current_slide_number = slide_index + 1
            proj_title_shape = self.find_shape_by_name(slide.shapes, "projTitle")

            if proj_title_shape:
                current_text = proj_title_shape.text
                if "[PROJECT]" in current_text:
                    replacements = {
                        "[PROJECT]": display_entity,
                        "[Current]": str(current_slide_number),
                        "[Total]": str(len(self.presentation.slides))
                    }
                    self.replace_text_preserve_formatting(proj_title_shape, replacements)
                else:
                    # Replace the entire title
                    if proj_title_shape.has_text_frame:
                        proj_title_shape.text_frame.text = title_template

    def generate_full_report(self, markdown_content: str, summary_md: Optional[str] = None,
                           output_path: str = None):
        """Generate full PowerPoint report from markdown content"""
        if not self.presentation:
            self.load_template()

        # Process markdown content
        processed_content = self._process_markdown_content(markdown_content)

        # Apply content to presentation
        self._apply_content_to_presentation(processed_content)

        # Save if output path provided
        if output_path:
            self.save(output_path)

    def _process_markdown_content(self, content: str) -> Dict:
        """Process markdown content into structured data"""
        if not content:
            logger.warning("Empty content provided to _process_markdown_content")
            return {}

        logger.info(f"Processing markdown content, length: {len(content)}")
        logger.debug(f"Content preview (first 500 chars): {content[:500]}")

        # Split by headers (## Account Name)
        sections = re.split(r'^##\s+(.+)$', content, flags=re.MULTILINE)

        logger.info(f"Found {len(sections)} sections after splitting")

        processed_sections = {}

        # Process each section
        for i in range(1, len(sections), 2):
            if i + 1 < len(sections):
                account_name = sections[i].strip()
                account_content = sections[i + 1].strip()

                logger.info(f"Processing section: {account_name}, content length: {len(account_content)}")

                processed_sections[account_name] = {
                    'content': account_content,
                    'is_chinese': detect_chinese_text(account_content)
                }

        logger.info(f"Processed {len(processed_sections)} sections")
        return processed_sections

    def _apply_content_to_presentation(self, sections: Dict):
        """Apply processed content to presentation slides"""
        if not self.presentation:
            logger.warning("No presentation loaded")
            return

        logger.info(f"Applying {len(sections)} sections to presentation with {len(self.presentation.slides)} slides")

        # Find content placeholders and fill them
        slide_idx = 0
        for slide in self.presentation.slides:
            if slide_idx >= len(sections):
                logger.warning(f"More slides ({len(self.presentation.slides)}) than sections ({len(sections)})")
                break

            account_name = list(sections.keys())[slide_idx]
            section_data = sections[account_name]

            logger.info(f"Processing slide {slide_idx + 1} for account: {account_name}")

            # Find content shape using flexible name matching
            content_shape = self.find_content_shape(slide.shapes)
            if content_shape:
                logger.info(f"Found content shape '{content_shape.name}' on slide {slide_idx + 1}")
                if content_shape.has_text_frame:
                    # Apply content to shape
                    self._fill_content_shape(content_shape, section_data)
                    logger.info(f"Applied content to slide {slide_idx + 1}")
                else:
                    logger.warning(f"Content shape found but has no text_frame on slide {slide_idx + 1}")
            else:
                logger.warning(f"No content shape found on slide {slide_idx + 1}, available shapes: {[s.name if hasattr(s, 'name') else 'unnamed' for s in slide.shapes]}")
                # Try to use the first available text frame as fallback
                for shape in slide.shapes:
                    if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
                        shape_name = getattr(shape, 'name', 'unnamed')
                        if 'title' not in shape_name.lower() and 'proj' not in shape_name.lower():
                            logger.info(f"Using fallback shape '{shape_name}' on slide {slide_idx + 1}")
                            self._fill_content_shape(shape, section_data)
                            break

            slide_idx += 1

    def _fill_content_shape(self, shape, section_data: Dict):
        """Fill content shape with processed data"""
        if not shape.has_text_frame:
            logger.warning("Shape does not have text_frame")
            return

        content = section_data.get('content', '')
        is_chinese = section_data.get('is_chinese', False)

        logger.info(f"Filling shape with content length: {len(content)}")

        # Clear existing content
        shape.text_frame.clear()
        
        if not content or not content.strip():
            logger.warning("No content to fill")
            return
        
        # Split content into paragraphs if it contains newlines
        content_lines = content.split('\n')
        
        # Add content with proper formatting
        for idx, line in enumerate(content_lines):
            line = line.strip()
            if not line and idx > 0:
                # Skip empty lines except add a paragraph break
                continue
            
            if idx == 0:
                # Use first paragraph or create one
                if shape.text_frame.paragraphs:
                    p = shape.text_frame.paragraphs[0]
                else:
                    p = shape.text_frame.add_paragraph()
            else:
                p = shape.text_frame.add_paragraph()
            
            p.text = line
            
            # Apply formatting to runs
            for run in p.runs:
                run.font.size = get_font_size_for_text(line, force_chinese_mode=is_chinese)
                run.font.name = get_font_name_for_text(line)

            # Set paragraph formatting
            p.space_after = get_space_after_for_text(line, force_chinese_mode=is_chinese)
            p.space_before = get_space_before_for_text(line, force_chinese_mode=is_chinese)
            p.line_spacing = get_line_spacing_for_text(line, force_chinese_mode=is_chinese)
        
        logger.info(f"Successfully filled shape with {len([l for l in content_lines if l.strip()])} paragraphs")

    def apply_structured_data_to_slides(self, structured_data: List[Dict], start_slide: int, 
                                       project_name: str, statement_type: str):
        """Apply structured data directly to slides (slides 1-4 for BS, 5-8 for IS)"""
        if not self.presentation:
            self.load_template()
        
        logger.info(f"Applying {len(structured_data)} accounts to slides starting at {start_slide}")
        
        # Limit to 4 slides per statement type
        max_slides = min(len(structured_data), 4)
        slide_indices = list(range(start_slide - 1, start_slide - 1 + max_slides))  # Convert to 0-based
        
        for idx, account_data in enumerate(structured_data[:max_slides]):
            slide_idx = slide_indices[idx]
            if slide_idx >= len(self.presentation.slides):
                logger.warning(f"Slide index {slide_idx + 1} exceeds available slides ({len(self.presentation.slides)})")
                break
            
            slide = self.presentation.slides[slide_idx]
            account_name = account_data.get('account_name', '')
            financial_data = account_data.get('financial_data')
            commentary = account_data.get('commentary', '')
            summary = account_data.get('summary', '')
            
            logger.info(f"Processing slide {slide_idx + 1} for account: {account_name}")
            
            # Update projTitle
            proj_title_shape = self.find_shape_by_name(slide.shapes, "projTitle")
            if proj_title_shape and proj_title_shape.has_text_frame:
                if project_name:
                    proj_title_shape.text_frame.text = project_name
            
            # Fill Table Placeholder with financial data (only on first slide of each statement type)
            if idx == 0:  # Only fill table on first slide (slide 1 for BS, slide 5 for IS)
                # Try different table placeholder names
                table_shape = None
                for table_name in ["Table Placeholder", "Table Placeholder 2", "Table"]:
                    table_shape = self.find_shape_by_name(slide.shapes, table_name)
                    if table_shape:
                        logger.info(f"Found table shape '{table_name}' on slide {slide_idx + 1}")
                        break
                
                if table_shape and financial_data is not None:
                    self._fill_table_placeholder(table_shape, financial_data)
                elif financial_data is not None:
                    logger.warning(f"Table Placeholder not found on slide {slide_idx + 1}, available shapes: {[s.name if hasattr(s, 'name') else 'unnamed' for s in slide.shapes]}")
            
            # Fill coSummaryShape with summary
            summary_shape = self.find_shape_by_name(slide.shapes, "coSummaryShape")
            if summary_shape and summary_shape.has_text_frame:
                summary_shape.text_frame.clear()
                if summary:
                    p = summary_shape.text_frame.paragraphs[0] if summary_shape.text_frame.paragraphs else summary_shape.text_frame.add_paragraph()
                    p.text = summary
                    for run in p.runs:
                        run.font.size = get_font_size_for_text(summary)
                        run.font.name = get_font_name_for_text(summary)
            
            # Fill textMainBullets with commentary (AI output)
            bullets_shape = self.find_shape_by_name(slide.shapes, "textMainBullets")
            if bullets_shape and bullets_shape.has_text_frame:
                bullets_shape.text_frame.clear()
                if commentary:
                    # Split commentary into bullet points
                    commentary_lines = commentary.split('\n')
                    for line_idx, line in enumerate(commentary_lines):
                        line = line.strip()
                        if not line:
                            continue
                        if line_idx == 0:
                            p = bullets_shape.text_frame.paragraphs[0] if bullets_shape.text_frame.paragraphs else bullets_shape.text_frame.add_paragraph()
                        else:
                            p = bullets_shape.text_frame.add_paragraph()
                        p.text = line
                        p.level = 0  # Bullet level
                        for run in p.runs:
                            run.font.size = get_font_size_for_text(line, force_chinese_mode=is_chinese)
                            run.font.name = get_font_name_for_text(line)
                logger.info(f"Filled textMainBullets with commentary on slide {slide_idx + 1}")
            else:
                logger.warning(f"textMainBullets not found on slide {slide_idx + 1}, available shapes: {[s.name if hasattr(s, 'name') else 'unnamed' for s in slide.shapes]}")
    
    def _fill_table_placeholder(self, shape, df):
        """Fill table placeholder with DataFrame data"""
        try:
            # Check if shape has a table (Table Placeholder might be a table shape)
            if hasattr(shape, 'table') and shape.table:
                table = shape.table
                logger.info(f"Found table with {len(table.rows)} rows and {len(table.columns)} columns")
                
                # Fill table with DataFrame data
                max_rows = min(len(df) + 1, len(table.rows))  # +1 for header
                max_cols = min(len(df.columns), len(table.columns))
                
                # Fill header row
                if len(table.rows) > 0:
                    header_row = table.rows[0]
                    for col_idx, col_name in enumerate(df.columns[:max_cols]):
                        if col_idx < len(header_row.cells):
                            cell = header_row.cells[col_idx]
                            cell.text = str(col_name)
                            logger.debug(f"Filled header cell {col_idx}: {col_name}")
                
                # Fill data rows
                for row_idx in range(min(len(df), len(table.rows) - 1)):
                    table_row = table.rows[row_idx + 1]
                    df_row = df.iloc[row_idx]
                    for col_idx, col_name in enumerate(df.columns[:max_cols]):
                        if col_idx < len(table_row.cells):
                            cell = table_row.cells[col_idx]
                            value = df_row[col_name]
                            # Format numbers
                            if isinstance(value, (int, float)):
                                cell.text = f"{value:,.0f}" if value != 0 else "0"
                            else:
                                cell.text = str(value)
                    logger.debug(f"Filled table row {row_idx + 1}")
            else:
                # If no table, try to add text representation to text frame
                logger.info("Table Placeholder is not a table shape, using text representation")
                if shape.has_text_frame:
                    shape.text_frame.clear()
                    # Convert DataFrame to formatted text table
                    try:
                        # Limit rows for display
                        df_display = df.head(15)
                        text_table = df_display.to_string(index=False)
                        if len(df) > 15:
                            text_table += f"\n\n... and {len(df) - 15} more rows"
                    except:
                        text_table = str(df.head(15))
                    
                    p = shape.text_frame.paragraphs[0] if shape.text_frame.paragraphs else shape.text_frame.add_paragraph()
                    p.text = text_table
                    logger.info(f"Added text table representation ({len(text_table)} chars)")
        except Exception as e:
            logger.error(f"Could not fill table placeholder: {e}")
            import traceback
            logger.error(traceback.format_exc())
            # Fallback: add text representation
            if shape.has_text_frame:
                shape.text_frame.clear()
                text_repr = str(df.head(10))
                p = shape.text_frame.paragraphs[0] if shape.text_frame.paragraphs else shape.text_frame.add_paragraph()
                p.text = text_repr

    def embed_financial_tables(self, excel_path: str, sheet_name: str, project_name: str, language: str):
        """Embed financial tables: BS to page 1, IS to page 5"""
        try:
            import pandas as pd
            from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement
            
            logger.info(f"Embedding financial tables from {excel_path}, sheet: {sheet_name}")
            
            # Use the existing extraction function
            bs_is_results = extract_balance_sheet_and_income_statement(excel_path, sheet_name, debug=False)
            if not bs_is_results:
                logger.warning("No BS/IS data extracted")
                return
            
            # Extract BS and IS DataFrames from results
            # The structure should have 'balance_sheet' and 'income_statement' keys with DataFrames
            bs_df = bs_is_results.get('balance_sheet')
            is_df = bs_is_results.get('income_statement')
            
            logger.info(f"Extracted BS: {bs_df.shape if bs_df is not None else 'None'}, IS: {is_df.shape if is_df is not None else 'None'}")
            
            # Embed BS table to slide 0 (page 1)
            if bs_df is not None and not bs_df.empty and len(self.presentation.slides) > 0:
                slide_0 = self.presentation.slides[0]
                table_shape = self.find_shape_by_name(slide_0.shapes, "Table Placeholder")
                if not table_shape:
                    # Try alternative names
                    for name in ["Table Placeholder 2", "Table"]:
                        table_shape = self.find_shape_by_name(slide_0.shapes, name)
                        if table_shape:
                            break
                
                if table_shape:
                    logger.info(f"Found table shape on slide 1, embedding BS table ({bs_df.shape})")
                    self._fill_table_placeholder(table_shape, bs_df)
                else:
                    logger.warning(f"Table Placeholder not found on slide 1 for BS table")
            
            # Embed IS table to slide 4 (page 5, 0-indexed)
            if is_df is not None and not is_df.empty and len(self.presentation.slides) > 4:
                slide_4 = self.presentation.slides[4]
                table_shape = self.find_shape_by_name(slide_4.shapes, "Table Placeholder")
                if not table_shape:
                    # Try alternative names
                    for name in ["Table Placeholder 2", "Table"]:
                        table_shape = self.find_shape_by_name(slide_4.shapes, name)
                        if table_shape:
                            break
                
                if table_shape:
                    logger.info(f"Found table shape on slide 5, embedding IS table ({is_df.shape})")
                    self._fill_table_placeholder(table_shape, is_df)
                else:
                    logger.warning(f"Table Placeholder not found on slide 5 for IS table")
                    
        except Exception as e:
            logger.error(f"Error embedding financial tables: {e}")
            import traceback
            logger.error(traceback.format_exc())

    def save(self, output_path: str):
        """Save the presentation"""
        if not self.presentation:
            raise ValueError("No presentation loaded")

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        self.presentation.save(output_path)
        logger.info(f"Presentation saved to: {output_path}")


class ReportGenerator:
    """Report generator that orchestrates the PPTX creation"""

    def __init__(self, template_path: str, markdown_file: str, output_path: str,
                 project_name: Optional[str] = None, language: str = 'english', row_limit: int = 20):
        self.template_path = template_path
        self.markdown_file = markdown_file
        self.output_path = output_path
        self.project_name = project_name
        self.language = language
        self.row_limit = row_limit

    def generate(self):
        """Generate the report"""
        logger.info(f"Starting PPTX generation...")
        logger.info(f"Template: {self.template_path}")
        logger.info(f"Markdown: {self.markdown_file}")
        logger.info(f"Output: {self.output_path}")
        logger.info(f"Language: {self.language}")
        logger.info(f"Project: {self.project_name}")

        # Read markdown content
        with open(self.markdown_file, 'r', encoding='utf-8') as f:
            md_content = f.read()

        logger.info(f"Content length: {len(md_content)} characters")

        # Create PowerPoint generator
        generator = PowerPointGenerator(self.template_path, self.language, self.row_limit)

        try:
            # Generate the report
            generator.generate_full_report(md_content, None, self.output_path)

            # Update project titles if project name provided
            if self.project_name:
                generator.update_project_titles(self.project_name, 'BS')  # Default to BS

        except Exception as e:
            logger.error(f"Report generation failed: {e}")
            raise

        logger.info(f"✅ PPTX generation completed: {self.output_path}")


def export_pptx(template_path: str, markdown_path: str, output_path: str,
                project_name: Optional[str] = None, excel_file_path: Optional[str] = None,
                language: str = 'english', statement_type: str = 'BS', row_limit: int = 20):
    """
    Export PowerPoint presentation from markdown content

    Args:
        template_path: Path to PPTX template
        markdown_path: Path to markdown content file
        output_path: Output PPTX file path
        project_name: Project/entity name for titles
        excel_file_path: Optional Excel file for embedding data
        language: Language ('english' or 'chinese')
        statement_type: Statement type ('BS' or 'IS')
        row_limit: Maximum rows per slide
    """
    generator = ReportGenerator(template_path, markdown_path, output_path,
                              project_name, language, row_limit)
    generator.generate()

    if not os.path.exists(output_path):
        raise FileNotFoundError(f"PPTX file was not created at {output_path}")

    # Update project titles with correct statement type
    if project_name:
        temp_presentation = Presentation(output_path)
        pptx_gen = PowerPointGenerator(template_path, language, row_limit)
        pptx_gen.presentation = temp_presentation
        pptx_gen.update_project_titles(project_name, statement_type)
        temp_presentation.save(output_path)
        # Ensure presentation is properly closed
        del temp_presentation
        del pptx_gen

    # Note: Excel embedding functionality would need to be implemented
    # if excel_file_path and project_name:
    #     embed_excel_data_in_pptx(output_path, excel_file_path, sheet_name, project_name, statement_type=statement_type)

    logger.info(f"✅ PowerPoint presentation successfully exported to: {output_path}")
    return output_path


def export_pptx_from_structured_data_combined(template_path: str, bs_data: List[Dict], is_data: List[Dict], 
                                              output_path: str, project_name: Optional[str] = None, 
                                              language: str = 'english', temp_path: Optional[str] = None,
                                              selected_sheet: Optional[str] = None):
    """
    Export ONE combined PowerPoint presentation with both BS and IS
    
    Args:
        template_path: Path to PPTX template
        bs_data: List of BS account data dicts
        is_data: List of IS account data dicts
        output_path: Output PPTX file path
        project_name: Project/entity name for titles
        language: Language ('english' or 'chinese')
        temp_path: Path to Excel file for table embedding
        selected_sheet: Sheet name for table embedding
    """
    try:
        logger.info(f"Starting COMBINED PPTX generation...")
        logger.info(f"Template: {template_path}")
        logger.info(f"Output: {output_path}")
        logger.info(f"Language: {language}")
        logger.info(f"BS accounts: {len(bs_data)}, IS accounts: {len(is_data)}")

        # Load template
        generator = PowerPointGenerator(template_path, language, row_limit=20)
        generator.load_template()

        # Apply BS data to slides 1-4
        if bs_data:
            generator.apply_structured_data_to_slides(bs_data, 1, project_name, 'BS')
        
        # Apply IS data to slides 5-8
        if is_data:
            generator.apply_structured_data_to_slides(is_data, 5, project_name, 'IS')
        
        # Embed financial tables: BS to page 1, IS to page 5
        if temp_path and selected_sheet:
            generator.embed_financial_tables(temp_path, selected_sheet, project_name, language)
        
        # Save presentation
        generator.save(output_path)
        
        logger.info(f"✅ Combined PPTX generation completed: {output_path}")
        return output_path
        
    except Exception as e:
        logger.error(f"PPTX generation failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise


def export_pptx_from_structured_data(template_path: str, structured_data: List[Dict], output_path: str,
                                     project_name: Optional[str] = None, language: str = 'english',
                                     statement_type: str = 'BS', start_slide: int = 1):
    """
    Export PowerPoint presentation from structured data (not markdown)
    
    Args:
        template_path: Path to PPTX template
        structured_data: List of account data dicts with keys: account_name, financial_data, commentary, summary
        output_path: Output PPTX file path
        project_name: Project/entity name for titles
        language: Language ('english' or 'chinese')
        statement_type: Statement type ('BS' or 'IS')
        start_slide: Starting slide index (1-4 for BS, 5-8 for IS)
    """
    try:
        logger.info(f"Starting PPTX generation from structured data...")
        logger.info(f"Template: {template_path}")
        logger.info(f"Output: {output_path}")
        logger.info(f"Language: {language}")
        logger.info(f"Statement type: {statement_type}, Start slide: {start_slide}")
        logger.info(f"Accounts to process: {len(structured_data)}")

        # Load template
        generator = PowerPointGenerator(template_path, language, row_limit=20)
        generator.load_template()

        # Apply structured data to slides
        generator.apply_structured_data_to_slides(structured_data, start_slide, project_name, statement_type)

        # Save presentation
        generator.save(output_path)
        
        logger.info(f"✅ PPTX generation completed: {output_path}")
        return output_path
        
    except Exception as e:
        logger.error(f"PPTX generation failed: {e}")
        raise


def merge_presentations(bs_presentation_path: str, is_presentation_path: str, output_path: str):
    """
    Merge Balance Sheet and Income Statement presentations into a single presentation.

    Args:
        bs_presentation_path: Path to Balance Sheet presentation
        is_presentation_path: Path to Income Statement presentation
        output_path: Path for merged output presentation
    """
    try:
        logger.info("🔄 Starting presentation merge...")
        logger.info(f"   BS: {bs_presentation_path}")
        logger.info(f"   IS: {is_presentation_path}")

        # Load BS presentation as base
        merged_prs = Presentation(bs_presentation_path)

        # Load IS presentation
        is_prs = Presentation(is_presentation_path)

        # Copy all slides from IS to BS
        # Use XML-level copying for reliable slide duplication
        import xml.etree.ElementTree as ET
        from copy import deepcopy
        
        for slide_idx, slide in enumerate(is_prs.slides):
            try:
                # Get the slide layout
                slide_layout = slide.slide_layout
                
                # Create new slide with same layout
                new_slide = merged_prs.slides.add_slide(slide_layout)
                
                # Get XML elements
                source_slide_xml = slide._element
                target_slide_xml = new_slide._element
                
                # Remove placeholder shapes from new slide (from layout)
                # We'll replace them with actual content
                shapes_to_remove = list(new_slide.shapes)
                for shape in shapes_to_remove:
                    try:
                        sp_tree = target_slide_xml.get_or_add_spTree()
                        sp_tree.remove(shape._element)
                    except:
                        pass
                
                # Copy all shapes from source slide
                source_sp_tree = source_slide_xml.get_or_add_spTree()
                target_sp_tree = target_slide_xml.get_or_add_spTree()
                
                for shape_element in source_sp_tree:
                    # Deep copy the shape element
                    new_shape_element = deepcopy(shape_element)
                    # Add to target slide
                    target_sp_tree.append(new_shape_element)
                    
            except Exception as e:
                logger.error(f"Error copying slide {slide_idx}, using fallback method: {e}")
                # Fallback: simple text copy
                slide_layout = slide.slide_layout
                new_slide = merged_prs.slides.add_slide(slide_layout)
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for new_shape in new_slide.shapes:
                            if (hasattr(new_shape, 'name') and hasattr(shape, 'name') and
                                new_shape.name == shape.name and new_shape.has_text_frame):
                                new_shape.text_frame.text = shape.text_frame.text
                                break

        # Save merged presentation
        merged_prs.save(output_path)
        
        # Ensure presentation objects are properly closed
        del merged_prs
        del is_prs
        
        # Force garbage collection to ensure file handles are released
        import gc
        gc.collect()

        logger.info("✅ Presentation merge completed successfully")
    except Exception as e:
        logger.error(f"❌ Presentation merge failed: {e}")
        raise
