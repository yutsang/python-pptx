#!/usr/bin/env python3
"""
PowerPoint Generation Module for Financial Reports
Based on the backup methods but implemented fresh for the new system
"""

import os
import re
import logging
from typing import Dict, List, Optional, Tuple
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


def get_font_size_for_text(text: str, base_size: int = Pt(9), force_chinese_mode: bool = False) -> Pt:
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
            return {}

        # Split by headers (## Account Name)
        sections = re.split(r'^##\s+(.+)$', content, flags=re.MULTILINE)

        processed_sections = {}

        # Process each section
        for i in range(1, len(sections), 2):
            if i + 1 < len(sections):
                account_name = sections[i].strip()
                account_content = sections[i + 1].strip()

                processed_sections[account_name] = {
                    'content': account_content,
                    'is_chinese': detect_chinese_text(account_content)
                }

        return processed_sections

    def _apply_content_to_presentation(self, sections: Dict):
        """Apply processed content to presentation slides"""
        if not self.presentation:
            return

        # Find content placeholders and fill them
        slide_idx = 0
        for slide in self.presentation.slides:
            if slide_idx >= len(sections):
                break

            account_name = list(sections.keys())[slide_idx]
            section_data = sections[account_name]

            # Find content shape (usually named 'Content' or similar)
            content_shape = self.find_shape_by_name(slide.shapes, "Content")
            if content_shape and content_shape.has_text_frame:
                # Apply content to shape
                self._fill_content_shape(content_shape, section_data)

            slide_idx += 1

    def _fill_content_shape(self, shape, section_data: Dict):
        """Fill content shape with processed data"""
        if not shape.has_text_frame:
            return

        content = section_data.get('content', '')
        is_chinese = section_data.get('is_chinese', False)

        # Clear existing content
        shape.text_frame.text = ""

        # Add content with proper formatting
        p = shape.text_frame.paragraphs[0]
        p.text = content

        # Apply formatting
        for run in p.runs:
            run.font.size = get_font_size_for_text(content, force_chinese_mode=is_chinese)
            run.font.name = get_font_name_for_text(content)

        # Set paragraph formatting
        p.space_after = get_space_after_for_text(content, force_chinese_mode=is_chinese)
        p.space_before = get_space_before_for_text(content, force_chinese_mode=is_chinese)
        p.line_spacing = get_line_spacing_for_text(content, force_chinese_mode=is_chinese)

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
