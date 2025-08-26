import os
import textwrap
import logging
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR
from pptx.oxml.xmlchemy import OxmlElement
from dataclasses import dataclass
from typing import List, Tuple
from datetime import datetime
import re
from pptx.oxml.ns import qn  # Required import for XML namespace handling

logging.basicConfig(level=logging.INFO)

def get_tab_name(project_name):
    """Get tab name based on project name - more flexible approach"""
    try:
        # Hardcoded mappings for known entities
        if project_name == 'Haining':
            return "BSHN"
        elif project_name == 'Nanjing':
            return "BSNJ"
        elif project_name == 'Ningbo':
            return "BSNB"
        else:
            # Try to find a sheet that contains the project name
            return None
    except Exception:
        return None

def clean_content_quotes(content):
    """
    Remove outermost quotation marks from content while preserving legitimate internal quotes.
    Supports both straight quotes (") and curly quotes (" ").
    """
    if not content:
        return content
    
    # Handle straight quotes
    content = re.sub(r'^"([^"]*)"$', r'\1', content)
    
    # Handle curly quotes
    content = re.sub(r'^"([^"]*)"$', r'\1', content)
    content = re.sub(r'^"([^"]*)"$', r'\1', content)
    
    return content

@dataclass
class FinancialItem:
    accounting_type: str
    account_title: str
    descriptions: List[str]
    layer1_continued: bool = False
    layer2_continued: bool = False
    is_table: bool = False

class PowerPointGenerator:
    def __init__(self, template_path: str):
        self.prs = Presentation(template_path)
        self.current_slide_index = 0
        self.LINE_HEIGHT = Pt(12)
        # Calculate ROWS_PER_SECTION based on shape dimensions
        shape_height_inches = shape.height / 914400
        if shape_height_inches <= 4.2:  # New template with 4.13" height
            self.ROWS_PER_SECTION = 25  # Fewer rows for shorter shapes
        else:
            self.ROWS_PER_SECTION = 30  # Original value for taller shapes
        
        # Find textMainBullets shape more robustly
        slide = self.prs.slides[0]
        shape = None
        try:
            shape = next(s for s in slide.shapes if s.name == "textMainBullets")
        except StopIteration:
            # Try to find any text shape if textMainBullets is not found
            for s in slide.shapes:
                if hasattr(s, 'text_frame'):
                    shape = s
                    break
        
        if not shape:
            raise ValueError("No suitable text shape found in template")
            
        self.CHARS_PER_ROW = 50
        self.BULLET_CHAR = '■ '
        self.DARK_BLUE = RGBColor(0, 50, 150)
        self.DARK_GREY = RGBColor(169, 169, 169)
        self.prev_layer1 = None
        self.prev_layer2 = None

    def log_template_shapes(self):
        # Removed shape audit logging to reduce debug output
        pass

    def _calculate_max_rows(self):
        slide = self.prs.slides[0]
        # Find textMainBullets shape more robustly
        shape = None
        try:
            shape = next(s for s in slide.shapes if s.name == "textMainBullets")
        except StopIteration:
            # Try to find any text shape if textMainBullets is not found
            for s in slide.shapes:
                if hasattr(s, 'text_frame'):
                    shape = s
                    break
        
        if not shape:
            return 25  # Default fallback
        
        # Calculate available height in inches
        shape_height_inches = shape.height / 914400
        
        # For new template with 4.13" height, adjust calculation
        if shape_height_inches <= 4.2:  # New template dimensions
            # More conservative line calculation for shorter shapes
            # 4.13" height with 12pt line height = approximately 33 lines
            # Reduce by 6 to account for margins and spacing
            return int(shape.height / self.LINE_HEIGHT) - 6
        else:
            # Original calculation for taller shapes
            return int(shape.height / self.LINE_HEIGHT) - 4

    def parse_markdown(self, md_content: str) -> List[FinancialItem]:
        items = []
        current_type = ""
        current_title = ""
        current_descs = []
        is_table = False

        for line in md_content.strip().split('\n'):
            stripped = line.strip()
            
            if stripped.startswith('## '):
                if current_descs:
                    items.append(FinancialItem(
                        current_type, current_title, current_descs, 
                        is_table=is_table
                    ))
                current_type = stripped[3:]
                current_title = ""
                current_descs = []
                is_table = False
            elif stripped.startswith('### '):
                if current_descs:
                    items.append(FinancialItem(
                        current_type, current_title, current_descs,
                        is_table=is_table
                    ))
                current_title = stripped[4:]
                current_descs = []
                is_table = "taxes payables" in current_title.lower()
            elif re.match(r'^={20,}', stripped):
                is_table = True
                current_descs.append("="*40)
            else:
                if stripped:
                    # Clean quotes from the content
                    cleaned_stripped = clean_content_quotes(stripped)
                    current_descs.append(cleaned_stripped)

        if current_descs:
            items.append(FinancialItem(current_type, current_title, current_descs, is_table=is_table))
        return items

    def _plan_content_distribution(self, items: List[FinancialItem]):
        distribution = []
        content_queue = items.copy()
        slide_idx = 0
        
        while content_queue:
            if slide_idx == 0:
                # First slide: only center section
                sections = ['c']
            else:
                # Subsequent slides: left then right sections
                sections = ['b', 'c']
            
            for section in sections:
                if not content_queue:
                    break
                    
                section_items = []
                lines_used = 0
                
                # Fill the current section
                while content_queue and lines_used < self.ROWS_PER_SECTION:
                    item = content_queue[0]
                    item_lines = self._calculate_item_lines(item)
                    
                    if lines_used + item_lines <= self.ROWS_PER_SECTION:
                        # Item fits completely
                        section_items.append(item)
                        content_queue.pop(0)
                        lines_used += item_lines
                    else:
                        # Item needs to be split
                        remaining_lines = self.ROWS_PER_SECTION - lines_used
                        if remaining_lines > 1:
                            split_item, remaining_item = self._split_item(item, remaining_lines)
                            section_items.append(split_item)
                            
                            if remaining_item and remaining_item.descriptions and remaining_item.descriptions[0]:
                                remaining_item.layer2_continued = True
                                remaining_item.layer1_continued = True
                                content_queue[0] = remaining_item
                            else:
                                content_queue.pop(0)
                        else:
                            # Not enough space, move to next section
                            break
                
                # Add section to distribution if it has content
                if section_items:
                    distribution.append((slide_idx, section, section_items))
            
            slide_idx += 1
            if slide_idx >= len(self.prs.slides):
                break
                
        return distribution

    def _calculate_chars_per_line(self, shape):
        """Calculate characters per line based on actual shape dimensions and text frame margins"""
        if not shape:
            return self.CHARS_PER_ROW
        
        # Get the actual available text width (shape width minus margins)
        shape_width_inches = shape.width / 914400  # Convert EMU to inches
        
        # Get text frame margins if available
        margin_left = 0
        margin_right = 0
        if hasattr(shape, 'text_frame'):
            tf = shape.text_frame
            margin_left = tf.margin_left / 914400 if tf.margin_left else 0
            margin_right = tf.margin_right / 914400 if tf.margin_right else 0
        
        # Calculate available text width
        available_width_inches = shape_width_inches - margin_left - margin_right
        
        # For the new template with 4.78" width shapes, use more conservative calculation
        # The new shapes are narrower, so we need to adjust character width accordingly
        if available_width_inches <= 4.8:  # New template dimensions
            # More conservative character width for narrower shapes
            avg_char_px = 8.0  # Wider character estimate for better fit
        else:
            # Original calculation for wider shapes
            avg_char_px = 7.2
        
        # Convert to pixels (1 inch = 96 pixels)
        available_width_px = available_width_inches * 96
        
        # Calculate characters per line
        chars_per_line = int(available_width_px / avg_char_px)
        
        # Ensure reasonable bounds for new template
        chars_per_line = max(25, min(chars_per_line, 80))  # More conservative bounds
        
        # Apply special handling for first page (more conservative)
        if hasattr(self, 'current_slide_index') and self.current_slide_index == 0:
            # Reduce characters per line for first page to prevent overflow
            chars_per_line = min(chars_per_line, 70)
        
        return chars_per_line

    def _wrap_text_to_shape(self, text, shape):
        """Wrap text to fit the actual shape dimensions with better control"""
        chars_per_line = self._calculate_chars_per_line(shape)
        
        # Use more conservative wrapping for better control
        # Reduce chars per line slightly to account for variable character widths
        adjusted_chars = max(25, chars_per_line - 5)
        
        # Use textwrap with better settings
        wrapped = textwrap.wrap(
            text, 
            width=adjusted_chars,
            break_long_words=True,
            break_on_hyphens=False,
            replace_whitespace=False
        )
        
        return wrapped

    def _calculate_item_lines(self, item: FinancialItem) -> int:
        """Calculate lines needed for an item with more accurate text wrapping"""
        shape = getattr(self, 'current_shape', None)
        chars_per_line = self._calculate_chars_per_line(shape) if shape else self.CHARS_PER_ROW
        
        lines = 0
        
        # Header line
        header = f"{item.accounting_type} (continued)" if item.layer1_continued else item.accounting_type
        header_lines = len(textwrap.wrap(header, width=chars_per_line, break_long_words=True)) or 1
        lines += header_lines
        
        # Description lines
        desc_lines = 0
        for desc in item.descriptions:
            for para in desc.split('\n'):
                # Use the same wrapping logic as _wrap_text_to_shape
                adjusted_chars = max(25, chars_per_line - 5)
                para_lines = len(textwrap.wrap(
                    para, 
                    width=adjusted_chars,
                    break_long_words=True,
                    break_on_hyphens=False,
                    replace_whitespace=False
                )) or 1
                desc_lines += para_lines
        
        lines += desc_lines
        return lines

    def _split_item(self, item: FinancialItem, max_lines: int) -> tuple[FinancialItem, FinancialItem | None]:
        """Split item at paragraph boundaries with improved text wrapping"""
        shape = getattr(self, 'current_shape', None)
        chars_per_line = self._calculate_chars_per_line(shape) if shape else self.CHARS_PER_ROW
        header = f"{self.BULLET_CHAR}{item.account_title} - "
        desc_paras = item.descriptions
        lines_used = 0
        split_index = 0
        
        # Try to fit as many whole paragraphs as possible
        for i, para in enumerate(desc_paras):
            # Use the same wrapping logic as _wrap_text_to_shape
            adjusted_chars = max(25, chars_per_line - 5)
            para_lines = len(textwrap.wrap(
                para, 
                width=adjusted_chars,
                break_long_words=True,
                break_on_hyphens=False,
                replace_whitespace=False
            )) or 1
            if lines_used + para_lines > max_lines:
                break
            lines_used += para_lines
            split_index = i + 1
        # If all paragraphs fit, no split needed
        if split_index == len(desc_paras):
            return (
                FinancialItem(
                    item.accounting_type,
                    item.account_title,
                    desc_paras,
                    layer1_continued=item.layer1_continued,
                    layer2_continued=item.layer2_continued,
                    is_table=item.is_table
                ),
                None
            )
        # If no paragraph fits, split the first paragraph at line boundary
        if split_index == 0:
            para = desc_paras[0]
            adjusted_chars = max(25, chars_per_line - 5)
            wrapped = textwrap.wrap(
                para, 
                width=adjusted_chars,
                break_long_words=True,
                break_on_hyphens=False,
                replace_whitespace=False
            )
            first_part = wrapped[:max_lines]
            remaining_part = wrapped[max_lines:]
            split_item = FinancialItem(
                item.accounting_type,
                item.account_title,
                [' '.join(first_part)],
                layer1_continued=item.layer1_continued,
                layer2_continued=item.layer2_continued,
                is_table=item.is_table
            )
            remaining_item = FinancialItem(
                item.accounting_type,
                item.account_title,
                [' '.join(remaining_part)] + desc_paras[1:],
                layer1_continued=True,
                layer2_continued=True,
                is_table=item.is_table
            ) if remaining_part or len(desc_paras) > 1 else None
            return split_item, remaining_item
        # Otherwise, split at paragraph boundary
        split_item = FinancialItem(
            item.accounting_type,
            item.account_title,
            desc_paras[:split_index],
            layer1_continued=item.layer1_continued,
            layer2_continued=item.layer2_continued,
            is_table=item.is_table
        )
        remaining_item = FinancialItem(
            item.accounting_type,
            item.account_title,
            desc_paras[split_index:],
            layer1_continued=True,
            layer2_continued=True,
            is_table=item.is_table
        ) if split_index < len(desc_paras) else None
        return split_item, remaining_item

    def _format_table_text(self, text: str) -> str:
        lines = []
        current_line = []
        for part in text.split():
            if part == self.BULLET_CHAR.strip():
                if current_line:
                    lines.append(" ".join(current_line))
                current_line = ["  " + part]
            else:
                current_line.append(part)
        lines.append(" ".join(current_line))
        return "\n".join(lines)

    def _validate_content_placement(self, distribution):
        seen = set()
        for slide_idx, section, items in distribution:
            for item in items:
                key = (item.accounting_type, item.account_title, tuple(item.descriptions))
                if key in seen:
                    raise ValueError(f"Duplicate content detected: {key}")
                seen.add(key)

    def parse_summary_markdown(self, md_content: str) -> str:
        summary_lines = []
        in_summary = False
        for line in md_content.strip().split('\n'):
            stripped = line.strip()
            if stripped.startswith('## Summary'):
                in_summary = True
                continue
            if in_summary:
                if stripped.startswith('##'):
                    break  # Next section found
                if stripped:  # Skip empty lines
                    summary_lines.append(stripped)
        summary_text = " ".join(summary_lines)
        return summary_text

    def _get_section_shape(self, slide, section: str):
        if self.current_slide_index == 0:
            if section == 'c':
                shape = next((s for s in slide.shapes if s.name == "textMainBullets"), None)
                # Apply special handling for first page
                if shape:
                    self._optimize_first_page_shape(shape)
                return shape
            return None  # No 'b' section on first slide
        # For subsequent slides
        target_name = "textMainBullets_L" if section == 'b' else "textMainBullets_R"
        return next((s for s in slide.shapes if s.name == target_name), None)

    def _populate_section(self, shape, items: List[FinancialItem]):
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        
        # Optimize text frame for better control
        self._optimize_text_frame(tf, shape)
        
        self.prev_layer1 = None  # Reset for each new section
        self.current_shape = shape  # For dynamic line calculation
        prev_account_title = None
        prev_continued = False
        paragraph_count = 0  # Track paragraph index in the shape
        for idx, item in enumerate(items):
            is_first_part = not (item.layer1_continued or item.layer2_continued)
            # Layer 1 Header
            if (item.accounting_type != self.prev_layer1) or (self.prev_layer1 is None):
                p = tf.add_paragraph()
                self._apply_paragraph_formatting(p, is_layer2_3=False)
                if paragraph_count > 0:
                    try: p.space_before = Pt(3)
                    except: pass
                run = p.add_run()
                cont_text = " (continued)" if item.layer1_continued else ""
                run.text = f"{item.accounting_type}{cont_text}"
                run.font.size = Pt(9)
                run.font.bold = True
                run.font.name = 'Arial'
                run.font.color.rgb = self.DARK_BLUE
                self.prev_layer1 = item.accounting_type
                paragraph_count += 1
            # Treat all items the same (no special handling for taxes payables)
            for para_idx, desc in enumerate(item.descriptions):
                # Split on '\n' and add an empty row between each part
                desc_parts = desc.split('\n')
                for part_idx, part in enumerate(desc_parts):
                    if is_first_part and para_idx == 0 and part_idx == 0:
                        # Main bullet with heading
                        p = tf.add_paragraph()
                        self._apply_paragraph_formatting(p, is_layer2_3=True)
                        if paragraph_count > 0:
                            try: p.space_before = Pt(3)
                            except: pass
                        bullet_run = p.add_run()
                        bullet_run.text = self.BULLET_CHAR
                        bullet_run.font.color.rgb = self.DARK_GREY
                        bullet_run.font.name = 'Arial'
                        bullet_run.font.size = Pt(9)
                        title_run = p.add_run()
                        cont_text = " (continued)" if item.layer2_continued else ""
                        title_run.text = f"{item.account_title}{cont_text}"
                        title_run.font.bold = True
                        title_run.font.name = 'Arial'
                        title_run.font.size = Pt(9)
                        desc_run = p.add_run()
                        desc_run.text = f" - {part}"
                        desc_run.font.bold = False
                        desc_run.font.name = 'Arial'
                        desc_run.font.size = Pt(9)
                        paragraph_count += 1
                    else:
                        # Continuation: visually subordinate (indented, no bullet)
                        p = tf.add_paragraph()
                        self._apply_paragraph_formatting(p, is_layer2_3=True)
                        if paragraph_count > 0:
                            try: p.space_before = Pt(3)
                            except: pass
                        try: p.left_indent = Inches(0.4)
                        except: pass
                        try: p.first_line_indent = Inches(-0.19)
                        except: pass
                        cont_run = p.add_run()
                        cont_run.text = part
                        cont_run.font.bold = False
                        cont_run.font.name = 'Arial'
                        cont_run.font.size = Pt(9)
                        paragraph_count += 1
                    # Add an empty row between split parts (but not after the last)
                    if part_idx < len(desc_parts) - 1:
                        tf.add_paragraph()
                        paragraph_count += 1
            prev_account_title = item.account_title
            prev_continued = item.layer1_continued or item.layer2_continued
        self.current_shape = None

    def _optimize_text_frame(self, text_frame, shape):
        """Optimize text frame properties for better text control"""
        try:
            # Set auto-size to none for better control
            text_frame.auto_size = MSO_AUTO_SIZE.NONE
            
            # Set vertical anchor to top
            text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
            
            # Adjust margins for better text control based on shape dimensions
            shape_width_inches = shape.width / 914400
            shape_height_inches = shape.height / 914400
            
            # For new template with 4.78" width and 4.13" height
            if shape_width_inches <= 4.8 and shape_height_inches <= 4.2:
                # New template: use smaller margins for better text utilization
                margin_left = Inches(0.03)
                margin_right = Inches(0.03)
                margin_top = Inches(0.03)
                margin_bottom = Inches(0.03)
            elif shape_width_inches > 8:  # Wide template
                margin_left = Inches(0.05)
                margin_right = Inches(0.05)
                margin_top = Inches(0.05)
                margin_bottom = Inches(0.05)
            else:  # Standard template
                margin_left = Inches(0.1)
                margin_right = Inches(0.1)
                margin_top = Inches(0.05)
                margin_bottom = Inches(0.05)
            
            text_frame.margin_left = margin_left
            text_frame.margin_right = margin_right
            text_frame.margin_top = margin_top
            text_frame.margin_bottom = margin_bottom
            
        except Exception as e:
            # If optimization fails, continue with default settings
            pass

    def _optimize_first_page_shape(self, shape):
        """Apply special optimizations for the first page shape"""
        try:
            # For the first page, use more conservative text wrapping
            # This helps prevent text from being too wide and uncontrollable
            
            # Adjust the shape's text frame for better control
            if hasattr(shape, 'text_frame'):
                tf = shape.text_frame
                
                # Get shape dimensions to determine appropriate margins
                shape_width_inches = shape.width / 914400
                shape_height_inches = shape.height / 914400
                
                # For new template with 4.78" width and 4.13" height
                if shape_width_inches <= 4.8 and shape_height_inches <= 4.2:
                    # New template: use very small margins for maximum text utilization
                    tf.margin_left = Inches(0.02)
                    tf.margin_right = Inches(0.02)
                    tf.margin_top = Inches(0.02)
                    tf.margin_bottom = Inches(0.02)
                else:
                    # Original template: use standard margins
                    tf.margin_left = Inches(0.08)
                    tf.margin_right = Inches(0.08)
                    tf.margin_top = Inches(0.05)
                    tf.margin_bottom = Inches(0.05)
                
                # Ensure word wrap is enabled
                tf.word_wrap = True
                
                # Set auto-size to none for better control
                tf.auto_size = MSO_AUTO_SIZE.NONE
                
                # Set vertical anchor to top
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
                
        except Exception as e:
            # If optimization fails, continue with default settings
            pass

    def _detect_bullet_content(self, text):
        """
        Enhanced bullet detection that handles multi-level bullet formats up to 4 levels deep
        with special handling for layer 3 bullet points that might contain dashes indicating layer 4
        """
        lines = text.split('\n')
        bullet_lines = []
        has_bullets = False

        # Check if we have any lines starting with "- "
        for line in lines:
            if line.strip().startswith('- '):
                has_bullets = True
                break
        
        # If we have bullets, process the content
        if has_bullets:
            main_content = []
            
            for line in lines:
                stripped = line.strip()
                original_line = line
                
                # Detect bullet lines with '- ' prefix
                if original_line.lstrip().startswith('- '):
                    # If we have accumulated main content, add it first
                    if main_content:
                        main_text = ' '.join(main_content).strip()
                        if main_text:
                            bullet_lines.append((0, main_text))
                        main_content = []
                    
                    # Calculate indentation level (based on spaces/tabs before the bullet)
                    indent_spaces = len(original_line) - len(original_line.lstrip())
                    
                    # Determine bullet level based on indentation
                    # We divide by 2 assuming 2 spaces per indentation level
                    level = min(4, (indent_spaces // 2) + 1)  # Cap at 4 levels
                    
                    # Clean and store bullet line
                    clean_line = stripped[2:]  # Remove '- '
                    
                    # Special handling for level 3 bullets that contain a dash indicating level 4
                    if level == 3 and " - " in clean_line:
                        # Split at the first occurrence of " - "
                        parts = clean_line.split(" - ", 1)
                        if len(parts) > 1:
                            # Add level 3 content
                            bullet_lines.append((level, parts[0].strip()))
                            # Add level 4 content
                            bullet_lines.append((4, parts[1].strip()))
                        else:
                            bullet_lines.append((level, clean_line))
                    else:
                        bullet_lines.append((level, clean_line))
                
                elif stripped:
                    # This is regular content - could be main content or continuation
                    main_content.append(stripped)
            
            # Add any remaining main content
            if main_content:
                main_text = ' '.join(main_content).strip()
                if main_text:
                    bullet_lines.insert(0, (0, main_text))
        
        return bullet_lines if bullet_lines else None

    def generate_full_report(self, md_content: str, summary_md: str, output_path: str):
        try:
            items = self.parse_markdown(md_content)
            distribution = self._plan_content_distribution(items)
            self._validate_content_placement(distribution)
            
            # Remove summary processing - no longer needed
            max_slide_used = max((slide_idx for slide_idx, _, _ in distribution), default=0)
            total_slides_needed = max_slide_used + 1
            
            while len(self.prs.slides) < total_slides_needed:
                self.prs.slides.add_slide(self.prs.slide_layouts[1])
            for slide_idx, section, section_items in distribution:
                if slide_idx >= len(self.prs.slides):
                    raise ValueError("Insufficient slides in template")
                slide = self.prs.slides[slide_idx]
                self.current_slide_index = slide_idx
                shape = self._get_section_shape(slide, section)
                if shape:
                    self._populate_section(shape, section_items)
            # Summary placement removed - no longer needed
            pass
            unused_slides = self._detect_unused_slides(distribution)
            if unused_slides:
                logging.info(f"Removing unused slides: {[idx+1 for idx in unused_slides]}")
                self._remove_slides(unused_slides)
            self.prs.save(output_path)
        except Exception as e:
            logging.error(f"Generation failed: {str(e)}")
            raise

    def _apply_paragraph_formatting(self, paragraph, is_layer2_3=False):
        # Only use legacy assignments, never .paragraph_format
        try:
            if is_layer2_3:
                try: paragraph.left_indent = Inches(0.21)
                except: pass
                try: paragraph.first_line_indent = Inches(-0.19)
                except: pass
                try: paragraph.space_before = Pt(0)
                except: pass
            else:
                try: paragraph.left_indent = Inches(0.3)
                except: pass
                try: paragraph.first_line_indent = Inches(-0.3)
                except: pass
                try: paragraph.space_before = Pt(0)
                except: pass
            try: paragraph.space_after = Pt(0)
            except: pass
            try: paragraph.line_spacing = 1.0
            except: pass
            try: paragraph.alignment = PP_ALIGN.LEFT
            except: self._handle_legacy_alignment(paragraph)
        except Exception as e:
            print(f"[ERROR] _apply_paragraph_formatting failed: {e} (type: {type(paragraph)})")

    def _handle_legacy_alignment(self, paragraph):
        """XML-based left alignment for legacy versions"""
        try:
            pPr = paragraph._element.get_or_add_pPr()
            for child in list(pPr):
                if child.tag.endswith('jc'):
                    pPr.remove(child)
            align = OxmlElement('a:jc')
            align.set('val', 'left')
            pPr.append(align)
        except Exception as e:
            logging.warning(f"Legacy alignment failed: {str(e)}")

    def _create_taxes_table(self, shape, item):
        tf = shape.text_frame
        # Header
        p = tf.add_paragraph()
        self._apply_paragraph_formatting(p, is_layer2_3=True)
        header_run = p.add_run()
        cont_text = " (continued)" if item.layer2_continued else ""
        header_run.text = f"{self.BULLET_CHAR}{item.account_title}{cont_text}"
        header_run.font.bold = True
        header_run.font.name = 'Arial'
        header_run.font.size = Pt(9)
        # Process table content
        content = ' '.join(item.descriptions)
        current_category = None
        for line in content.split('\n'):
            line = line.strip()
            if not line or '===' in line:
                continue
            if 'Tax:' in line or 'Payable:' in line:
                current_category = line.replace(':', '')
                p = tf.add_paragraph()
                self._apply_paragraph_formatting(p, is_layer2_3=True)
                try: p.left_indent = Inches(0.4)
                except: pass
                run = p.add_run()
                run.text = f"• {current_category}"
                run.font.name = 'Arial'
                run.font.bold = True
                run.font.size = Pt(9)
            elif line.startswith('- '):
                p = tf.add_paragraph()
                self._apply_paragraph_formatting(p, is_layer2_3=True)
                try: p.left_indent = Inches(0.6)
                except: pass
                run = p.add_run()
                run.text = f"◦ {line[2:]}"
                run.font.name = 'Arial'
                run.font.size = Pt(9)

    def _populate_summary_section_safe(self, shape, chunks: List[str]):
        """Summary section with updated formatting"""
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = Inches(0.07)
        tf.margin_right = Inches(0.07)
        p = tf.add_paragraph()
        full_text = " ".join(chunks)
        run = p.add_run()
        run.text = full_text
        run.font.size = Pt(9)  # All text pt 9
        run.font.bold = True
        run.font.name = 'Arial'
        run.font.color.rgb = RGBColor(255, 255, 255)
        # Set LEFT alignment
        try:
            p.alignment = PP_ALIGN.LEFT
        except AttributeError:
            try:
                p.alignment = PP_ALIGN.LEFT
            except AttributeError:
                self._handle_legacy_alignment(p)
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP

    def _detect_unused_slides(self, distribution):
        """Adjusted slide retention logic with content-aware detection"""
        used_slides = set()
        content_slides = set()
        # Track slides that actually contain content
        for slide_idx, section, items in distribution:
            if items:  # Only consider slides with actual content
                content_slides.add(slide_idx)
            used_slides.add(slide_idx)
        # Calculate maximum content-bearing slide
        max_content_slide = max(content_slides) if content_slides else 0
        # Determine minimum slides to keep (content slides + buffer)
        min_slides_to_keep = max(2, max_content_slide + 1)
        # Preserve all slides up to min_slides_to_keep
        remove_slides = []
        for slide_idx in range(len(self.prs.slides)-1, min_slides_to_keep-1, -1):
            remove_slides.append(slide_idx)
        return sorted(remove_slides, reverse=True)

    def _remove_slides(self, slide_indices):
        """Remove slides by indices (must be in descending order)"""
        for slide_idx in slide_indices:
            if slide_idx < len(self.prs.slides):
                # Method 1: Using _sldIdLst (more reliable)
                xml_slides = self.prs.slides._sldIdLst
                slides = list(xml_slides)
                # Remove relationship
                rId = slides[slide_idx].rId
                self.prs.part.drop_rel(rId)
                # Remove from slide list
                xml_slides.remove(slides[slide_idx])

    def _populate_bullet_content(self, shape, item, bullet_lines):
        tf = shape.text_frame
        p = tf.add_paragraph()
        self._apply_paragraph_formatting(p, is_layer2_3=True)
        bullet_run = p.add_run()
        bullet_run.text = self.BULLET_CHAR
        bullet_run.font.color.rgb = self.DARK_GREY
        bullet_run.font.name = 'Arial'
        bullet_run.font.size = Pt(9)
        title_run = p.add_run()
        cont_text = " (continued)" if item.layer2_continued else ""
        title_run.text = f"{item.account_title}{cont_text}"
        title_run.font.bold = True
        title_run.font.name = 'Arial'
        title_run.font.size = Pt(9)
        for line in bullet_lines:
            p = tf.add_paragraph()
            self._apply_paragraph_formatting(p, is_layer2_3=True)
            try: p.left_indent = Inches(0.4)
            except: pass
            try: p.first_line_indent = Inches(-0.19)
            except: pass
            run = p.add_run()
            clean_line = line.lstrip('- ').strip() if isinstance(line, str) else line[1].lstrip('- ').strip()
            run.text = f"• {clean_line}"
            run.font.name = 'Arial'
            run.font.size = Pt(9)
            run.font.bold = False

class ReportGenerator:
    def __init__(self, template_path, markdown_file, output_path, project_name=None):
        self.template_path = template_path
        self.markdown_file = markdown_file
        self.output_path = output_path
        self.project_name = project_name
        
    def generate(self):
        # Remove summary content - no longer needed
        summary_content = ""
        with open(self.markdown_file, 'r', encoding='utf-8') as f:
            md_content = f.read()
        generator = PowerPointGenerator(self.template_path)
        try: 
            generator.generate_full_report(md_content, summary_content, self.output_path)
        except Exception as e:
            print(f"Generation failed: {str(e)}")

# --- Project Title Update Logic (from 3.wrap_up.py) ---
def find_shape_by_name(shapes, name):
    for shape in shapes:
        if shape.name == name:
            return shape
    return None

def replace_text_preserve_formatting(shape, replacements):
    if not shape.has_text_frame:
        return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            for old_text, new_text in replacements.items():
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)

def update_project_titles(presentation_path, project_name, output_path=None):
    prs = Presentation(presentation_path)
    total_slides = len(prs.slides)
    
    # Extract base entity name (e.g., "Haining" from "Haining Wanpu Limited")
    base_entity = project_name.split()[0] if project_name else project_name
    
    for slide_index, slide in enumerate(prs.slides):
        current_slide_number = slide_index + 1
        projTitle_shape = find_shape_by_name(slide.shapes, "projTitle")
        if projTitle_shape:
            replacements = {
                "[PROJECT]": base_entity,  # Use base entity name instead of full project name
                "[Current]": str(current_slide_number),
                "[Total]": str(total_slides)
            }
            replace_text_preserve_formatting(projTitle_shape, replacements)
    # Save the presentation
    if output_path is None:
        output_path = presentation_path
    prs.save(output_path)
    return output_path

# --- High-level function for app.py ---
def embed_excel_data_in_pptx(presentation_path, excel_file_path, sheet_name, project_name, output_path=None):
    """
    Update existing financialData shape in PowerPoint with Excel data
    """
    try:
        from pptx import Presentation
        from pptx.util import Inches
        import pandas as pd
        import os
        
        # Load the presentation
        prs = Presentation(presentation_path)
        
        # Read Excel data
        if sheet_name:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        else:
            # If no sheet name provided, read the first sheet
            df = pd.read_excel(excel_file_path, sheet_name=0)
        
        # Find the financialData shape in all slides
        financial_data_shape = None
        target_slide = None
        
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.name == "financialData":
                    financial_data_shape = shape
                    target_slide = slide
                    break
            if financial_data_shape:
                break
        
        if financial_data_shape:
            # Found the financialData shape - update its content
            if hasattr(financial_data_shape, 'table'):
                # It's a table shape - update the table data while preserving formatting
                table = financial_data_shape.table
                
                # Store original formatting for each cell
                original_formats = {}
                for row_idx in range(len(table.rows)):
                    for col_idx in range(len(table.columns)):
                        cell = table.cell(row_idx, col_idx)
                        # Store font properties
                        if cell.text_frame.paragraphs[0].runs:
                            run = cell.text_frame.paragraphs[0].runs[0]
                            original_formats[(row_idx, col_idx)] = {
                                'font_name': run.font.name,
                                'font_size': run.font.size,
                                'font_bold': run.font.bold,
                                'font_italic': run.font.italic,
                                'font_color': run.font.color.rgb if run.font.color.rgb else None
                            }
                
                # Clear existing content but preserve formatting
                for row in range(len(table.rows)):
                    for col in range(len(table.columns)):
                        if row < len(table.rows) and col < len(table.columns):
                            cell = table.cell(row, col)
                            cell.text = ""
                
                # Update with new data and apply original formatting
                # Header row
                for col_idx, col_name in enumerate(df.columns):
                    if col_idx < len(table.columns):
                        cell = table.cell(0, col_idx)
                        cell.text = str(col_name)
                        # Apply header formatting (usually bold, different color)
                        if (0, col_idx) in original_formats:
                            format_info = original_formats[(0, col_idx)]
                            if cell.text_frame.paragraphs[0].runs:
                                run = cell.text_frame.paragraphs[0].runs[0]
                                if format_info['font_name']:
                                    run.font.name = format_info['font_name']
                                if format_info['font_size']:
                                    run.font.size = format_info['font_size']
                                if format_info['font_bold'] is not None:
                                    run.font.bold = format_info['font_bold']
                                if format_info['font_italic'] is not None:
                                    run.font.italic = format_info['font_italic']
                                if format_info['font_color']:
                                    run.font.color.rgb = format_info['font_color']
                
                # Data rows
                for row_idx, row in enumerate(df.values):
                    if row_idx + 1 < len(table.rows):
                        for col_idx, value in enumerate(row):
                            if col_idx < len(table.columns):
                                cell = table.cell(row_idx + 1, col_idx)
                                cell.text = str(value)
                                # Apply data row formatting
                                if (row_idx + 1, col_idx) in original_formats:
                                    format_info = original_formats[(row_idx + 1, col_idx)]
                                    if cell.text_frame.paragraphs[0].runs:
                                        run = cell.text_frame.paragraphs[0].runs[0]
                                        if format_info['font_name']:
                                            run.font.name = format_info['font_name']
                                        if format_info['font_size']:
                                            run.font.size = format_info['font_size']
                                        if format_info['font_bold'] is not None:
                                            run.font.bold = format_info['font_bold']
                                        if format_info['font_italic'] is not None:
                                            run.font.italic = format_info['font_italic']
                                        if format_info['font_color']:
                                            run.font.color.rgb = format_info['font_color']
                
                logging.info(f"✅ Updated financialData table with Excel data (formatting preserved)")
                
            elif hasattr(financial_data_shape, 'text_frame'):
                # It's a text shape - convert to table or update text
                # For now, let's create a table in its place
                left = financial_data_shape.left
                top = financial_data_shape.top
                width = financial_data_shape.width
                height = financial_data_shape.height
                
                # Remove the old shape
                target_slide.shapes._spTree.remove(financial_data_shape._element)
                
                # Add new table with same position and size
                table = target_slide.shapes.add_table(
                    rows=len(df) + 1,  # +1 for header
                    cols=len(df.columns),
                    left=left,
                    top=top,
                    width=width,
                    height=height
                ).table
                
                # Name the new table
                table._element.getparent().getparent().set('name', 'financialData')
                
                # Fill table with data
                for col_idx, col_name in enumerate(df.columns):
                    table.cell(0, col_idx).text = str(col_name)
                
                for row_idx, row in enumerate(df.values):
                    for col_idx, value in enumerate(row):
                        table.cell(row_idx + 1, col_idx).text = str(value)
                
                logging.info(f"✅ Replaced financialData text with table")
                
        else:
            # financialData shape not found - create a new table in the middle
            logging.warning("⚠️ financialData shape not found, creating new table")
            
            if len(prs.slides) > 0:
                slide = prs.slides[0]
                
                # Calculate position for the table (middle of slide)
                slide_width = prs.slide_width
                slide_height = prs.slide_height
                object_width = Inches(8)
                object_height = Inches(4)
                left = (slide_width - object_width) / 2
                top = (slide_height - object_height) / 2
                
                # Add a table with the Excel data
                table = slide.shapes.add_table(
                    rows=len(df) + 1,  # +1 for header
                    cols=len(df.columns),
                    left=left,
                    top=top,
                    width=object_width,
                    height=object_height
                ).table
                
                # Name the table
                table._element.getparent().getparent().set('name', 'financialData')
                
                # Apply professional formatting to the new table
                from pptx.dml.color import RGBColor
                
                # Fill table with data and apply formatting
                for col_idx, col_name in enumerate(df.columns):
                    cell = table.cell(0, col_idx)
                    cell.text = str(col_name)
                    # Header formatting
                    if cell.text_frame.paragraphs[0].runs:
                        run = cell.text_frame.paragraphs[0].runs[0]
                        run.font.bold = True
                        run.font.size = Pt(12)
                        run.font.name = 'Arial'
                        # Header background color (light blue)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(217, 225, 242)
                
                for row_idx, row in enumerate(df.values):
                    for col_idx, value in enumerate(row):
                        cell = table.cell(row_idx + 1, col_idx)
                        cell.text = str(value)
                        # Data row formatting
                        if cell.text_frame.paragraphs[0].runs:
                            run = cell.text_frame.paragraphs[0].runs[0]
                            run.font.size = Pt(10)
                            run.font.name = 'Arial'
                            # Alternate row colors for readability
                            if row_idx % 2 == 0:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = RGBColor(242, 242, 242)
        
        # Save the presentation
        if output_path is None:
            output_path = presentation_path
        prs.save(output_path)
        
        logging.info(f"✅ Excel data updated in PowerPoint: {output_path}")
        return output_path
        
    except Exception as e:
        logging.error(f"❌ Failed to update Excel data: {str(e)}")
        raise

def export_pptx(template_path, markdown_path, output_path, project_name=None, excel_file_path=None):
    generator = ReportGenerator(template_path, markdown_path, output_path, project_name)
    generator.generate()
    if not os.path.exists(output_path):
        logging.error(f"PPTX file was not created at {output_path}")
        raise FileNotFoundError(f"PPTX file was not created at {output_path}")
    if project_name:
        update_project_titles(output_path, project_name)
    
    # Embed Excel data if provided
    if excel_file_path and project_name:
        # Get the appropriate sheet name based on project
        sheet_name = get_tab_name(project_name)
        if sheet_name:
            try:
                embed_excel_data_in_pptx(output_path, excel_file_path, sheet_name, project_name)
            except Exception as e:
                logging.warning(f"⚠️ Could not embed Excel data: {str(e)}")
        else:
            logging.warning(f"⚠️ No valid sheet name found for project: {project_name}")
            # Try to embed with first sheet as fallback
            try:
                embed_excel_data_in_pptx(output_path, excel_file_path, None, project_name)
            except Exception as e:
                logging.warning(f"⚠️ Could not embed Excel data with fallback: {str(e)}")
    
    # Add success message
    logging.info(f"✅ PowerPoint presentation successfully exported to: {output_path}")
    return output_path

def merge_presentations(bs_presentation_path, is_presentation_path, output_path):
    """
    Merge Balance Sheet and Income Statement presentations into a single presentation.
    
    Args:
        bs_presentation_path: Path to the Balance Sheet presentation
        is_presentation_path: Path to the Income Statement presentation  
        output_path: Path for the merged output presentation
    """
    try:
        from pptx import Presentation
        import logging
        
        logging.info(f"🔄 Starting presentation merge...")
        
        # Load the Balance Sheet presentation as the base
        bs_prs = Presentation(bs_presentation_path)
        is_prs = Presentation(is_presentation_path)
        
        logging.info(f"📊 BS presentation has {len(bs_prs.slides)} slides")
        logging.info(f"📈 IS presentation has {len(is_prs.slides)} slides")
        
        # Copy all slides from Income Statement to Balance Sheet presentation
        for slide in is_prs.slides:
            # Get the slide layout from the BS presentation
            slide_layout = bs_prs.slide_layouts[0]  # Use first layout for all slides
            
            # Create new slide in BS presentation
            new_slide = bs_prs.slides.add_slide(slide_layout)
            
            # Copy all shapes from IS slide to new slide
            for shape in slide.shapes:
                # Get shape position and size
                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height
                
                # Copy shape based on type
                if shape.shape_type == 17:  # Text box
                    # Copy text box
                    textbox = new_slide.shapes.add_textbox(left, top, width, height)
                    textbox.text_frame.text = shape.text_frame.text
                    
                    # Copy text formatting
                    for i, paragraph in enumerate(shape.text_frame.paragraphs):
                        if i < len(textbox.text_frame.paragraphs):
                            new_paragraph = textbox.text_frame.paragraphs[i]
                            new_paragraph.alignment = paragraph.alignment
                            for j, run in enumerate(paragraph.runs):
                                if j < len(new_paragraph.runs):
                                    new_run = new_paragraph.runs[j]
                                    new_run.font.bold = run.font.bold
                                    new_run.font.italic = run.font.italic
                                    new_run.font.size = run.font.size
                                    new_run.font.name = run.font.name
                                    if hasattr(run.font, 'color') and run.font.color.rgb:
                                        new_run.font.color.rgb = run.font.color.rgb
                
                elif shape.shape_type == 19:  # Table
                    # Copy table
                    table = new_slide.shapes.add_table(
                        rows=shape.table.rows.__len__(),
                        cols=shape.table.columns.__len__(),
                        left=left,
                        top=top,
                        width=width,
                        height=height
                    ).table
                    
                    # Copy table data
                    for row_idx in range(shape.table.rows.__len__()):
                        for col_idx in range(shape.table.columns.__len__()):
                            if row_idx < table.rows.__len__() and col_idx < table.columns.__len__():
                                table.cell(row_idx, col_idx).text = shape.table.cell(row_idx, col_idx).text
                
                else:
                    # For other shape types, try to copy as picture
                    try:
                        new_slide.shapes.add_picture(
                            shape.image.blob,
                            left,
                            top,
                            width,
                            height
                        )
                    except:
                        # If copying as picture fails, skip this shape
                        logging.warning(f"⚠️ Could not copy shape type {shape.shape_type}")
                        continue
        
        # Save the merged presentation
        bs_prs.save(output_path)
        
        logging.info(f"✅ Successfully merged presentations to: {output_path}")
        logging.info(f"📊 Final presentation has {len(bs_prs.slides)} slides")
        
    except Exception as e:
        logging.error(f"❌ Failed to merge presentations: {str(e)}")
        raise 