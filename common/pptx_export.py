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
        self.ROWS_PER_SECTION = 30  # Use the same value for all sections
        print(f"[DEBUG] ROWS_PER_SECTION: {self.ROWS_PER_SECTION}")
        slide = self.prs.slides[0]
        shape = next(s for s in slide.shapes if s.name == "textMainBullets")
        print(f"[DEBUG] textMainBullets shape height: {shape.height}, LINE_HEIGHT: {self.LINE_HEIGHT}")
        self.CHARS_PER_ROW = 50
        self.BULLET_CHAR = '■ '
        self.DARK_BLUE = RGBColor(0, 50, 150)
        self.DARK_GREY = RGBColor(169, 169, 169)
        self.prev_layer1 = None
        self.prev_layer2 = None
        self.log_template_shapes()  # Debug: print all shape names

    def log_template_shapes(self):
        print("=== Template Shape Audit ===")
        for slide_idx, slide in enumerate(self.prs.slides):
            print(f"Slide {slide_idx + 1} has {len(slide.shapes)} shapes:")
            for shape in slide.shapes:
                print(f"  - Name: '{shape.name}' | Type: {shape.shape_type}")
        print("=== End Shape Audit ===")

    def _calculate_max_rows(self):
        slide = self.prs.slides[0]
        shape = next(s for s in slide.shapes if s.name == "textMainBullets")
        return int(shape.height / self.LINE_HEIGHT) - 3

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
                    current_descs.append(stripped)

        if current_descs:
            items.append(FinancialItem(current_type, current_title, current_descs, is_table=is_table))
        return items

    def _plan_content_distribution(self, items: List[FinancialItem]):
        distribution = []
        content_queue = items.copy()
        slide_idx = 0
        while content_queue:
            if slide_idx == 0:
                sections = ['c']
            else:
                sections = ['b', 'c']
            for section in sections:
                if not content_queue:
                    break
                section_items = []
                lines_used = 0
                while content_queue and lines_used < self.ROWS_PER_SECTION:
                    item = content_queue[0]
                    item_lines = self._calculate_item_lines(item)
                    if lines_used + item_lines <= self.ROWS_PER_SECTION:
                        section_items.append(item)
                        content_queue.pop(0)
                        lines_used += item_lines
                    else:
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
                            break
                if section_items:
                    distribution.append((slide_idx, section, section_items))
            slide_idx += 1
            if slide_idx >= len(self.prs.slides):
                break
        return distribution

    def _calculate_chars_per_line(self, shape):
        # Convert EMU to pixels: 1 EMU = 1/914400 inches, 1 inch = 96 px
        px_width = int(shape.width * 96 / 914400)
        avg_char_px = 7  # Arial 9pt ~ 7px per char
        chars_per_line = max(20, px_width // avg_char_px)
        return chars_per_line

    def _wrap_text_to_shape(self, text, shape):
        # Use the actual shape width and font size to wrap text accurately
        chars_per_line = self._calculate_chars_per_line(shape)
        wrapped = textwrap.wrap(text, width=chars_per_line)
        return wrapped

    def _calculate_item_lines(self, item: FinancialItem) -> int:
        shape = getattr(self, 'current_shape', None)
        chars_per_line = self._calculate_chars_per_line(shape) if shape else self.CHARS_PER_ROW
        lines = 0
        header = f"{item.accounting_type} (continued)" if item.layer1_continued else item.accounting_type
        header_lines = len(textwrap.wrap(header, width=chars_per_line))
        lines += header_lines
        desc_lines = 0
        for desc in item.descriptions:
            for para in desc.split('\n'):
                para_lines = len(textwrap.wrap(para, width=chars_per_line)) or 1
                desc_lines += para_lines
        lines += desc_lines
        print(f"[DEBUG] _calculate_item_lines: Section: {item.account_title}, Header lines: {header_lines}, Desc lines: {desc_lines}, Total: {lines}")
        return lines

    def _split_item(self, item: FinancialItem, max_lines: int) -> tuple[FinancialItem, FinancialItem | None]:
        # Split at paragraph boundaries (not mid-paragraph) whenever possible
        shape = getattr(self, 'current_shape', None)
        chars_per_line = self._calculate_chars_per_line(shape) if shape else self.CHARS_PER_ROW
        header = f"{self.BULLET_CHAR}{item.account_title} - "
        desc_paras = item.descriptions
        lines_used = 0
        split_index = 0
        # Try to fit as many whole paragraphs as possible
        for i, para in enumerate(desc_paras):
            para_lines = len(textwrap.wrap(para, width=chars_per_line)) or 1
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
            wrapped = textwrap.wrap(para, width=chars_per_line)
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
                return next((s for s in slide.shapes if s.name == "textMainBullets"), None)
            return None  # No 'b' section on first slide
        # For subsequent slides
        target_name = "textMainBullets_L" if section == 'b' else "textMainBullets_R"
        return next((s for s in slide.shapes if s.name == target_name), None)

    def _populate_section(self, shape, items: List[FinancialItem]):
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
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
        print("[DEBUG] Entered generate_full_report")
        print(f"[DEBUG] ROWS_PER_SECTION: {self.ROWS_PER_SECTION}")
        print(f"[DEBUG] md_content preview: {md_content[:200]!r}")
        try:
            print("[DEBUG] Parsing markdown...")
            items = self.parse_markdown(md_content)
            print(f"[DEBUG] Parsed {len(items)} items from markdown.")
            print("[DEBUG] Planning content distribution...")
            distribution = self._plan_content_distribution(items)
            print(f"[DEBUG] Planned distribution for {len(distribution)} slide sections.")
            for slide_idx, section, section_items in distribution:
                total_lines = sum(self._calculate_item_lines(item) for item in section_items)
                print(f"[DEBUG] Slide {slide_idx} section {section}: {len(section_items)} items, {total_lines} lines")
            self._validate_content_placement(distribution)
            print("[DEBUG] Content placement validated.")
            summary_text = self.parse_summary_markdown(summary_md)
            print("[DEBUG] Parsed summary markdown.")
            max_slide_used = max((slide_idx for slide_idx, _, _ in distribution), default=0)
            total_slides_needed = max_slide_used + 1
            print(f"[DEBUG] Total slides needed: {total_slides_needed}")
            wrapper = textwrap.TextWrapper(
                width=self.CHARS_PER_ROW, 
                break_long_words=True,
                replace_whitespace=False
            )
            summary_chunks = wrapper.wrap(summary_text)
            chunks_per_slide = len(summary_chunks) // total_slides_needed + 1
            slide_summary_chunks = [
                summary_chunks[i:i+chunks_per_slide] 
                for i in range(0, len(summary_chunks), chunks_per_slide)
            ]
            while len(self.prs.slides) < total_slides_needed:
                self.prs.slides.add_slide(self.prs.slide_layouts[1])
            print("[DEBUG] Populating main content sections...")
            for slide_idx, section, section_items in distribution:
                if slide_idx >= len(self.prs.slides):
                    raise ValueError("Insufficient slides in template")
                slide = self.prs.slides[slide_idx]
                self.current_slide_index = slide_idx
                shape = self._get_section_shape(slide, section)
                if shape:
                    self._populate_section(shape, section_items)
            print("[DEBUG] Populating summary content on slides...")
            for slide_idx in range(total_slides_needed):
                slide = self.prs.slides[slide_idx]
                summary_shape = next((s for s in slide.shapes if s.name == "coSummaryShape"), None)
                if summary_shape:
                    self._populate_summary_section_safe(
                        summary_shape, 
                        slide_summary_chunks[slide_idx] if slide_idx < len(slide_summary_chunks) else []
                    )
            print("[DEBUG] Removing unused slides if any...")
            unused_slides = self._detect_unused_slides(distribution)
            if unused_slides:
                logging.info(f"Removing unused slides: {[idx+1 for idx in unused_slides]}")
                self._remove_slides(unused_slides)
            print(f"[DEBUG] Saving presentation to {output_path}")
            print(f"[DEBUG] Current working directory: {os.getcwd()}")
            self.prs.save(output_path)
            print(f"[DEBUG] Successfully generated PowerPoint with {len(self.prs.slides)} slides and summary content")
        except Exception as e:
            print(f"[ERROR] Generation failed: {str(e)}")
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
        location = self.project_name  # Use the provided project_name/entity
        if location == "Haining":
            summary_content = """
            ## Summary
            The company demonstrates strong financial health with total assets of $180 million, liabilities of $75 million, and shareholder's equity of $105 million. Current assets including $45 million cash and $30 million receivables provide ample liquidity to cover short-term obligations of $50 milliion. Long-term invesmtnets in property and equipment total $90 million, supported by conservative debt levels with a debt-to-equity ratio of 0.71. Retained earnings of $80 million reflect consistent profitability and prudent dividend policies. The balance sheet structure shows optimal asset allocation with 60% long-term investments and 40% working capital. Financial ratios indicate robust solvency with current ratio of 2.4 and quick ratio of 1.8. Equity growth of 12% year-over-year demonstrates sustainable value creation. Conservative accounting practives ensure asset valutations remain realistic, while liability management maintains healthy interest coverage. Overall, the balance sheet positions the company for strategic investments while maintaining financial stability.
            """
        elif location == "Ningbo":
            summary_content = """
            ## Summary
            The company demonstrates strong financial health with total assets of $180 million, liabilities of $75 million, and shareholder's equity of $105 million. Current assets including $45 million cash and $30 million receivables provide ample liquidity to cover short-term obligations of $50 milliion. Long-term invesmtnets in property and equipment total $90 million, supported by conservative debt levels with a debt-to-equity ratio of 0.71. Retained earnings of $80 million reflect consistent profitability and prudent dividend policies. The balance sheet structure shows optimal asset allocation with 60% long-term investments and 40% working capital. Financial ratios indicate robust solvency with current ratio of 2.4 and quick ratio of 1.8. Equity growth of 12% year-over-year demonstrates sustainable value creation. Conservative accounting practives ensure asset valutations remain realistic, while liability management maintains healthy interest coverage. Overall, the balance sheet positions the company for strategic investments while maintaining financial stability.
            """
        elif location == "Nanjing":
            summary_content = """
            ## Summary
            The company demonstrates strong financial health with total assets of $180 million, liabilities of $75 million, and shareholder's equity of $105 million. Current assets including $45 million cash and $30 million receivables provide ample liquidity to cover short-term obligations of $50 milliion. Long-term invesmtnets in property and equipment total $90 million, supported by conservative debt levels with a debt-to-equity ratio of 0.71. Retained earnings of $80 million reflect consistent profitability and prudent dividend policies. The balance sheet structure shows optimal asset allocation with 60% long-term investments and 40% working capital. Financial ratios indicate robust solvency with current ratio of 2.4 and quick ratio of 1.8. Equity growth of 12% year-over-year demonstrates sustainable value creation. Conservative accounting practives ensure asset valutations remain realistic, while liability management maintains healthy interest coverage. Overall, the balance sheet positions the company for strategic investments while maintaining financial stability.
            """
        else:
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
    for slide_index, slide in enumerate(prs.slides):
        current_slide_number = slide_index + 1
        projTitle_shape = find_shape_by_name(slide.shapes, "projTitle")
        if projTitle_shape:
            replacements = {
                "[PROJECT]": project_name,
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
def export_pptx(template_path, markdown_path, output_path, project_name=None):
    logging.info(f"[DEBUG] Starting export_pptx: template={template_path}, markdown={markdown_path}, output={output_path}, project={project_name}")
    generator = ReportGenerator(template_path, markdown_path, output_path, project_name)
    generator.generate()
    if not os.path.exists(output_path):
        logging.error(f"[ERROR] PPTX file was not created at {output_path}")
        raise FileNotFoundError(f"PPTX file was not created at {output_path}")
    logging.info(f"[DEBUG] PPTX file successfully saved at {output_path}")
    if project_name:
        logging.info("[DEBUG] Updating project titles...")
        update_project_titles(output_path, project_name)
        logging.info("[DEBUG] Project titles updated.")
    logging.info(f"[DEBUG] Export complete: {output_path}")
    return output_path 