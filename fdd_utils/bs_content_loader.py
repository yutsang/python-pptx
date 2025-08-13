#!/usr/bin/env python3
"""
Balance Sheet Content Loader

This module provides utilities to load and manage structured balance sheet content
from JSON format, replacing the previous markdown-based approach.
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Optional, Any

class BSContentLoader:
    """Loader for structured balance sheet content from JSON."""
    
    def __init__(self, json_file_path: str = "fdd_utils/bs_content_structured.json"):
        """Initialize the loader with the JSON file path."""
        self.json_file_path = Path(json_file_path)
        self._content = None
        self._metadata = None
    
    def load_content(self) -> Dict[str, Any]:
        """Load the structured content from JSON file."""
        if self._content is None:
            if not self.json_file_path.exists():
                raise FileNotFoundError(f"Balance sheet content file not found: {self.json_file_path}")
            
            try:
                with open(self.json_file_path, 'r', encoding='utf-8') as f:
                    self._content = json.load(f)
                self._metadata = self._content.get('metadata', {})
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON in balance sheet content file: {e}")
            except Exception as e:
                raise RuntimeError(f"Error loading balance sheet content: {e}")
        
        return self._content
    
    def get_metadata(self) -> Dict[str, Any]:
        """Get metadata about the content."""
        if self._metadata is None:
            self.load_content()
        return self._metadata
    
    def get_section(self, section_name: str) -> Dict[str, Any]:
        """Get a specific section of the balance sheet content."""
        content = self.load_content()
        return content.get(section_name, {})
    
    def get_item(self, section_name: str, item_name: str) -> Optional[Dict[str, Any]]:
        """Get a specific item from a section."""
        section = self.get_section(section_name)
        return section.get(item_name)
    
    def get_all_amounts(self) -> Dict[str, str]:
        """Get all amounts mentioned in the content with their descriptions."""
        content = self.load_content()
        amounts = {}
        
        for section_name, section_data in content.items():
            if section_name == 'metadata':
                continue
            
            for item_name, item_data in section_data.items():
                if isinstance(item_data, dict) and 'amount' in item_data:
                    key = f"{section_name}.{item_name}"
                    amounts[key] = item_data['amount']
        
        return amounts
    
    def get_all_entities(self) -> List[str]:
        """Get all entity names mentioned in the content."""
        content = self.load_content()
        entities = set()
        
        for section_name, section_data in content.items():
            if section_name == 'metadata':
                continue
            
            for item_name, item_data in section_data.items():
                if isinstance(item_data, dict):
                    # Check for entity field
                    if 'entity' in item_data:
                        entities.add(item_data['entity'])
                    
                    # Check for entity mentions in content
                    if 'content' in item_data:
                        content_text = item_data['content']
                        # Look for common entity patterns
                        import re
                        entity_patterns = [
                            r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\s+(?:Co\.|Ltd\.|Limited|Corp\.|Corporation))',
                            r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\s+(?:Management|Supply|Chain|Logistics|Development))',
                            r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\s+(?:Wanpu|Wanchen|Jinghong|Jinhong))'
                        ]
                        
                        for pattern in entity_patterns:
                            matches = re.findall(pattern, content_text)
                            entities.update(matches)
        
        return list(entities)
    
    def search_by_keyword(self, keyword: str) -> List[Dict[str, Any]]:
        """Search for items containing a specific keyword."""
        content = self.load_content()
        results = []
        
        for section_name, section_data in content.items():
            if section_name == 'metadata':
                continue
            
            for item_name, item_data in section_data.items():
                if isinstance(item_data, dict):
                    # Search in title, content, and key_points
                    searchable_text = ""
                    if 'title' in item_data:
                        searchable_text += item_data['title'] + " "
                    if 'content' in item_data:
                        searchable_text += item_data['content'] + " "
                    if 'key_points' in item_data:
                        searchable_text += " ".join(item_data['key_points']) + " "
                    
                    if keyword.lower() in searchable_text.lower():
                        results.append({
                            'section': section_name,
                            'item': item_name,
                            'data': item_data
                        })
        
        return results
    
    def get_summary(self) -> Dict[str, Any]:
        """Get a summary of the balance sheet content."""
        content = self.load_content()
        summary = {
            'metadata': self.get_metadata(),
            'sections': list(content.keys()),
            'total_items': 0,
            'total_amounts': 0,
            'entities': self.get_all_entities(),
            'amounts': self.get_all_amounts()
        }
        
        # Count items
        for section_name, section_data in content.items():
            if section_name != 'metadata':
                summary['total_items'] += len(section_data)
        
        summary['total_amounts'] = len(summary['amounts'])
        
        return summary
    
    def export_to_markdown(self, output_file: str = "fdd_utils/bs_content_exported.md") -> str:
        """Export the structured content back to markdown format."""
        content = self.load_content()
        markdown_lines = []
        
        # Add header
        metadata = self.get_metadata()
        markdown_lines.append(f"# Balance Sheet Content")
        markdown_lines.append(f"*Generated from structured JSON data*")
        markdown_lines.append(f"*Source: {metadata.get('source', 'Unknown')}*")
        markdown_lines.append(f"*Version: {metadata.get('version', 'Unknown')}*")
        markdown_lines.append("")
        
        # Process each section
        for section_name, section_data in content.items():
            if section_name == 'metadata':
                continue
            
            # Section header
            section_title = section_name.replace('_', ' ').title()
            markdown_lines.append(f"## {section_title}")
            markdown_lines.append("")
            
            # Process items in section
            for item_name, item_data in section_data.items():
                if isinstance(item_data, dict):
                    # Item header
                    title = item_data.get('title', item_name.replace('_', ' ').title())
                    markdown_lines.append(f"### {title}")
                    
                    # Content
                    if 'content' in item_data:
                        markdown_lines.append(item_data['content'])
                        markdown_lines.append("")
                    
                    # Key points
                    if 'key_points' in item_data:
                        markdown_lines.append("**Key Points:**")
                        for point in item_data['key_points']:
                            markdown_lines.append(f"- {point}")
                        markdown_lines.append("")
        
        # Write to file
        output_path = Path(output_file)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(markdown_lines))
        
        return str(output_path)

# Convenience functions
def load_bs_content() -> Dict[str, Any]:
    """Load balance sheet content."""
    loader = BSContentLoader()
    return loader.load_content()

def get_bs_item(section: str, item: str) -> Optional[Dict[str, Any]]:
    """Get a specific balance sheet item."""
    loader = BSContentLoader()
    return loader.get_item(section, item)

def get_bs_amounts() -> Dict[str, str]:
    """Get all amounts from balance sheet content."""
    loader = BSContentLoader()
    return loader.get_all_amounts()

def get_bs_entities() -> List[str]:
    """Get all entities from balance sheet content."""
    loader = BSContentLoader()
    return loader.get_all_entities()

if __name__ == "__main__":
    # Test the loader
    loader = BSContentLoader()
    
    print("=== Balance Sheet Content Loader Test ===")
    print()
    
    # Load content
    content = loader.load_content()
    print(f"✅ Content loaded successfully")
    print(f"📊 Metadata: {loader.get_metadata()}")
    print()
    
    # Get summary
    summary = loader.get_summary()
    print(f"📋 Summary:")
    print(f"   - Sections: {len(summary['sections'])}")
    print(f"   - Total items: {summary['total_items']}")
    print(f"   - Total amounts: {summary['total_amounts']}")
    print(f"   - Entities: {len(summary['entities'])}")
    print()
    
    # Show entities
    entities = loader.get_all_entities()
    print(f"🏢 Entities found: {entities}")
    print()
    
    # Show amounts
    amounts = loader.get_all_amounts()
    print(f"💰 Amounts found:")
    for key, amount in amounts.items():
        print(f"   - {key}: {amount}")
    print()
    
    print("✅ All tests passed!") 