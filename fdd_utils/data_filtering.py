"""
Data filtering utilities for FDD application
Removes zero-value rows and cleanses data before AI processing
"""

import pandas as pd
import re
from typing import Dict, List


def filter_zero_value_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filter out rows where ALL value columns are exactly zero (not NaN)
    This ensures AI commentary doesn't include accounts with no activity
    
    Args:
        df: DataFrame with financial data
        
    Returns:
        Filtered DataFrame with zero-value rows removed
    """
    if df.empty:
        return df
    
    # Make a copy to avoid modifying original
    filtered_df = df.copy()
    
    # Identify value columns (numeric columns that are not the first column)
    # Typically: column 0 is description, columns 1+ are date/value columns
    value_cols = []
    for col_idx, col in enumerate(filtered_df.columns):
        if col_idx > 0:  # Skip first column (description)
            # Check if column has numeric data
            sample_values = filtered_df[col].dropna().head(10)
            if len(sample_values) > 0:
                try:
                    pd.to_numeric(sample_values, errors='coerce')
                    if pd.to_numeric(sample_values, errors='coerce').notna().any():
                        value_cols.append(col)
                except:
                    pass
    
    if not value_cols:
        print(f"⚠️ No value columns found for filtering")
        return filtered_df
    
    print(f"🔍 Filtering zero-value rows. Value columns: {value_cols}")
    
    # Convert value columns to numeric
    for col in value_cols:
        filtered_df[col] = pd.to_numeric(filtered_df[col], errors='coerce')
    
    # Create mask to identify rows to keep
    mask = []
    desc_col = filtered_df.columns[0]  # First column is description
    
    kept_count = 0
    removed_count = 0
    
    for idx, row in filtered_df.iterrows():
        desc_value = str(row[desc_col]).strip()
        has_description = desc_value not in ['', 'nan', 'None', 'NaN']
        
        # Check if ALL value columns are exactly zero (not NaN)
        all_values_zero = True
        has_any_data = False
        
        for col in value_cols:
            val = pd.to_numeric(row[col], errors='coerce')
            if pd.notna(val):
                has_any_data = True
                if val != 0:
                    all_values_zero = False
                    break
        
        # Keep row if:
        # 1. Has description AND has non-zero values
        # 2. Has description AND has no data (NaN) - might be header
        # 3. Remove only if has description AND ALL values are exactly 0
        if has_description:
            if not has_any_data or not all_values_zero:
                mask.append(True)
                kept_count += 1
            else:
                mask.append(False)
                removed_count += 1
                print(f"   🗑️ Removing zero-value row: {desc_value}")
        else:
            mask.append(False)  # Remove rows without description
    
    print(f"✅ Filtering complete: Kept {kept_count} rows, Removed {removed_count} rows")
    
    if mask:
        filtered_df = filtered_df[mask].reset_index(drop=True)
    
    return filtered_df


def filter_sections_by_key_for_ai(sections_by_key: Dict[str, List], debug: bool = False) -> Dict[str, List]:
    """
    Filter sections_by_key data structure to remove zero-value content before AI processing
    
    Args:
        sections_by_key: Dict mapping keys to list of table sections (as strings)
        debug: Enable debug printing
        
    Returns:
        Filtered sections_by_key with zero-value rows removed
    """
    filtered_sections = {}
    
    for key, sections in sections_by_key.items():
        if not sections:
            filtered_sections[key] = sections
            continue
        
        if debug:
            print(f"\n🔍 Filtering key: {key}")
            print(f"   Original sections count: {len(sections)}")
        
        filtered_key_sections = []
        
        for section_idx, section in enumerate(sections):
            # Parse section text into DataFrame for filtering
            try:
                # Section is typically a table in text format
                # Try to convert to DataFrame
                lines = section.strip().split('\n')
                
                if len(lines) < 2:
                    # Not enough data to be a table, keep as-is
                    filtered_key_sections.append(section)
                    continue
                
                # Try to parse as table
                # Look for pipe-separated or tab-separated data
                if '|' in lines[0]:
                    # Markdown table format
                    data_rows = []
                    for line in lines:
                        if line.strip() and not line.strip().startswith('|--'):
                            row = [cell.strip() for cell in line.split('|') if cell.strip()]
                            if row:
                                data_rows.append(row)
                    
                    if len(data_rows) > 1:
                        # Create DataFrame
                        df = pd.DataFrame(data_rows[1:], columns=data_rows[0])
                        
                        # Apply zero-value filtering
                        filtered_df = filter_zero_value_rows(df)
                        
                        if not filtered_df.empty:
                            # Convert back to table format
                            filtered_section = df_to_markdown_table(filtered_df)
                            filtered_key_sections.append(filtered_section)
                            
                            if debug:
                                print(f"   ✅ Section {section_idx}: {len(df)} -> {len(filtered_df)} rows")
                        else:
                            if debug:
                                print(f"   🗑️ Section {section_idx}: Completely filtered out (all zeros)")
                    else:
                        filtered_key_sections.append(section)
                else:
                    # Not a table format, keep as-is
                    filtered_key_sections.append(section)
                    
            except Exception as e:
                if debug:
                    print(f"   ⚠️ Section {section_idx}: Could not filter ({e}), keeping as-is")
                filtered_key_sections.append(section)
        
        filtered_sections[key] = filtered_key_sections
        
        if debug:
            print(f"   Filtered sections count: {len(filtered_key_sections)}")
    
    return filtered_sections


def df_to_markdown_table(df: pd.DataFrame) -> str:
    """Convert DataFrame to markdown table format"""
    # Create header
    header = "| " + " | ".join(df.columns) + " |"
    separator = "|" + "|".join(["---" for _ in df.columns]) + "|"
    
    # Create rows
    rows = []
    for _, row in df.iterrows():
        row_str = "| " + " | ".join([str(val) for val in row]) + " |"
        rows.append(row_str)
    
    return "\n".join([header, separator] + rows)


def clean_table_data_for_ai(table_text: str) -> str:
    """
    Clean table data by removing rows with all zero values
    Simpler text-based approach for when DataFrame parsing is not possible
    
    Args:
        table_text: Table data as text
        
    Returns:
        Cleaned table text with zero-value rows removed
    """
    lines = table_text.strip().split('\n')
    filtered_lines = []
    
    for line in lines:
        # Skip empty lines
        if not line.strip():
            continue
        
        # Keep header lines (usually contain text, not numbers)
        # Split by common delimiters
        cells = re.split(r'[\t|,]', line)
        
        # Check if this is a data row (has numbers)
        has_numbers = any(re.search(r'\d', cell) for cell in cells)
        
        if has_numbers:
            # Check if all numeric values are zero
            numeric_values = []
            for cell in cells[1:]:  # Skip first cell (description)
                try:
                    # Extract numeric value
                    num_match = re.search(r'[-+]?\d*\.?\d+', cell.replace(',', ''))
                    if num_match:
                        numeric_values.append(float(num_match.group()))
                except:
                    pass
            
            # Keep row if not all zeros
            if numeric_values:
                if not all(val == 0 for val in numeric_values):
                    filtered_lines.append(line)
                else:
                    print(f"   🗑️ Removing zero-value line: {line[:100]}")
            else:
                filtered_lines.append(line)  # No numeric values found, keep
        else:
            filtered_lines.append(line)  # Header or non-data line, keep
    
    return '\n'.join(filtered_lines)


# Example usage
if __name__ == "__main__":
    # Test data
    test_df = pd.DataFrame({
        'Account': ['Cash', 'AR', 'Zero Account', 'AP'],
        '2023': [100, 200, 0, 300],
        '2024': [150, 250, 0, 350]
    })
    
    print("Original DataFrame:")
    print(test_df)
    
    filtered_df = filter_zero_value_rows(test_df)
    
    print("\nFiltered DataFrame:")
    print(filtered_df)

