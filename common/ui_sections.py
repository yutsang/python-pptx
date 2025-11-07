import pandas as pd
import streamlit as st


def clean_header_rows(df):
    """
    Clean up dataframe by removing header rows that contain unit indicators 
    (like 人民幣千元) combined with year values (like 2021000).
    """
    if df.empty:
        return df
    
    # Check first few rows for problematic patterns
    rows_to_drop = []
    
    for row_idx in range(min(3, len(df))):  # Check first 3 rows
        row = df.iloc[row_idx]
        row_str = ' '.join(str(val) for val in row if pd.notna(val))
        
        # Check if row contains unit indicators
        unit_indicators = ['人民幣千元', '人民币千元', "CNY'000", 'thousands']
        has_unit_indicator = any(indicator in row_str for indicator in unit_indicators)
        
        if has_unit_indicator:
            # Check if row also contains year-like values (e.g., 2021000, 2020000)
            # BUT NOT large numbers like 2,021,000 which are actual financial figures
            for val in row:
                if pd.notna(val):
                    val_str = str(val)
                    # Skip if it has commas and is a large number (definitely financial data)
                    if ',' in val_str:
                        try:
                            val_num = float(val_str.replace(',', ''))
                            if val_num > 100000:  # Large numbers are financial data, not years
                                continue
                        except:
                            pass
                    # Check if it's a numeric value that looks like year + 000 (e.g., 2021000)
                    if val_str.replace(',', '').replace('.', '').isdigit():
                        val_num = float(val_str.replace(',', ''))
                        # If it's between 2000000 and 2100000 AND doesn't have commas, it's likely year + 000
                        # Exclude values with commas (like 2,021,000) which are actual amounts
                        if 2000000 <= val_num <= 2100000:
                            rows_to_drop.append(row_idx)
                            print(f"   🧹 Removing header row {row_idx} with unit indicator and year value: {val_str}")
                            break
    
    # Drop the identified rows
    if rows_to_drop:
        df = df.drop(df.index[rows_to_drop]).reset_index(drop=True)
        print(f"   ✅ Cleaned {len(rows_to_drop)} header rows")
    
    return df


def preprocess_income_statement_table(df):
    """
    Preprocess income statement table:
    - Remove rows with '示意性调整后'
    - Round numbers to 1 decimal place
    - Add thousand separators
    """
    if df.empty:
        return df
    
    print(f"   🔧 Preprocessing Income Statement table...")
    
    # 1. Remove rows containing "示意性调整后" or "Indicative adjusted"
    mask = df.apply(lambda row: not any(
        '示意性调整后' in str(val) or '示意性調整後' in str(val) or 
        '经示意性调整后' in str(val) or '經示意性調整後' in str(val) or
        'indicative adjusted' in str(val).lower()
        for val in row
    ), axis=1)
    rows_before = len(df)
    df = df[mask]
    if rows_before != len(df):
        print(f"   ✅ Removed {rows_before - len(df)} rows with '示意性调整后'")
    
    # 1.5. Limit table to rows up to and including "净利润" row
    net_profit_keywords = ['净利润', '淨利潤', 'net profit', 'net income']
    net_profit_row_idx = None
    for idx, row in df.iterrows():
        row_str = ' '.join(str(val).lower() for val in row if pd.notna(val))
        if any(keyword in row_str for keyword in net_profit_keywords):
            net_profit_row_idx = idx
            print(f"   📍 Found '净利润' at row index {idx}")
            break
    
    if net_profit_row_idx is not None:
        # Keep only rows up to and including 净利润
        df = df.loc[:net_profit_row_idx]
        print(f"   ✅ Limited table to {len(df)} rows (up to 净利润)")
    
    # 2. Round numerical columns to 1 decimal place and add thousand separators
    for col in df.columns:
        if df[col].dtype in ['float64', 'int64', 'float32', 'int32']:
            # Round to 1 decimal and format with thousand separators
            # Skip if values look like years (2020-2030 range)
            def format_with_check(x):
                if pd.notna(x):
                    # Don't format if it's a year value (between 2000 and 2100)
                    if 2000 <= x <= 2100:
                        return x
                    return f"{x:,.1f}"
                return x
            df[col] = df[col].apply(format_with_check)
        elif df[col].dtype == 'object':
            # Try to convert string numbers to formatted numbers
            def format_number(val):
                if pd.isna(val) or val == '':
                    return val
                try:
                    # Skip if already contains Chinese characters or is descriptive text
                    if any('\u4e00' <= char <= '\u9fff' for char in str(val)):
                        return val
                    # Skip if it looks like a date or year (e.g., "2024", "2021-01-01", "2024年1-5月")
                    val_str = str(val)
                    if any(char in val_str for char in ['年', '月', '日', '-', '/', '至']):
                        return val
                    # Try to parse as number
                    num = float(str(val).replace(',', '').replace(' ', ''))
                    # Don't format if it's a year value (between 2000 and 2100)
                    if 2000 <= num <= 2100:
                        return val
                    return f"{num:,.1f}"
                except (ValueError, TypeError):
                    return val
            df[col] = df[col].apply(format_number)
    print(f"   ✅ Applied number formatting: 1 decimal place with thousand separators")
    
    return df


def render_balance_sheet_sections(
    sections_by_key: dict,
    get_key_display_name,
    selected_entity: str,
    format_date_to_dd_mmm_yyyy,
):
    """Render Balance Sheet sections UI using existing parsed/cleaned data.

    Parameters
    - sections_by_key: mapping key -> list of sections
    - get_key_display_name: function to map key code to display
    - selected_entity: entity string for context
    - format_date_to_dd_mmm_yyyy: callable to format dates
    """
    
    # Debug data flow
    keys_with_data = [key for key, sections in sections_by_key.items() if sections]
    print(f"🔍 UI BS: {len(keys_with_data)} keys with data")
    
    if not keys_with_data:
        st.warning("No data found for any financial keys.")
        print("❌ UI BS: No keys with data - this explains missing tables")
        return

    key_tabs = st.tabs([get_key_display_name(key) for key in keys_with_data])
    for i, key in enumerate(keys_with_data):
        with key_tabs[i]:
            sections = sections_by_key[key]
            if not sections:
                st.info("No sections found for this key.")
                continue

            # Simple entity filtering - only show sections that match the selected entity
            filtered_sections = []
            for section in sections:
                section_entity = section.get('entity_name', '')
                # Use flexible matching - if either entity name contains the other
                if (selected_entity.lower() in section_entity.lower() or 
                    section_entity.lower() in selected_entity.lower() or
                    selected_entity.lower() == section_entity.lower()):
                    filtered_sections.append(section)
            
            if not filtered_sections:
                st.info(f"No sections found for entity '{selected_entity}' in this key.")
                continue
            
            sections = filtered_sections
            print(f"🔍 BS Showing {len(sections)} sections for entity '{selected_entity}' in key '{key}'")

            # Debug information (only shown if needed)
            if 'parsed_data' in sections[0] and sections[0]['parsed_data']:
                metadata = sections[0]['parsed_data']['metadata']
                # Keep minimal debug info for troubleshooting if needed
                pass
            
            first_section = sections[0]

            # If we have structured data, prefer it
            if 'parsed_data' in first_section and first_section['parsed_data']:
                parsed_data = first_section['parsed_data']
                metadata = parsed_data['metadata']
                data_rows = parsed_data['data']
                
                # Metadata summary row (info bar)
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(f"**Table:** {metadata.get('table_name', 'N/A')}")
                with col2:
                    date_value = metadata.get('date')
                    if date_value:
                        try:
                            formatted_date = format_date_to_dd_mmm_yyyy(date_value)
                            st.markdown(f"**Date:** {formatted_date}")
                        except Exception as e:
                            st.markdown(f"**Date:** {str(date_value)}")
                    else:
                        st.markdown("**Date:** N/A")
                with col3:
                    currency = metadata.get('currency', 'CNY')
                    multiplier = metadata.get('multiplier', 1)
                    if multiplier > 1:
                        if multiplier == 1000:
                            st.markdown(f"**Currency:** {currency}'000")
                        elif multiplier == 1000000:
                            st.markdown(f"**Currency:** {currency}'000,000")
                        else:
                            st.markdown(f"**Currency:** {currency} (×{multiplier})")
                    else:
                        st.markdown(f"**Currency:** {currency}")
                with col4:
                    st.markdown(f"**Type:** Balance Sheet")

                if data_rows:
                    # Check if we have multiple description columns
                    num_desc_cols = data_rows[0].get('num_desc_columns', 1) if data_rows and isinstance(data_rows[0], dict) else 1
                    
                    structured_data = []
                    for row in data_rows:
                        # Handle different data structures safely
                        if isinstance(row, dict):
                            value = row.get('value') or row.get('Value') or row.get('amount')
                        else:
                            value = None
                        actual_value = value
                        
                        # SKIP ROWS WITH ZERO OR NULL VALUES IMMEDIATELY - DON'T SHOW THEM AT ALL
                        if actual_value is None:
                            continue
                        
                        # Try to convert to number for zero-checking
                        try:
                            # Convert to float for comparison
                            if isinstance(actual_value, str):
                                # Remove commas and convert
                                val_num = float(str(actual_value).replace(',', '').strip())
                            else:
                                val_num = float(actual_value)
                            
                            # Skip if zero or very close to zero - DON'T DISPLAY AT ALL
                            if abs(val_num) < 0.01:
                                continue
                                
                            # Format the value for display (only if non-zero)
                            formatted_value = f"{val_num:,.0f}"
                        except (ValueError, TypeError):
                            # If can't convert to number, it might be text/date
                            # Convert datetime objects to strings
                            if hasattr(actual_value, 'strftime'):
                                formatted_value = actual_value.strftime('%Y-%m-%d')
                            else:
                                # For text values, keep as-is
                                formatted_value = str(actual_value)
                        
                        # Create row with separate description columns if available
                        if num_desc_cols > 1 and isinstance(row, dict):
                            row_data = {
                                'Category': row.get('description_0', ''),
                                'Subcategory': row.get('description_1', ''),
                                'Value': formatted_value
                            }
                        else:
                            if isinstance(row, dict):
                                description = row.get('description') or row.get('Description') or row.get('desc') or row.get('item') or 'Unknown'
                            else:
                                description = str(row)
                            row_data = {'Description': description, 'Value': formatted_value}
                        
                        structured_data.append(row_data)

                    if not structured_data:
                        # If no data found, show a "Total: 0" row instead of info message
                        if num_desc_cols > 1:
                            structured_data = [{'Category': 'Total | 总计', 'Subcategory': '', 'Value': '0'}]
                        else:
                            structured_data = [{'Description': 'Total | 总计', 'Value': '0'}]
                    
                    df_structured = pd.DataFrame(structured_data)

                    def highlight_totals(row):
                        # Check first column for total keywords
                        first_col_val = str(row.iloc[0]).lower()
                        if any(keyword in first_col_val for keyword in ['total', 'subtotal', '总计', '合计', '小计']):
                            return ['background-color: rgba(173, 216, 230, 0.3)'] * len(row)
                        return [''] * len(row)

                    styled_df = df_structured.style.apply(highlight_totals, axis=1)
                    
                    # Configure dataframe display based on number of columns
                    if num_desc_cols > 1:
                        st.dataframe(
                            styled_df, 
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Category": st.column_config.TextColumn(
                                    "Category",
                                    width="medium"
                                ),
                                "Subcategory": st.column_config.TextColumn(
                                    "Subcategory",
                                    width="medium"
                                ),
                                "Value": st.column_config.TextColumn(
                                    "Value", 
                                    width="medium"
                                )
                            }
                        )
                    else:
                        st.dataframe(
                            styled_df, 
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Description": st.column_config.TextColumn(
                                    "Description",
                                    width="large"
                                ),
                                "Value": st.column_config.TextColumn(
                                    "Value", 
                                    width="medium"
                                )
                            }
                        )
                else:
                    st.info("No structured data rows found")

                with st.expander("📋 Structured Markdown", expanded=False):
                    st.code(first_section.get('markdown', 'No markdown available'), language='markdown')
                continue

            # Fallback: render raw DataFrame with cleaning
            raw_df = first_section['data'].copy()
            
            # Clean header rows first
            raw_df = clean_header_rows(raw_df)
            
            # Drop all-NaN/None columns
            for col in list(raw_df.columns):
                if raw_df[col].isna().all() or (raw_df[col].astype(str) == 'None').all():
                    raw_df = raw_df.drop(columns=[col])

            # Rename columns for display clarity
            if len(raw_df.columns) >= 2:
                new_column_names = [f"{key} (Description)", f"{key} (Balance)"]
                if len(raw_df.columns) > 2:
                    for i2 in range(2, len(raw_df.columns)):
                        new_column_names.append(f"{key} (Column {i2+1})")
                raw_df.columns = new_column_names
            elif len(raw_df.columns) == 1:
                raw_df.columns = [f"{key} (Description)"]

            if len(raw_df.columns) > 0:
                # Filter out rows with all zero values
                filtered_df = raw_df.copy()
                
                # Identify numeric columns
                numeric_cols = []
                for col in filtered_df.columns:
                    try:
                        # Convert to numeric, errors='coerce' will set non-numeric to NaN
                        pd.to_numeric(filtered_df[col], errors='coerce')
                        numeric_cols.append(col)
                    except:
                        pass
                
                # Filter out rows where ALL non-description columns are zero
                if len(filtered_df.columns) > 1:  # Only filter if we have more than description column
                    # Assume first column is description, check others for zero values
                    desc_col = filtered_df.columns[0]
                    value_cols = [col for col in filtered_df.columns if col != desc_col]
                    
                    if value_cols:
                        # Convert value columns to numeric
                        for col in value_cols:
                            filtered_df[col] = pd.to_numeric(filtered_df[col], errors='coerce')
                        
                        # Keep rows where at least one value column is non-zero OR has description
                        mask = pd.Series([False] * len(filtered_df), index=filtered_df.index)
                        
                        # Check each row using improved logic
                        for idx, row in filtered_df.iterrows():
                            desc_value = str(row[desc_col]).strip()
                            has_description = desc_value not in ['', 'nan', 'None', 'NaN']
                            
                            if has_description:
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
                                # 1. Has non-zero values
                                # 2. Is header row (NaN values)
                                # 3. Only filter if ALL values are exactly 0
                                if not has_any_data or not all_values_zero:
                                    mask[idx] = True
                        
                        if mask.any():
                            filtered_df = filtered_df[mask]
                        else:
                            # If no rows pass the filter, keep original data
                            filtered_df = raw_df.copy()
                
                # Convert datetime objects to strings to avoid Arrow serialization errors
                for col in filtered_df.columns:
                    if filtered_df[col].dtype == 'object':
                        filtered_df[col] = filtered_df[col].astype(str)
                
                # Configure display
                if len(filtered_df) > 0:
                    st.dataframe(
                        filtered_df, 
                        use_container_width=True,
                        hide_index=True,
                        column_config={col: st.column_config.TextColumn(col, width="large") for col in filtered_df.columns}
                    )
                else:
                    st.info("No rows with non-zero values found")
            else:
                st.error("No valid columns found after cleaning")
                st.write("**Original DataFrame:**")
                st.dataframe(first_section['data'], use_container_width=True)

            # Show parsed data json if present (debug)
            if 'parsed_data' in first_section:
                with st.expander("🔍 Debug: Parsed Data Structure", expanded=False):
                    st.json(first_section['parsed_data'])


def render_combined_sections(
    sections_by_key: dict,
    get_key_display_name,
    selected_entity: str,
    format_date_to_dd_mmm_yyyy,
):
    """Render combined Balance Sheet and Income Statement sections UI."""

    st.markdown("#### 📊")
    
    # High-level debug only
    keys_with_data = [key for key, sections in sections_by_key.items() if sections]
    
    if not keys_with_data:
        st.warning("No data found for any financial keys.")
        return

    key_tabs = st.tabs([get_key_display_name(key) for key in keys_with_data])
    for i, key in enumerate(keys_with_data):
        with key_tabs[i]:
            sections = sections_by_key[key]
            if not sections:
                st.info("No sections found for this key.")
                continue

            # Simple entity filtering - only show sections that match the selected entity
            filtered_sections = []
            for section in sections:
                section_entity = section.get('entity_name', '')
                # Use flexible matching - if either entity name contains the other
                if (selected_entity.lower() in section_entity.lower() or 
                    section_entity.lower() in selected_entity.lower() or
                    selected_entity.lower() == section_entity.lower()):
                    filtered_sections.append(section)
                else:
                    print(f"❌ Excluding section: entity='{section_entity}' doesn't match selected='{selected_entity}'")
            
            if not filtered_sections:
                st.info(f"No sections found for entity '{selected_entity}' in this key.")
                continue
            
            sections = filtered_sections
            print(f"🔍 Showing {len(sections)} sections for entity '{selected_entity}' in key '{key}'")

            # Debug information (only shown if needed)
            if 'parsed_data' in sections[0] and sections[0]['parsed_data']:
                metadata = sections[0]['parsed_data']['metadata']
                # Keep minimal debug info for troubleshooting if needed
                pass
            
            first_section = sections[0]

            # If we have structured data, prefer it
            if 'parsed_data' in first_section and first_section['parsed_data']:
                parsed_data = first_section['parsed_data']
                metadata = parsed_data['metadata']
                data_rows = parsed_data['data']
                
                # Metadata summary row (info bar) - matching BS/IS format
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(f"**Table:** {metadata.get('table_name', 'N/A')}")
                with col2:
                    date_value = metadata.get('date')
                    if date_value:
                        try:
                            formatted_date = format_date_to_dd_mmm_yyyy(date_value)
                            st.markdown(f"**Date:** {formatted_date}")
                        except Exception as e:
                            st.markdown(f"**Date:** {str(date_value)}")
                    else:
                        st.markdown("**Date:** N/A")
                with col3:
                    currency = metadata.get('currency', 'CNY')
                    multiplier = metadata.get('multiplier', 1)
                    if multiplier > 1:
                        if multiplier == 1000:
                            st.markdown(f"**Currency:** {currency}'000")
                        elif multiplier == 1000000:
                            st.markdown(f"**Currency:** {currency}'000,000")
                        else:
                            st.markdown(f"**Currency:** {currency} (×{multiplier})")
                    else:
                        st.markdown(f"**Currency:** {currency}")
                with col4:
                    st.markdown(f"**Type:** Combined (BS+IS)")
                
                # Display the data table in Description/Value format (matching BS/IS mode)
                if data_rows:
                    structured_data = []
                    for row in data_rows:
                        # Handle different data structures safely
                        if isinstance(row, dict):
                            description = row.get('description') or row.get('Description') or row.get('desc') or row.get('item') or str(list(row.keys())[0] if row.keys() else 'Unknown')
                        else:
                            description = str(row)
                        if isinstance(row, dict):
                            value = row.get('value') or row.get('Value') or row.get('amount') or str(list(row.values())[1] if len(row.values()) > 1 else list(row.values())[0] if row.values() else 'N/A')
                        else:
                            value = 'N/A'
                        actual_value = value
                        
                        # Skip rows with zero values
                        try:
                            if isinstance(actual_value, (int, float)) and actual_value == 0:
                                continue
                            elif isinstance(actual_value, str) and actual_value.strip() in ['0', '0.0', '0.00', '-']:
                                continue
                        except:
                            pass  # If conversion fails, keep the row
                        
                        # Convert datetime objects to strings to avoid Arrow serialization errors
                        if hasattr(actual_value, 'strftime'):  # datetime object
                            formatted_value = actual_value.strftime('%Y-%m-%d')
                        elif isinstance(actual_value, (int, float)):
                            formatted_value = f"{actual_value:,.0f}"  # No decimals for cleaner display
                        else:
                            formatted_value = str(actual_value)
                        structured_data.append({'Description': description, 'Value': formatted_value})

                    if structured_data:  # Only display if we have data after filtering zeros
                        df_structured = pd.DataFrame(structured_data)
                        display_df = df_structured[["Description", "Value"]].copy()

                        def highlight_items(row):
                            desc_lower = row['Description'].lower()
                            if any(term in desc_lower for term in ['total', 'subtotal']):
                                return ['background-color: rgba(173, 216, 230, 0.3)'] * len(row)
                            return [''] * len(row)

                        styled_df = display_df.style.apply(highlight_items, axis=1)
                        
                        # Configure dataframe display to prevent wrapping and hide index
                        st.dataframe(
                            styled_df, 
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Description": st.column_config.TextColumn(
                                    "Description",
                                    width="large"
                                ),
                                "Value": st.column_config.TextColumn(
                                    "Value", 
                                    width="medium"
                                )
                            }
                        )
                    else:
                        st.info("No data rows with non-zero values found")
                else:
                    st.info("No structured data rows found")
            else:
                # Fallback to raw data display
                st.markdown("**📋 Raw Data:**")
                st.json(first_section)

def render_income_statement_sections(
    sections_by_key: dict,
    get_key_display_name,
    selected_entity: str,
    format_date_to_dd_mmm_yyyy,
):
    """Render Income Statement sections UI using existing parsed/cleaned data.

    Parameters
    - sections_by_key: mapping key -> list of sections
    - get_key_display_name: function to map key code to display
    - selected_entity: entity string for context
    - format_date_to_dd_mmm_yyyy: callable to format dates
    """
    
    # High-level debug only
    keys_with_data = [key for key, sections in sections_by_key.items() if sections]
    

    
    if not keys_with_data:
        st.warning("No income statement data found for any financial keys.")
        return

    key_tabs = st.tabs([get_key_display_name(key) for key in keys_with_data])
    for i, key in enumerate(keys_with_data):
        with key_tabs[i]:
            sections = sections_by_key[key]
            if not sections:
                st.info("No sections found for this key.")
                continue

            # Simple entity filtering - only show sections that match the selected entity
            filtered_sections = []
            for section in sections:
                section_entity = section.get('entity_name', '')
                # Use flexible matching - if either entity name contains the other
                if (selected_entity.lower() in section_entity.lower() or 
                    section_entity.lower() in selected_entity.lower() or
                    selected_entity.lower() == section_entity.lower()):
                    filtered_sections.append(section)
                else:
                    print(f"❌ Excluding section: entity='{section_entity}' doesn't match selected='{selected_entity}'")
            
            if not filtered_sections:
                st.info(f"No sections found for entity '{selected_entity}' in this key.")
                continue
            
            sections = filtered_sections
            print(f"🔍 Showing {len(sections)} sections for entity '{selected_entity}' in key '{key}'")

            # Debug information (only shown if needed)
            if 'parsed_data' in sections[0] and sections[0]['parsed_data']:
                metadata = sections[0]['parsed_data']['metadata']
                # Keep minimal debug info for troubleshooting if needed
                pass
            
            first_section = sections[0]

            # If we have structured data, prefer it
            if 'parsed_data' in first_section and first_section['parsed_data']:
                parsed_data = first_section['parsed_data']
                metadata = parsed_data['metadata']
                data_rows = parsed_data['data']
                
                # Metadata summary row (info bar)
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(f"**Table:** {metadata.get('table_name', 'N/A')}")
                with col2:
                    date_value = metadata.get('date')
                    if date_value:
                        try:
                            formatted_date = format_date_to_dd_mmm_yyyy(date_value)
                            st.markdown(f"**Date:** {formatted_date}")
                        except Exception as e:
                            st.markdown(f"**Date:** {str(date_value)}")
                    else:
                        st.markdown("**Date:** N/A")
                with col3:
                    currency = metadata.get('currency', 'CNY')
                    multiplier = metadata.get('multiplier', 1)
                    if multiplier > 1:
                        if multiplier == 1000:
                            st.markdown(f"**Currency:** {currency}'000")
                        elif multiplier == 1000000:
                            st.markdown(f"**Currency:** {currency}'000,000")
                        else:
                            st.markdown(f"**Currency:** {currency} (×{multiplier})")
                    else:
                        st.markdown(f"**Currency:** {currency}")
                with col4:
                    st.markdown(f"**Type:** Income Statement")

                if data_rows:
                    structured_data = []
                    for row in data_rows:
                        # Handle different data structures safely
                        if isinstance(row, dict):
                            description = row.get('description') or row.get('Description') or row.get('desc') or row.get('item') or str(list(row.keys())[0] if row.keys() else 'Unknown')
                        else:
                            description = str(row)
                        if isinstance(row, dict):
                            value = row.get('value') or row.get('Value') or row.get('amount') or str(list(row.values())[1] if len(row.values()) > 1 else list(row.values())[0] if row.values() else 'N/A')
                        else:
                            value = 'N/A'
                        actual_value = value
                        
                        # Skip rows with zero values
                        try:
                            if isinstance(actual_value, (int, float)) and actual_value == 0:
                                continue
                            elif isinstance(actual_value, str) and actual_value.strip() in ['0', '0.0', '0.00', '-']:
                                continue
                        except:
                            pass  # If conversion fails, keep the row
                        
                        # Convert datetime objects to strings to avoid Arrow serialization errors
                        if hasattr(actual_value, 'strftime'):  # datetime object
                            formatted_value = actual_value.strftime('%Y-%m-%d')
                        elif isinstance(actual_value, (int, float)):
                            formatted_value = f"{actual_value:,.0f}"  # No decimals for cleaner display
                        else:
                            formatted_value = str(actual_value)
                        structured_data.append({'Description': description, 'Value': formatted_value})

                    if structured_data:  # Only display if we have data after filtering zeros
                        df_structured = pd.DataFrame(structured_data)
                        display_df = df_structured[["Description", "Value"]].copy()

                        def highlight_income_items(row):
                            desc_lower = row['Description'].lower()
                            if any(term in desc_lower for term in ['total', 'subtotal', 'net income', 'gross profit', 'operating income']):
                                return ['background-color: rgba(173, 216, 230, 0.3)'] * len(row)
                            elif any(term in desc_lower for term in ['income', 'revenue', 'gain']):
                                return ['background-color: rgba(144, 238, 144, 0.2)'] * len(row)  # Light green for income
                            elif any(term in desc_lower for term in ['expense', 'cost', 'loss']):
                                return ['background-color: rgba(255, 182, 193, 0.2)'] * len(row)  # Light red for expenses
                            return [''] * len(row)

                        styled_df = display_df.style.apply(highlight_income_items, axis=1)
                        
                        # Configure dataframe display to prevent wrapping and hide index
                        st.dataframe(
                            styled_df, 
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Description": st.column_config.TextColumn(
                                    "Description",
                                    width="large"
                                ),
                                "Value": st.column_config.TextColumn(
                                    "Value", 
                                    width="medium"
                                )
                            }
                        )
                    else:
                        st.info("No data rows with non-zero values found")
                else:
                    st.info("No structured data rows found")

                with st.expander("📋 Structured Markdown", expanded=False):
                    st.code(first_section.get('markdown', 'No markdown available'), language='markdown')
                continue

            # Fallback: render raw DataFrame with cleaning
            raw_df = first_section['data'].copy()
            
            # Clean header rows first
            raw_df = clean_header_rows(raw_df)
            
            # Apply income statement preprocessing
            raw_df = preprocess_income_statement_table(raw_df)
            
            # Drop all-NaN/None columns
            for col in list(raw_df.columns):
                if raw_df[col].isna().all() or (raw_df[col].astype(str) == 'None').all():
                    raw_df = raw_df.drop(columns=[col])

            # Rename columns for display clarity (income statement specific)
            if len(raw_df.columns) >= 2:
                new_column_names = [f"{key} (Description)", f"{key} (Amount)"]
                if len(raw_df.columns) > 2:
                    for i2 in range(2, len(raw_df.columns)):
                        new_column_names.append(f"{key} (Column {i2+1})")
                raw_df.columns = new_column_names
            elif len(raw_df.columns) == 1:
                raw_df.columns = [f"{key} (Description)"]

            if len(raw_df.columns) > 0:
                # Apply income statement specific highlighting
                def highlight_income_statement(row):
                    row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
                    row_lower = row_str.lower()
                    
                    if any(term in row_lower for term in ['total', 'subtotal', 'net income', 'gross profit', 'operating income']):
                        return ['background-color: rgba(173, 216, 230, 0.3)'] * len(row)
                    elif any(term in row_lower for term in ['income', 'revenue', 'gain']):
                        return ['background-color: rgba(144, 238, 144, 0.2)'] * len(row)  # Light green for income
                    elif any(term in row_lower for term in ['expense', 'cost', 'loss']):
                        return ['background-color: rgba(255, 182, 193, 0.2)'] * len(row)  # Light red for expenses
                    return [''] * len(row)
                
                styled_df = raw_df.style.apply(highlight_income_statement, axis=1)
                st.dataframe(styled_df, use_container_width=True, hide_index=True)
            else:
                st.error("No valid columns found after cleaning")
                st.write("**Original DataFrame:**")
                st.dataframe(first_section['data'], use_container_width=True)

            # Show parsed data json if present (debug)
            if 'parsed_data' in first_section:
                with st.expander("🔍 Debug: Parsed Data Structure", expanded=False):
                    st.json(first_section['parsed_data'])


