import pandas as pd
import streamlit as st


def render_balance_sheet_sections(
    sections_by_key: dict,
    get_key_display_name,
    selected_entity: str,
    format_date_to_dd_mmm_yyyy,
):
    """Render Balance Sheet sections UI using existing parsed/cleaned data."""

    st.markdown("#### View Table by Key")
    
    # High-level debug only
    keys_with_data = [key for key, sections in sections_by_key.items() if sections]
    print(f"DEBUG UI: Processing {len(keys_with_data)} keys with data")
    
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

            # Debug: Show all sections and their entity names
            print(f"\n🔍 DEBUG DETAILED: Key '{key}' has {len(sections)} total sections")
            print(f"🔍 DEBUG DETAILED: Selected entity: '{selected_entity}'")
            
            for idx, section in enumerate(sections):
                section_entity = section.get('entity_name', 'NO_ENTITY_NAME')
                sheet_name = section.get('sheet_name', 'NO_SHEET_NAME')
                print(f"🔍 DEBUG DETAILED: Section {idx}: entity='{section_entity}', sheet='{sheet_name}'")
                
                # Check if section has parsed_data and what entity info is there
                if 'parsed_data' in section and section['parsed_data']:
                    metadata = section['parsed_data'].get('metadata', {})
                    metadata_entity = metadata.get('entity_name', 'NO_METADATA_ENTITY')
                    print(f"🔍 DEBUG DETAILED: Section {idx} metadata entity: '{metadata_entity}'")

            # Filter sections by selected entity
            filtered_sections = []
            for section in sections:
                section_entity = section.get('entity_name', '')
                
                # Try multiple matching strategies
                match_found = False
                if section_entity:
                    # Strategy 1: Exact match (case insensitive)
                    if selected_entity.lower() == section_entity.lower():
                        match_found = True
                        print(f"✅ EXACT MATCH: '{section_entity}' == '{selected_entity}'")
                    
                    # Strategy 2: Selected entity contains section entity
                    elif selected_entity.lower() in section_entity.lower():
                        match_found = True
                        print(f"✅ CONTAINS MATCH: '{selected_entity}' in '{section_entity}'")
                    
                    # Strategy 3: Section entity contains selected entity
                    elif section_entity.lower() in selected_entity.lower():
                        match_found = True
                        print(f"✅ REVERSE CONTAINS MATCH: '{section_entity}' in '{selected_entity}'")
                    
                    # Strategy 4: Check for partial word matches
                    else:
                        selected_words = selected_entity.lower().split()
                        section_words = section_entity.lower().split()
                        if any(word in section_words for word in selected_words):
                            match_found = True
                            print(f"✅ WORD MATCH: Common words between '{selected_entity}' and '{section_entity}'")
                
                if match_found:
                    filtered_sections.append(section)
                else:
                    print(f"❌ NO MATCH: '{section_entity}' vs '{selected_entity}'")
            
            print(f"🔍 DEBUG DETAILED: Filtered {len(sections)} -> {len(filtered_sections)} sections")
            
            if not filtered_sections:
                st.info(f"No sections found for entity '{selected_entity}' in this key.")
                continue
            
            sections = filtered_sections
            print(f"🔍 DEBUG: Showing {len(sections)} sections for entity '{selected_entity}' in key '{key}'")

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

                # Metadata summary row
                col1, col2, col3, col4, col5, col6 = st.columns(6)
                with col1:
                    st.markdown(f"**Table:** {metadata['table_name']}")
                with col2:
                    date_value = metadata.get('date')
                    if date_value:
                        try:
                            formatted_date = format_date_to_dd_mmm_yyyy(date_value)
                            st.markdown(f"**Date:** {formatted_date}")
                        except Exception as e:
                            st.markdown(f"**Date:** Error formatting: {e}")
                    else:
                        st.markdown("**Date:** Unknown")
                with col3:
                    currency = metadata.get('currency', 'CNY')
                    st.markdown(f"**Currency:** {currency}")
                with col4:
                    multiplier = metadata.get('multiplier', 1)
                    st.markdown(f"**Multiplier:** {multiplier}x")
                with col5:
                    # Show Indicative adjusted column info instead of Excel row
                    selected_column = metadata.get('selected_column')
                    if selected_column:
                        st.markdown(f"**Indicative adjusted:** {selected_column}")
                    else:
                        # Fallback to currency info
                        currency_info = metadata.get('currency', 'CNY')
                        multiplier_info = metadata.get('multiplier', 1)
                        if multiplier_info > 1:
                            st.markdown(f"**Processed:** {currency_info} × {multiplier_info}")
                        else:
                            st.markdown(f"**Processed:** {currency_info}")
                with col6:
                    entity_mode = metadata.get('entity_mode', 'unknown')
                    entity_status = "✅" if first_section.get('entity_match', False) else "⚠️"
                    # Remove entity mode display as requested
                    st.markdown(f"**Entity:** {entity_status}")

                if data_rows:
                    structured_data = []
                    for row in data_rows:
                        description = row['description']
                        value = row['value']
                        actual_value = value
                        # Convert datetime objects to strings to avoid Arrow serialization errors
                        if hasattr(actual_value, 'strftime'):  # datetime object
                            formatted_value = actual_value.strftime('%Y-%m-%d')
                        elif isinstance(actual_value, (int, float)):
                            formatted_value = f"{actual_value:,.2f}"
                        else:
                            formatted_value = str(actual_value)
                        structured_data.append({'Description': description, 'Value': formatted_value})

                    df_structured = pd.DataFrame(structured_data)
                    display_df = df_structured[["Description", "Value"]].copy()

                    def highlight_totals(row):
                        if row['Description'].lower() in ['total', 'subtotal']:
                            return ['background-color: rgba(173, 216, 230, 0.3)'] * len(row)
                        return [''] * len(row)

                    styled_df = display_df.style.apply(highlight_totals, axis=1)
                    st.dataframe(styled_df, use_container_width=True)
                else:
                    st.info("No structured data rows found")

                with st.expander("📋 Structured Markdown", expanded=False):
                    st.code(first_section.get('markdown', 'No markdown available'), language='markdown')
                continue

            # Fallback: render raw DataFrame with cleaning
            raw_df = first_section['data'].copy()
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
                # Convert datetime objects to strings to avoid Arrow serialization errors
                raw_df_copy = raw_df.copy()
                for col in raw_df_copy.columns:
                    if raw_df_copy[col].dtype == 'object':
                        raw_df_copy[col] = raw_df_copy[col].astype(str)
                st.dataframe(raw_df_copy, use_container_width=True)
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
    
    st.markdown("#### View Combined Financial Statements by Key")
    
    # High-level debug only
    keys_with_data = [key for key, sections in sections_by_key.items() if sections]
    print(f"DEBUG UI Combined: Processing {len(keys_with_data)} combined keys with data")
    
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

            # Filter sections by selected entity
            filtered_sections = []
            for section in sections:
                section_entity = section.get('entity_name', '')
                if selected_entity.lower() in section_entity.lower() or section_entity.lower() in selected_entity.lower():
                    filtered_sections.append(section)
                else:
                    print(f"🔍 DEBUG Combined: Filtering out section for entity '{section_entity}' (selected: '{selected_entity}')")
            
            if not filtered_sections:
                st.info(f"No sections found for entity '{selected_entity}' in this key.")
                continue
            
            sections = filtered_sections
            print(f"🔍 DEBUG Combined: Showing {len(sections)} sections for entity '{selected_entity}' in key '{key}'")

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
                
                # Metadata summary row
                col1, col2, col3, col4, col5, col6 = st.columns(6)
                with col1:
                    st.markdown(f"**Table:** {metadata['table_name']}")
                with col2:
                    date_value = metadata.get('date')
                    if date_value:
                        try:
                            formatted_date = format_date_to_dd_mmm_yyyy(date_value)
                            st.markdown(f"**Date:** {formatted_date}")
                        except Exception as e:
                            st.markdown(f"**Date:** Error formatting: {e}")
                    else:
                        st.markdown("**Date:** Unknown")
                with col3:
                    currency = metadata.get('currency', 'CNY')
                    st.markdown(f"**Currency:** {currency}")
                with col4:
                    multiplier = metadata.get('multiplier', 1)
                    st.markdown(f"**Multiplier:** {multiplier}x")
                with col5:
                    # Show currency and multiplier info instead of "Value Column"
                    currency_info = metadata.get('currency', 'CNY')
                    multiplier_info = metadata.get('multiplier', 1)
                    if multiplier_info > 1:
                        st.markdown(f"**Processed:** {currency_info} × {multiplier_info}")
                    else:
                        st.markdown(f"**Processed:** {currency_info}")
                with col6:
                    st.markdown(f"**Statement Type:** Combined")
                
                # Display the data table
                if data_rows is not None and len(data_rows) > 0:
                    st.markdown("**📊 Financial Data:**")
                    st.dataframe(data_rows, use_container_width=True)
                else:
                    st.info("No structured data available for this key.")
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
    """Render Income Statement sections UI using existing parsed/cleaned data."""

    st.markdown("#### View Income Statement by Key")
    
    # High-level debug only
    keys_with_data = [key for key, sections in sections_by_key.items() if sections]
    print(f"DEBUG UI IS: Processing {len(keys_with_data)} income statement keys with data")
    
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

            # Filter sections by selected entity
            filtered_sections = []
            for section in sections:
                section_entity = section.get('entity_name', '')
                if selected_entity.lower() in section_entity.lower() or section_entity.lower() in selected_entity.lower():
                    filtered_sections.append(section)
                else:
                    print(f"🔍 DEBUG IS: Filtering out section for entity '{section_entity}' (selected: '{selected_entity}')")
            
            if not filtered_sections:
                st.info(f"No sections found for entity '{selected_entity}' in this key.")
                continue
            
            sections = filtered_sections
            print(f"🔍 DEBUG IS: Showing {len(sections)} sections for entity '{selected_entity}' in key '{key}'")

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

                # Metadata summary row
                col1, col2, col3, col4, col5, col6 = st.columns(6)
                with col1:
                    st.markdown(f"**Table:** {metadata['table_name']}")
                with col2:
                    date_value = metadata.get('date')
                    if date_value:
                        try:
                            formatted_date = format_date_to_dd_mmm_yyyy(date_value)
                            st.markdown(f"**Date:** {formatted_date}")
                        except Exception as e:
                            st.markdown(f"**Date:** Error formatting: {e}")
                    else:
                        st.markdown("**Date:** Unknown")
                with col3:
                    currency = metadata.get('currency', 'CNY')
                    st.markdown(f"**Currency:** {currency}")
                with col4:
                    multiplier = metadata.get('multiplier', 1)
                    st.markdown(f"**Multiplier:** {multiplier}x")
                with col5:
                    # Show Indicative adjusted column info instead of Excel row
                    selected_column = metadata.get('selected_column')
                    if selected_column:
                        st.markdown(f"**Indicative adjusted:** {selected_column}")
                    else:
                        # Fallback to currency info
                        currency_info = metadata.get('currency', 'CNY')
                        multiplier_info = metadata.get('multiplier', 1)
                        if multiplier_info > 1:
                            st.markdown(f"**Processed:** {currency_info} × {multiplier_info}")
                        else:
                            st.markdown(f"**Processed:** {currency_info}")
                with col6:
                    entity_mode = metadata.get('entity_mode', 'unknown')
                    entity_status = "✅" if first_section.get('entity_match', False) else "⚠️"
                    # Remove entity mode display as requested
                    st.markdown(f"**Entity:** {entity_status}")

                if data_rows:
                    structured_data = []
                    for row in data_rows:
                        description = row['description']
                        value = row['value']
                        actual_value = value
                        # Convert datetime objects to strings to avoid Arrow serialization errors
                        if hasattr(actual_value, 'strftime'):  # datetime object
                            formatted_value = actual_value.strftime('%Y-%m-%d')
                        elif isinstance(actual_value, (int, float)):
                            formatted_value = f"{actual_value:,.2f}"
                        else:
                            formatted_value = str(actual_value)
                        structured_data.append({'Description': description, 'Value': formatted_value})

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
                    st.dataframe(styled_df, use_container_width=True)
                else:
                    st.info("No structured data rows found")

                with st.expander("📋 Structured Markdown", expanded=False):
                    st.code(first_section.get('markdown', 'No markdown available'), language='markdown')
                continue

            # Fallback: render raw DataFrame with cleaning
            raw_df = first_section['data'].copy()
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
                st.dataframe(styled_df, use_container_width=True)
            else:
                st.error("No valid columns found after cleaning")
                st.write("**Original DataFrame:**")
                st.dataframe(first_section['data'], use_container_width=True)

            # Show parsed data json if present (debug)
            if 'parsed_data' in first_section:
                with st.expander("🔍 Debug: Parsed Data Structure", expanded=False):
                    st.json(first_section['parsed_data'])