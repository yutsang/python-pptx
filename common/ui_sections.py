import pandas as pd
import streamlit as st


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

    st.markdown("#### View Table by Key")
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
                    if metadata.get('date'):
                        formatted_date = format_date_to_dd_mmm_yyyy(metadata['date'])
                        st.markdown(f"**Date:** {formatted_date}")
                    else:
                        st.markdown("**Date:** Unknown")
                with col3:
                    st.markdown(f"**Currency:** {metadata['currency_info']}")
                with col4:
                    st.markdown(f"**Multiplier:** {metadata['multiplier']}x")
                with col5:
                    st.markdown(f"**Value Column:** {metadata['value_column']}")
                with col6:
                    st.markdown("**Entity:** ✅" if first_section.get('entity_match', False) else "**Entity:** ⚠️")

                if data_rows:
                    structured_data = []
                    for row in data_rows:
                        description = row['description']
                        value = row['value']
                        actual_value = value
                        if isinstance(actual_value, (int, float)):
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
                st.dataframe(raw_df, use_container_width=True)
            else:
                st.error("No valid columns found after cleaning")
                st.write("**Original DataFrame:**")
                st.dataframe(first_section['data'], use_container_width=True)

            # Show parsed data json if present (debug)
            if 'parsed_data' in first_section:
                with st.expander("🔍 Debug: Parsed Data Structure", expanded=False):
                    st.json(first_section['parsed_data'])


