import streamlit as st
import pandas as pd
import json
import os
from io import BytesIO
from pathlib import Path

from fdd_utils.app_helpers import derive_entity_parts, get_financial_keys


def _safe_rerun():
    """Compatibility rerun for different Streamlit versions."""
    try:
        # Prefer modern API; silently no-op if unavailable
        if hasattr(st, "rerun"):
            st.rerun()
    except Exception:
        pass


def render_settings_wizard(uploaded_file, load_config_files, get_sections_by_key):
    """Render the 3-step setup wizard in the main area.

    This function manages its own session state and will stop the Streamlit
    script when setup is not completed, to avoid the rest of the app rendering.
    """
    # Ensure settings opens by default
    if 'settings_open' not in st.session_state:
        st.session_state['settings_open'] = True

    show_settings = st.session_state.get('settings_open', False) or not st.session_state.get('setup', {}).get('completed')
    if not show_settings:
        return

    st.markdown("### ⚙️ Setup Wizard")
    st.caption("Configure entities, preview a sheet, and choose keys. After completion, this panel is hidden.")
    setup = st.session_state.setdefault('setup', {"entities": [], "ack_preview": False, "keys_selected": [], "completed": False})
    st.session_state.setdefault('setup_step', 1)
    st.session_state.setdefault('entity_input_nonce', 0)
    step = st.session_state['setup_step']

    # Step 1: Entities
    with st.expander("Step 1: Add Entities", expanded=(step == 1)):
        input_key = f"setup_new_entity_main_{st.session_state['entity_input_nonce']}"
        new_entity = st.text_input("Enter entity full name (e.g., 'Haining Wanpu Limited')", key=input_key)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Add Entity", key="btn_add_entity_main") and new_entity.strip():
                base, suffixes = derive_entity_parts(new_entity.strip())
                setup['entities'].append({"full": new_entity.strip(), "base": base, "suffixes": suffixes})
                st.session_state['setup'] = setup
                st.session_state['entity_input_nonce'] += 1
                _safe_rerun()
        with c2:
            if st.button("Clear Entities", key="btn_clear_entities_main"):
                setup['entities'] = []
                st.session_state['setup'] = setup
        for idx, e in enumerate(list(setup.get('entities', []))):
            cols = st.columns([6, 2])
            with cols[0]:
                st.markdown(f"- {e['full']} → base: `{e['base']}`, suffixes: `{', '.join(e['suffixes']) or '—'}`")
            with cols[1]:
                if st.button("Delete", key=f"btn_del_entity_{idx}"):
                    setup['entities'].pop(idx)
                    st.session_state['setup'] = setup
                    _safe_rerun()
        can_next_1 = bool(setup.get('entities'))
        if st.button("Next", key="btn_next_step1", disabled=not can_next_1):
            st.session_state['setup_step'] = 2

    # Step 2: Select Sheet and Preview
    with st.expander("Step 2: Select Sheet and Preview", expanded=(step == 2)):
        # Resolve Excel source robustly (path or BytesIO)
        sheet_names = []
        try:
            excel_source = None
            if uploaded_file is None:
                # Try project root and module-relative default
                candidates = [
                    "databook.xlsx",
                    str(Path(__file__).resolve().parent.parent / "databook.xlsx"),
                ]
                for p in candidates:
                    if os.path.exists(p):
                        excel_source = p
                        break
            else:
                if hasattr(uploaded_file, "getbuffer"):
                    excel_source = BytesIO(uploaded_file.getbuffer())
                elif hasattr(uploaded_file, "read"):
                    data = uploaded_file.read()
                    # Reset pointer if possible so later consumers can re-read
                    if hasattr(uploaded_file, "seek"):
                        try:
                            uploaded_file.seek(0)
                        except Exception:
                            pass
                    excel_source = BytesIO(data)
                elif hasattr(uploaded_file, "name") and os.path.exists(uploaded_file.name):
                    excel_source = uploaded_file.name
            if excel_source is not None:
                try:
                    # Prefer openpyxl for .xlsx; fall back to auto
                    with pd.ExcelFile(excel_source, engine="openpyxl") as xl:
                        sheet_names = xl.sheet_names
                except Exception:
                    try:
                        with pd.ExcelFile(excel_source) as xl:
                            sheet_names = xl.sheet_names
                    except Exception as inner_e:
                        st.warning(f"Could not read Excel sheets: {inner_e}")
                try:
                    src_label = excel_source if isinstance(excel_source, str) else "uploaded file"
                    st.caption(f"Detected Excel source: {src_label} — {len(sheet_names)} sheet(s)")
                except Exception:
                    pass
        except Exception as e:
            st.warning(f"Could not read Excel file: {e}")
        
        # Show sheet list as dropdown (demo section)
        if sheet_names:
            sel_sheet = st.selectbox("Sheet to test", sheet_names, key="setup_sheet_main")
        else:
            st.warning("No sheets found in Excel file")
            sel_sheet = None
        
        sel_key = st.selectbox("Key to preview", get_financial_keys(), key="setup_key_main")
        if st.button("Generate Preview", key="btn_gen_preview_main") and setup.get('entities') and sel_sheet and uploaded_file is not None:
            _, mapping, _, _ = load_config_files()
            preview = {}
            debug_matches = {}
            for e in setup['entities']:
                sections = get_sections_by_key(
                    uploaded_file=uploaded_file,
                    tab_name_mapping=mapping,
                    entity_name=e['base'],
                    entity_suffixes=e['suffixes'],
                    debug=False,
                )
                filtered = [sec for sec in sections.get(sel_key, []) if sec.get('sheet') == sel_sheet]
                preview[e['full']] = filtered
                debug_matches[e['full']] = list({sec.get('sheet') for k, secs in sections.items() for sec in secs})
            st.session_state['setup_preview'] = preview
            st.session_state['setup_debug_matches'] = debug_matches
        preview = st.session_state.get('setup_preview', {})
        if preview:
            tabs = st.tabs(list(preview.keys()))
            for i, (ename, sections) in enumerate(preview.items()):
                with tabs[i]:
                    if not sections:
                        st.info("No matching table found on this sheet for this entity.")
                        continue
                    first = sections[0]
                    if 'parsed_data' in first and first['parsed_data']:
                        meta = first['parsed_data']['metadata']
                        rows = first['parsed_data']['data']
                        st.markdown(f"**Table:** {meta.get('table_name','')} | **Date:** {meta.get('date','')} | **Currency:** {meta.get('currency_info','')} | **Multiplier:** {meta.get('multiplier','')}x")
                        if rows:
                            df = pd.DataFrame([{ 'Description': r['description'], 'Value': r['value'] } for r in rows])
                            st.dataframe(df, use_container_width=True)
                        show_md = st.checkbox("Show Markdown", key=f"show_md_{i}")
                        if show_md:
                            st.code(first.get('markdown',''), language='markdown')
                    else:
                        st.dataframe(first.get('data'), use_container_width=True)
            debug_matches = st.session_state.get('setup_debug_matches', {})
            show_debug = st.checkbox("Show debug matches (sheets by entity)", key="show_debug_matches")
            if show_debug:
                if not debug_matches:
                    st.caption("No match data yet.")
                for ename, sheets in debug_matches.items():
                    st.markdown(f"- **{ename}** → {', '.join(sorted(sheets)) if sheets else '—'}")
            setup['ack_preview'] = st.checkbox("✅ I confirm the data preview looks correct and I want to proceed to key selection", value=setup.get('ack_preview', False), key="setup_ack_preview_main")
            st.session_state['setup'] = setup
        can_next_2 = bool(setup.get('ack_preview'))
        if st.button("Next", key="btn_next_step2", disabled=not can_next_2):
            st.session_state['setup_step'] = 3

    # Step 3: Key Selection
    with st.expander("Step 3: Select Keys", expanded=(step == 3)):
        # Get available keys from the financial keys function
        available_keys = get_financial_keys()
        selected_keys = setup.get('keys_selected', [])
        
        # Pre-populate with default balance sheet keys if not already set
        if not selected_keys:
            default_bs_keys = [
                "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                "AP", "Taxes payable", "OP", "Capital", "Reserve"
            ]
            selected_keys = [k for k in default_bs_keys if k in available_keys]
            setup['keys_selected'] = selected_keys
            st.session_state['setup'] = setup
        
        st.markdown("**Available keys (click to select/deselect):**")
        
        # Create a grid of checkboxes to show all keys
        cols = st.columns(3)  # 3 columns for better layout
        new_selected_keys = selected_keys.copy()
        
        for i, key in enumerate(available_keys):
            col_idx = i % 3
            with cols[col_idx]:
                is_selected = key in selected_keys
                if st.checkbox(
                    f"{key} {'✓' if is_selected else ''}", 
                    value=is_selected, 
                    key=f"key_{key}",
                    help=f"Select {key} for your presentation"
                ):
                    if key not in new_selected_keys:
                        new_selected_keys.append(key)
                else:
                    if key in new_selected_keys:
                        new_selected_keys.remove(key)
        
        # Update the setup with new selection
        if new_selected_keys != selected_keys:
            setup['keys_selected'] = new_selected_keys
            st.session_state['setup'] = setup
            st.rerun()
        
        can_next_3 = bool(setup.get('keys_selected'))
        if st.button("Next", key="btn_next_step3", disabled=not can_next_3):
            st.session_state['setup_step'] = 4

    # Step 4: Key Ordering
    with st.expander("Step 4: Arrange Key Order", expanded=(step == 4)):
        selected_keys = setup.get('keys_selected', [])
        if selected_keys:
            st.markdown("**Arrange the order of your selected keys:**")
            for i, key in enumerate(selected_keys):
                col1, col2, col3 = st.columns([8, 1, 1])
                with col1:
                    st.markdown(f"{i+1}. {key}")
                with col2:
                    if st.button("↑", key=f"up_{i}", disabled=(i==0)):
                        if i > 0:
                            selected_keys[i-1], selected_keys[i] = selected_keys[i], selected_keys[i-1]
                            setup['keys_selected'] = selected_keys
                            st.session_state['setup'] = setup
                            st.rerun()
                with col3:
                    if st.button("↓", key=f"down_{i}", disabled=(i==len(selected_keys)-1)):
                        if i < len(selected_keys)-1:
                            selected_keys[i], selected_keys[i+1] = selected_keys[i+1], selected_keys[i]
                            setup['keys_selected'] = selected_keys
                            st.session_state['setup'] = setup
                            st.rerun()
        else:
            st.markdown("*No keys to arrange*")
        
        ready = bool(setup.get('entities')) and bool(setup.get('ack_preview')) and bool(setup.get('keys_selected'))
        
        # Show completion status
        if ready:
            st.success("✅ All setup steps completed! You can now complete the setup.")
        else:
            missing_items = []
            if not setup.get('entities'):
                missing_items.append("entities")
            if not setup.get('ack_preview'):
                missing_items.append("sheet preview")
            if not setup.get('keys_selected'):
                missing_items.append("key selection")
            st.warning(f"⚠️ Please complete: {', '.join(missing_items)}")
        
        if st.button("Complete Setup", type="primary", key="btn_complete_setup_main", disabled=not ready):
            setup['completed'] = True
            st.session_state['setup'] = setup
            st.session_state['settings_open'] = False
            st.session_state['setup_step'] = 1
            st.session_state['setup_last_hash'] = json.dumps(st.session_state.get('setup'), sort_keys=True)
            st.success("Setup completed. Settings hidden; proceed to normal operations.")
            st.rerun()


