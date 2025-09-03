"""
Display utilities for FDD application
Functions for displaying content, comparisons, and UI elements
"""

import streamlit as st
import re
import difflib
from fdd_utils.content_utils import clean_content_quotes

def display_ai_content_by_key(key, agent_choice, language='English'):
    """
    Display AI content based on the financial key using actual AI processing
    """
    try:
        import re

        # Get AI data from session state
        ai_data = st.session_state.get('ai_data')
        if not ai_data:
            st.info("No AI data available. Please process with AI first.")
            return

        sections_by_key = ai_data['sections_by_key']
        pattern = ai_data['pattern']
        mapping = ai_data['mapping']
        config = ai_data['config']
        entity_name = ai_data['entity_name']
        entity_keywords = ai_data['entity_keywords']
        mode = ai_data.get('mode', 'AI Mode')

        # Get sections for this key
        sections = sections_by_key.get(key, [])
        if not sections:
            st.info(f"No data found for {get_key_display_name(key)}")
            return

        # Show processing status
        with st.spinner(f"🤖 Processing {get_key_display_name(key)} with {agent_choice}..."):

            # Process the key with the selected agent
            if agent_choice == 'Agent 1':
                result = run_single_agent_1(key, sections, pattern, entity_name, entity_keywords, language)
            elif agent_choice == 'Agent 2':
                result = run_single_agent_2(key, sections, pattern, entity_name, entity_keywords, language)
            elif agent_choice == 'Agent 3':
                result = run_single_agent_3(key, sections, pattern, entity_name, entity_keywords, language)
            else:
                st.error(f"Unknown agent choice: {agent_choice}")
                return

        if result and 'content' in result:
            content = result['content']

            # Clean up quotes if needed
            content = clean_content_quotes(content)

            # Display the content
            st.markdown(f"### 🤖 {agent_choice} Result for {get_key_display_name(key)}")
            st.markdown(f"**Entity:** {entity_name}")
            st.markdown(f"**Language:** {language}")
            st.markdown(f"**Mode:** {mode}")

            # Content statistics
            st.markdown("**📊 Content Statistics:**")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Characters", len(str(content)))
            with col2:
                st.metric("Words", len(str(content).split()))
            with col3:
                st.metric("Sentences", len(str(content).split('.')))

            # Content display
            st.markdown("**📝 Generated Content:**")
            with st.container():
                st.markdown(content)

            # Pattern compliance check
            if pattern and key in pattern:
                st.markdown("**🎯 Pattern Compliance:**")
                key_pattern = pattern[key]
                compliance_score = calculate_pattern_compliance(content, key_pattern)
                st.metric("Pattern Compliance", f"{compliance_score:.1f}%")

                if compliance_score < 70:
                    st.warning("⚠️ Content may not fully comply with the required pattern")

        else:
            st.error(f"Failed to generate content for {get_key_display_name(key)} with {agent_choice}")

    except Exception as e:
        st.error(f"Error displaying AI content: {e}")

def display_ai_prompt_by_key(key, agent_choice, language='English'):
    """
    Display AI prompt for the financial key using dynamic prompts from configuration
    """
    try:
        # Load prompts from configuration
        config, mapping, pattern, prompts = load_config_files()

        if not prompts:
            st.error("❌ Failed to load prompts configuration")
            return

        # Get system prompts from configuration - use language-specific prompts
        language_key = 'chinese' if language == '中文' else 'english'
        system_prompts = prompts.get('system_prompts', {}).get(language_key, {})

        # Get user prompts from configuration
        user_prompts_config = prompts.get('user_prompts', {})
        generic_prompt_config = prompts.get('generic_prompt', {})

        # Get AI data for context
        ai_data = st.session_state.get('ai_data', {})
        entity_name = ai_data.get('entity_name', 'Unknown Entity')

        # Import centralized prompt generation
        from fdd_utils.prompt_templates import generate_dynamic_user_prompt

        # Get the prompts for this key and agent
        system_prompt = system_prompts.get(agent_choice, system_prompts.get('Agent 1', ''))
        user_prompt_config = user_prompts_config.get(key, generic_prompt_config)
        user_prompt = generate_dynamic_user_prompt(key, user_prompt_config, entity_name, get_key_display_name(key))

        if user_prompt:
            st.markdown("### 🤖 AI Prompt Configuration")
            st.markdown(f"**Agent:** {agent_choice}")
            st.markdown(f"**Financial Key:** {get_key_display_name(key)}")

            # Collapsible prompt sections
            prompt_expander = st.expander("📝 View AI Prompts", expanded=False)
            with prompt_expander:
                st.markdown("#### 📋 System Prompt")
                st.code(system_prompt, language="text")

                st.markdown("#### 💬 User Prompt")
                st.code(user_prompt, language="text")

            # Get AI data for debug information
            ai_data = st.session_state.get('ai_data', {})
            sections_by_key = ai_data.get('sections_by_key', {})

            st.markdown("#### 📊 Debug Information")

            # Pattern information
            pattern = ai_data.get('pattern', {})
            key_patterns = pattern.get(key, {})
            if key_patterns:
                st.markdown("**📋 Available Patterns:**")
                for pattern_name, pattern_text in key_patterns.items():
                    with st.expander(f"Pattern: {pattern_name}", expanded=False):
                        st.code(pattern_text, language="text")

            # Data sections information
            sections = sections_by_key.get(key, [])
            if sections:
                st.markdown("**📄 Data Sections Available:**")
                st.metric("Total Sections", len(sections))

                # Show sample data
                if len(sections) > 0:
                    with st.expander("🔍 Sample Data", expanded=False):
                        sample_text = sections[0][:500] + "..." if len(sections[0]) > 500 else sections[0]
                        st.code(sample_text, language="text")

        else:
            st.info(f"No AI prompt template available for {get_key_display_name(key)}")
            from fdd_utils.ui_config import get_generic_prompt_template
            system_prompt = system_prompts.get(agent_choice, system_prompts.get('Agent 1', ''))
            key_display_name = get_key_display_name(key)
            generic_prompt = get_generic_prompt_template().format(
                system_prompt=system_prompt,
                key_display_name=key_display_name
            )
            st.markdown(generic_prompt)

    except Exception as e:
        st.error(f"Error generating AI prompt for {key}: {e}")

def display_sequential_agent_results(key, filtered_keys, ai_data):
    """Display sequential agent results for a specific key"""
    try:
        st.markdown(f"### 🔄 Sequential Agent Results for {get_key_display_name(key)}")

        agent_states = st.session_state.get('agent_states', {})
        agent1_results = agent_states.get('agent1_results', {})
        agent2_results = agent_states.get('agent2_results', {})
        agent3_results = agent_states.get('agent3_results', {})

        # Get content for each agent
        agent1_content = get_agent_content(agent1_results, key)
        agent2_content = get_agent_content(agent2_results, key)
        agent3_content = get_agent_content(agent3_results, key)

        # Display results in tabs
        tab1, tab2, tab3 = st.tabs(["🤖 Agent 1", "🎯 Agent 2", "✅ Agent 3"])

        with tab1:
            display_agent_result("Agent 1", agent1_content, agent_states.get('agent1_success', False))

        with tab2:
            display_agent_result("Agent 2", agent2_content, agent_states.get('agent2_success', False))

        with tab3:
            display_agent_result("Agent 3", agent3_content, agent_states.get('agent3_success', False))

        # Show comparison if all agents have results
        if agent1_content and agent2_content and agent3_content:
            st.markdown("---")
            display_step_by_step_comparison(key, agent1_content, agent2_content, agent3_content, agent_states)

    except Exception as e:
        st.error(f"Error displaying sequential agent results: {e}")

def display_before_after_comparison(key, before_content, after_content, agent_states):
    """Display before (AI1) vs after (AI3) comparison with visual diff"""
    st.markdown("#### 📊 Before vs After Comparison")

    # Status indicators
    col1, col2, col3 = st.columns(3)
    with col1:
        agent1_success = agent_states.get('agent1_success', False)
        st.metric("Agent 1", "✅ Success" if agent1_success else "❌ Failed")
    with col2:
        changes_made = before_content != after_content
        st.metric("Changes Made", "✅ Yes" if changes_made else "➖ No")
    with col3:
        agent3_success = agent_states.get('agent3_success', False)
        st.metric("Agent 3", "✅ Success" if agent3_success else "❌ Failed")

    # Side-by-side comparison
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("##### 📝 **BEFORE** (Agent 1 - Original)")
        if before_content:
            st.markdown(f"**Length:** {len(str(before_content))} characters, {len(str(before_content).split())} words")
            with st.container():
                st.markdown(before_content)
        else:
            st.warning("No original content available")

    with col2:
        st.markdown("##### 🎯 **AFTER** (Agent 3 - Final)")
        if after_content:
            st.markdown(f"**Length:** {len(str(after_content))} characters, {len(str(after_content).split())} words")
            with st.container():
                st.markdown(after_content)
        else:
            st.warning("No final content available")

    # Change analysis
    if before_content and after_content:
        st.markdown("---")
        st.markdown("#### 🔄 Change Analysis")

        # Calculate similarities
        similarity = calculate_content_similarity(str(before_content), str(after_content))

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Text Similarity", f"{similarity:.1f}%")
        with col2:
            added_words = len(str(after_content).split()) - len(str(before_content).split())
            st.metric("Words Added", f"{added_words:+d}")
        with col3:
            added_chars = len(str(after_content)) - len(str(before_content))
            st.metric("Characters Added", f"{added_chars:+d}")

        # Show differences if content is different
        if before_content != after_content:
            st.markdown("**📋 Key Differences:**")
            diff_html = show_text_differences(str(before_content), str(after_content))
            st.markdown(diff_html, unsafe_allow_html=True)

def display_step_by_step_comparison(key, agent1_content, agent2_content, agent3_content, agent_states):
    """Display step-by-step comparison of all three agents"""
    st.markdown("#### 📈 Step-by-Step Evolution")

    # Create tabs for each comparison
    tab1, tab2 = st.tabs(["Agent 1 → 2", "Agent 2 → 3"])

    with tab1:
        st.markdown("**Agent 1 → Agent 2 (Proofreading)**")
        display_before_after_comparison(key, agent1_content, agent2_content, agent_states)

    with tab2:
        st.markdown("**Agent 2 → Agent 3 (Validation)**")
        display_before_after_comparison(key, agent2_content, agent3_content, agent_states)

def display_validation_comparison(key, agent2_content, agent3_content, agent2_results, agent3_results):
    """Display validation comparison between Agent 2 and Agent 3"""
    st.markdown("#### 🎯 Validation Comparison")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Agent 2 (Proofread)**")
        if agent2_content:
            st.metric("Status", "✅ Active")
            st.text_area("Content", agent2_content, height=100, disabled=True)
        else:
            st.metric("Status", "❌ No Content")

    with col2:
        st.markdown("**Agent 3 (Validated)**")
        if agent3_content:
            st.metric("Status", "✅ Active")
            st.text_area("Content", agent3_content, height=100, disabled=True)
        else:
            st.metric("Status", "❌ No Content")

    # Show differences
    if agent2_content and agent3_content:
        similarity = calculate_content_similarity(agent2_content, agent3_content)
        st.metric("Content Similarity", f"{similarity:.1f}%")

        if similarity < 95:
            st.markdown("**Differences detected:**")
            diff_html = show_text_differences(agent2_content, agent3_content)
            st.markdown(diff_html, unsafe_allow_html=True)

def calculate_content_similarity(text1, text2):
    """Calculate similarity between two text contents"""
    try:
        # Use difflib for similarity calculation
        seq = difflib.SequenceMatcher(None, str(text1), str(text2))
        return seq.ratio() * 100
    except Exception:
        return 0.0

def show_text_differences(text1, text2):
    """Show differences between two texts with HTML formatting"""
    try:
        # Create unified diff
        diff = list(difflib.unified_diff(
            str(text1).splitlines(keepends=True),
            str(text2).splitlines(keepends=True),
            fromfile='Before',
            tofile='After',
            lineterm=''
        ))

        if not diff:
            return "<p style='color: green;'>✅ No differences found</p>"

        # Convert to HTML
        html_diff = "<div style='font-family: monospace; background: #f8f9fa; padding: 10px; border-radius: 5px;'>"
        for line in diff[2:]:  # Skip header lines
            if line.startswith('+'):
                html_diff += f"<span style='color: green;'>+ {line[1:]}</span><br>"
            elif line.startswith('-'):
                html_diff += f"<span style='color: red;'>- {line[1:]}</span><br>"
            elif line.startswith('@@'):
                html_diff += f"<span style='color: blue; font-weight: bold;'>{line}</span><br>"
            else:
                html_diff += f"{line}<br>"
        html_diff += "</div>"

        return html_diff

    except Exception as e:
        return f"<p style='color: red;'>Error generating diff: {e}</p>"

# Helper functions
def get_key_display_name(key):
    """Get display name for a financial key"""
    from fdd_utils.data_utils import get_key_display_name as get_display_name
    return get_display_name(key)

def load_config_files():
    """Load configuration files"""
    from fdd_utils.data_utils import load_config_files as load_configs
    return load_configs()

def get_agent_content(results_dict, key):
    """Get content for a specific key from agent results"""
    if key in results_dict:
        result = results_dict[key]
        if isinstance(result, dict):
            return result.get('content', result.get('corrected_content', ''))
        return str(result)
    return None

def display_agent_result(agent_name, content, success_status):
    """Display result for a specific agent"""
    if success_status:
        st.success(f"✅ {agent_name} completed successfully")
    else:
        st.error(f"❌ {agent_name} failed")

    if content:
        st.markdown(f"**Content Length:** {len(str(content))} characters")
        with st.expander(f"View {agent_name} Content", expanded=False):
            st.markdown(content)
    else:
        st.warning(f"No content generated by {agent_name}")

def calculate_pattern_compliance(content, pattern):
    """Calculate how well content complies with a pattern"""
    try:
        if not pattern or not content:
            return 0.0

        # Simple compliance check based on key terms
        pattern_lower = str(pattern).lower()
        content_lower = str(content).lower()

        # Count matching keywords
        keywords = ['balance', 'represents', 'total', 'amount', 'value']
        matches = sum(1 for keyword in keywords if keyword in content_lower)

        return (matches / len(keywords)) * 100
    except Exception:
        return 0.0

# Placeholder functions for agent processing (these would be imported from agent_utils)
def run_single_agent_1(key, sections, pattern, entity_name, entity_keywords, language):
    """Run single Agent 1 processing - placeholder"""
    # This would be imported from agent_utils when implemented
    return {"content": f"Agent 1 processed {key}", "success": True}

def run_single_agent_2(key, sections, pattern, entity_name, entity_keywords, language):
    """Run single Agent 2 processing - placeholder"""
    # This would be imported from agent_utils when implemented
    return {"content": f"Agent 2 processed {key}", "success": True}

def run_single_agent_3(key, sections, pattern, entity_name, entity_keywords, language):
    """Run single Agent 3 processing - placeholder"""
    # This would be imported from agent_utils when implemented
    return {"content": f"Agent 3 processed {key}", "success": True}
