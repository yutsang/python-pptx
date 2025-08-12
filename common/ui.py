import os
import streamlit as st


def configure_streamlit_page():
    """Centralize Streamlit page setup and config cleanup to avoid local unbound variables."""
    st.set_page_config(page_title="Financial Data Processor", page_icon="📊", layout="wide")
    # Clean any deprecated/invalid options that might be in session
    if 'client.caching' in st.session_state:
        del st.session_state['client.caching']


def select_ai_provider_and_model(config: dict):
    """Render provider/model selection and return (provider, model)."""
    st.markdown("---")
    st.markdown("### 🔧 AI Provider & Model")

    default_provider = config.get('DEFAULT_AI_PROVIDER', 'Server AI')
    providers = ["Open AI", "Local AI", "Server AI"]
    default_index = providers.index(default_provider) if default_provider in providers else 2
    provider = st.selectbox("Select AI Provider", options=providers, index=default_index, key="provider_select")

    openai_models = [config.get('OPENAI_CHAT_MODEL', 'gpt-4o-mini')]
    local_models = config.get('LOCAL_MODELS', ['local-qwen2', 'local-deep-seek', 'local-deep-seek-full'])
    server_models = config.get('SERVER_MODELS', local_models)

    if provider == 'Open AI':
        model = st.selectbox("Model", options=openai_models, key="model_select_openai")
    elif provider == 'Local AI':
        default_idx = local_models.index('local-qwen2') if 'local-qwen2' in local_models else 0
        model = st.selectbox("Model", options=local_models, index=default_idx, key="model_select_local")
    else:
        default_idx = server_models.index('local-qwen2') if 'local-qwen2' in server_models else 0
        model = st.selectbox("Model", options=server_models, index=default_idx, key="model_select_server")

    return provider, model


