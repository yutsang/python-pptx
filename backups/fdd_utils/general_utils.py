"""
General utilities for FDD application
Utility functions for text processing, similarity calculation, and debugging
"""

import pandas as pd
import streamlit as st
from fdd_utils.data_utils import get_key_display_name

def write_prompt_debug_content(filtered_keys, sections_by_key):
    """
    Write debug content for prompt analysis to a separate file
    This helps with debugging AI prompts without affecting main content files
    """
    try:
        with open("fdd_utils/bs_prompt_debug.md", "w", encoding="utf-8") as f:
            f.write("# AI Prompt Debug Content\n\n")
            f.write(f"**Generated:** {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            f.write(f"**Keys to Process:** {len(filtered_keys)}\n\n")

            for key in filtered_keys:
                if key in sections_by_key and sections_by_key[key]:
                    f.write(f"## {get_key_display_name(key)} ({key})\n\n")

                    # Write data sections
                    for i, section in enumerate(sections_by_key[key]):
                        f.write(f"### Section {i+1}\n\n")
                        if isinstance(section, dict) and 'data' in section:
                            df = section['data']
                            # Clean DataFrame - remove columns that are all None/NaN
                            df_clean = df.dropna(axis=1, how='all')

                            for idx, row in df_clean.iterrows():
                                row_str = " | ".join(str(x) for x in row if pd.notna(x) and str(x).strip() != "None")
                                if row_str:
                                    f.write(f"- {row_str}\n")
                        elif isinstance(section, str):
                            f.write(f"{section}\n")
                        f.write("\n")
                    f.write("\n")

            f.write("---\n*This file is for AI prompt debugging only*\n")
            st.success("✅ Debug content written to fdd_utils/bs_prompt_debug.md")

    except Exception as e:
        st.error(f"Error writing debug content: {e}")

def calculate_content_similarity(text1, text2):
    """
    Calculate similarity percentage between two texts using word-based comparison
    """
    if not text1 or not text2:
        return 0.0

    # Simple word-based similarity
    words1 = set(str(text1).lower().split())
    words2 = set(str(text2).lower().split())

    if not words1 and not words2:
        return 100.0

    intersection = words1.intersection(words2)
    union = words1.union(words2)

    return (len(intersection) / len(union)) * 100 if union else 0.0

def show_text_differences(text1, text2):
    """
    Show differences between two texts with simplified comparison
    """
    if str(text1).strip() == str(text2).strip():
        st.info("✅ No differences found")
        return

    # Split into sentences for comparison
    sentences1 = [s.strip() for s in str(text1).split('.') if s.strip()]
    sentences2 = [s.strip() for s in str(text2).split('.') if s.strip()]

    st.markdown("**🔄 Changes Summary:**")

    # Find added/removed sentences
    added = [s for s in sentences2 if s not in sentences1]
    removed = [s for s in sentences1 if s not in sentences2]

    if added:
        st.markdown("**✅ Added sentences:**")
        for sentence in added[:3]:  # Show first 3
            st.write(f"🟢 + {sentence}")
        if len(added) > 3:
            st.write(f"📝 ... and {len(added) - 3} more additions")

    if removed:
        st.markdown("**❌ Removed sentences:**")
        for sentence in removed[:3]:  # Show first 3
            st.write(f"🔴 - {sentence}")
        if len(removed) > 3:
            st.write(f"📝 ... and {len(removed) - 3} more removals")

    if not added and not removed:
        st.info("📝 Changes are mostly within existing sentences (minor edits)")

def calculate_text_metrics(text):
    """
    Calculate various metrics for text content
    """
    if not text:
        return {
            'characters': 0,
            'words': 0,
            'sentences': 0,
            'avg_word_length': 0
        }

    text_str = str(text)
    words = text_str.split()
    sentences = [s.strip() for s in text_str.split('.') if s.strip()]

    return {
        'characters': len(text_str),
        'words': len(words),
        'sentences': len(sentences),
        'avg_word_length': sum(len(word) for word in words) / len(words) if words else 0
    }

def format_file_size(size_bytes):
    """
    Format file size in human-readable format
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"

def safe_json_load(file_path, default=None):
    """
    Safely load JSON file with error handling
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading JSON from {file_path}: {e}")
        return default if default is not None else {}

def safe_json_dump(data, file_path):
    """
    Safely save JSON file with error handling
    """
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving JSON to {file_path}: {e}")
        return False

def create_backup_file(file_path):
    """
    Create a backup of a file with timestamp
    """
    import shutil
    from pathlib import Path

    if not Path(file_path).exists():
        return False

    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{file_path}.backup_{timestamp}"

    try:
        shutil.copy2(file_path, backup_path)
        return backup_path
    except Exception as e:
        print(f"Error creating backup: {e}")
        return False

def validate_file_exists(file_path, file_description="File"):
    """
    Validate that a file exists and show appropriate message
    """
    from pathlib import Path

    if Path(file_path).exists():
        return True
    else:
        st.error(f"❌ {file_description} not found: {file_path}")
        return False

def get_file_modification_time(file_path):
    """
    Get file modification time in readable format
    """
    from pathlib import Path

    try:
        mtime = Path(file_path).stat().st_mtime
        return pd.Timestamp.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return "Unknown"

def log_processing_step(step_name, details=None):
    """
    Log a processing step with timestamp for debugging
    """
    import logging

    timestamp = pd.Timestamp.now().strftime('%H:%M:%S')
    message = f"[{timestamp}] {step_name}"

    if details:
        message += f" - {details}"

    print(message)

    # If logging is configured, also log to file
    try:
        logging.info(message)
    except Exception:
        pass  # Logging not configured, just print

def validate_text_content(text, min_length=10, max_length=10000):
    """
    Validate text content meets basic requirements
    """
    if not text:
        return False, "Text is empty"

    text_str = str(text).strip()
    if len(text_str) < min_length:
        return False, f"Text too short (minimum {min_length} characters)"

    if len(text_str) > max_length:
        return False, f"Text too long (maximum {max_length} characters)"

    return True, "Valid"
