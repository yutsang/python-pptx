"""
Content cleanup utilities to remove placeholders and improve output quality
"""

import re


def remove_placeholders(text):
    """
    Remove all placeholder patterns from text that AI didn't fill in
    
    Removes:
    - {placeholder} patterns
    - [placeholder] patterns
    - xxx or XXX patterns
    - Common template artifacts
    
    Args:
        text: Input text with potential placeholders
        
    Returns:
        Cleaned text without placeholders
    """
    if not text:
        return text
    
    # Remove {placeholder} patterns
    text = re.sub(r'\{[^}]+\}', '', text)
    
    # Remove [placeholder] patterns  
    text = re.sub(r'\[[^\]]+\]', '', text)
    
    # Remove xxx or XXX as standalone words
    text = re.sub(r'\bxxx\b', '', text, flags=re.IGNORECASE)
    
    # Remove common placeholder phrases
    placeholder_phrases = [
        r'Which Entity\?',
        r'Which Entity',
        r'Project_Name',
        r'Project Name',
        r'PLACEHOLDER',
        r'TBD',
        r'TODO',
        r'FIXME',
        r'\?\?\?',
    ]
    
    for phrase in placeholder_phrases:
        text = re.sub(phrase, '', text, flags=re.IGNORECASE)
    
    # Clean up formatting issues from removals
    # Remove double spaces
    text = re.sub(r'\s+', ' ', text)
    
    # Remove space before punctuation
    text = re.sub(r'\s+([,.;:!?])', r'\1', text)
    
    # Remove orphaned prepositions
    text = re.sub(r'\b(of|from|to|at|in|on|with|for)\s+([,.;:])', r'\2', text)
    
    # Remove empty parentheses or brackets
    text = re.sub(r'\(\s*\)', '', text)
    text = re.sub(r'\[\s*\]', '', text)
    
    # Clean up multiple punctuation
    text = re.sub(r'[,.;:]+', lambda m: m.group()[0], text)
    
    return text.strip()


def validate_no_placeholders(text):
    """
    Check if text still contains placeholder patterns
    
    Returns:
        (bool, list): (has_placeholders, list_of_found_placeholders)
    """
    found_placeholders = []
    
    # Check for {placeholder}
    curly_placeholders = re.findall(r'\{[^}]+\}', text)
    found_placeholders.extend(curly_placeholders)
    
    # Check for [placeholder]
    bracket_placeholders = re.findall(r'\[[^\]]+\]', text)
    found_placeholders.extend(bracket_placeholders)
    
    # Check for xxx
    if re.search(r'\bxxx\b', text, re.IGNORECASE):
        found_placeholders.append('xxx')
    
    return len(found_placeholders) > 0, found_placeholders


def clean_financial_commentary(text):
    """
    Clean financial commentary text comprehensively
    
    - Removes placeholders
    - Fixes formatting
    - Ensures professional output
    
    Args:
        text: Raw commentary text
        
    Returns:
        Cleaned professional commentary
    """
    # Remove placeholders
    text = remove_placeholders(text)
    
    # Ensure proper capitalization at sentence starts
    sentences = re.split(r'([.!?]\s+)', text)
    cleaned_sentences = []
    
    for i, sentence in enumerate(sentences):
        if sentence and not sentence.strip() in ['.', '!', '?']:
            # Capitalize first letter of sentence
            sentence = sentence[0].upper() + sentence[1:] if len(sentence) > 1 else sentence.upper()
        cleaned_sentences.append(sentence)
    
    text = ''.join(cleaned_sentences)
    
    # Remove any remaining artifacts
    # Remove "Pattern X:" if AI left it
    text = re.sub(r'Pattern\s+\d+:\s*', '', text, flags=re.IGNORECASE)
    
    # Remove quotation marks around entire paragraphs
    text = text.strip('"\'')
    
    return text.strip()


# Example usage
if __name__ == "__main__":
    test_text = """
    balance as at {date} represented CNY[Amount] of cash at bank from {Which Entity?}.
    The xxx were located in {location} with {Project_Name} totaling [placeholder] amount.
    Management indicated no restrictions.
    """
    
    print("Original:")
    print(test_text)
    
    print("\nCleaned:")
    cleaned = clean_financial_commentary(test_text)
    print(cleaned)
    
    has_placeholders, found = validate_no_placeholders(cleaned)
    print(f"\nHas placeholders: {has_placeholders}")
    if found:
        print(f"Found: {found}")

