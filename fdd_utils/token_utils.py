"""
Token counting and management utilities for AI requests
Prevents token limit errors and provides visibility into token usage

NO ADDITIONAL DEPENDENCIES REQUIRED - Uses character-based estimation by default
Optional: Install tiktoken for accurate token counting (pip install tiktoken)
"""

import json
from typing import Dict, List, Tuple, Optional
from datetime import datetime

# Try to import tiktoken, but don't fail if not available
try:
    import tiktoken
    TIKTOKEN_AVAILABLE = True
except ImportError:
    TIKTOKEN_AVAILABLE = False
    tiktoken = None


class TokenCounter:
    """Utility class for counting and managing tokens in AI requests"""
    
    def __init__(self, model_name: str = "gpt-4"):
        """
        Initialize token counter for a specific model
        
        Args:
            model_name: Model name for encoding (default: gpt-4)
                       Supports: gpt-4, gpt-3.5-turbo, text-embedding-ada-002, etc.
        
        Note:
            If tiktoken is not installed, uses character-based estimation:
            - English: 1 token ≈ 4 characters
            - Chinese: 1 token ≈ 1.8 characters
            Install tiktoken for accurate counting: pip install tiktoken
        """
        self.model_name = model_name
        self.token_logs = []
        self.use_accurate_counting = False
        
        if TIKTOKEN_AVAILABLE:
            try:
                # Try model-specific encoding first
                self.encoding = tiktoken.encoding_for_model(model_name)
                self.use_accurate_counting = True
                # Silent success - no need to print
            except KeyError:
                # Fallback to cl100k_base encoding (silent)
                try:
                    self.encoding = tiktoken.get_encoding("cl100k_base")
                    self.use_accurate_counting = True
                    # Silent fallback - works fine
                except:
                    self.encoding = None
                    # Silent fallback to character-based
        else:
            self.encoding = None
            # Silent - only print once at app startup if needed
    
    def count_tokens(self, text: str) -> int:
        """
        Count tokens in a text string
        
        Args:
            text: Input text to count tokens
            
        Returns:
            Number of tokens in the text (estimated if tiktoken not available)
        """
        if not text:
            return 0
        
        text = str(text)
        
        # Use accurate tiktoken counting if available
        if self.use_accurate_counting and self.encoding:
            return len(self.encoding.encode(text))
        
        # Fallback: Character-based estimation
        # English: ~4 chars per token, Chinese: ~1.8 chars per token
        char_count = len(text)
        
        # Detect if text contains significant Chinese characters
        chinese_char_count = sum(1 for char in text if '\u4e00' <= char <= '\u9fff')
        chinese_ratio = chinese_char_count / char_count if char_count > 0 else 0
        
        if chinese_ratio > 0.3:
            # Mostly Chinese text: ~1.8 characters per token
            estimated_tokens = int(char_count / 1.8)
        else:
            # English/mixed text: ~4 characters per token
            estimated_tokens = int(char_count / 4)
        
        return estimated_tokens
    
    def count_messages_tokens(self, messages: List[Dict[str, str]]) -> int:
        """
        Count tokens in a list of messages (chat format)
        
        Args:
            messages: List of message dicts with 'role' and 'content'
            
        Returns:
            Total number of tokens including message formatting overhead
        """
        # Token counting for chat messages includes:
        # - 3 tokens per message overhead
        # - 1 token for message role
        # - tokens in message content
        # - 3 tokens for assistant reply priming
        
        num_tokens = 3  # Base overhead for reply priming
        
        for message in messages:
            num_tokens += 3  # Message overhead
            for key, value in message.items():
                num_tokens += self.count_tokens(str(value))
        
        return num_tokens
    
    def estimate_cost(self, input_tokens: int, output_tokens: int, 
                     model: str = None) -> Tuple[float, str]:
        """
        Estimate API cost based on token counts
        
        Args:
            input_tokens: Number of input tokens
            output_tokens: Number of output tokens
            model: Model name (uses self.model_name if not provided)
            
        Returns:
            Tuple of (cost in USD, currency string)
        """
        model = model or self.model_name
        
        # Pricing as of October 2024 (update as needed)
        pricing = {
            "gpt-4": (0.03, 0.06),  # per 1K tokens (input, output)
            "gpt-4-turbo": (0.01, 0.03),
            "gpt-3.5-turbo": (0.0015, 0.002),
            "deepseek-chat": (0.0014, 0.0028),  # DeepSeek pricing
            "deepseek-coder": (0.0014, 0.0028),
        }
        
        # Default pricing if model not found
        input_price, output_price = pricing.get(model, (0.002, 0.004))
        
        input_cost = (input_tokens / 1000) * input_price
        output_cost = (output_tokens / 1000) * output_price
        total_cost = input_cost + output_cost
        
        return total_cost, "USD"
    
    def check_token_limit(self, text: str, limit: int = 20000, 
                         buffer: int = 1000) -> Tuple[bool, int, str]:
        """
        Check if text exceeds token limit
        
        Args:
            text: Input text to check
            limit: Maximum token limit
            buffer: Safety buffer to leave for response
            
        Returns:
            Tuple of (is_within_limit, token_count, message)
        """
        token_count = self.count_tokens(text)
        effective_limit = limit - buffer
        
        if token_count <= effective_limit:
            return True, token_count, f"✅ Within limit ({token_count}/{effective_limit} tokens)"
        else:
            excess = token_count - effective_limit
            return False, token_count, f"❌ Exceeds limit by {excess} tokens ({token_count}/{effective_limit})"
    
    def truncate_to_limit(self, text: str, limit: int = 20000, 
                         preserve_start: bool = True) -> Tuple[str, int, int]:
        """
        Truncate text to fit within token limit
        
        Args:
            text: Input text to truncate
            limit: Maximum token limit
            preserve_start: If True, keep beginning of text; if False, keep end
            
        Returns:
            Tuple of (truncated_text, original_tokens, final_tokens)
        """
        tokens = self.encoding.encode(text)
        original_count = len(tokens)
        
        if original_count <= limit:
            return text, original_count, original_count
        
        # Truncate tokens
        if preserve_start:
            truncated_tokens = tokens[:limit]
        else:
            truncated_tokens = tokens[-limit:]
        
        truncated_text = self.encoding.decode(truncated_tokens)
        
        return truncated_text, original_count, len(truncated_tokens)
    
    def log_token_usage(self, operation: str, key: str, input_tokens: int, 
                       output_tokens: int, agent: str = "unknown", 
                       metadata: Dict = None):
        """
        Log token usage for an operation
        
        Args:
            operation: Operation name (e.g., "content_generation", "validation")
            key: Financial key being processed
            input_tokens: Number of input tokens
            output_tokens: Number of output tokens
            agent: Agent name (e.g., "agent1", "agent2")
            metadata: Additional metadata to log
        """
        cost, currency = self.estimate_cost(input_tokens, output_tokens)
        
        log_entry = {
            "timestamp": datetime.now().isoformat(),
            "operation": operation,
            "key": key,
            "agent": agent,
            "input_tokens": input_tokens,
            "output_tokens": output_tokens,
            "total_tokens": input_tokens + output_tokens,
            "estimated_cost": cost,
            "currency": currency,
            "model": self.model_name,
            "metadata": metadata or {}
        }
        
        self.token_logs.append(log_entry)
        
        return log_entry
    
    def get_usage_summary(self) -> Dict:
        """
        Get summary of token usage across all logged operations
        
        Returns:
            Dict with usage statistics
        """
        if not self.token_logs:
            return {
                "total_operations": 0,
                "total_input_tokens": 0,
                "total_output_tokens": 0,
                "total_tokens": 0,
                "total_cost": 0,
                "currency": "USD"
            }
        
        total_input = sum(log["input_tokens"] for log in self.token_logs)
        total_output = sum(log["output_tokens"] for log in self.token_logs)
        total_cost = sum(log["estimated_cost"] for log in self.token_logs)
        
        # Group by operation
        by_operation = {}
        for log in self.token_logs:
            op = log["operation"]
            if op not in by_operation:
                by_operation[op] = {
                    "count": 0,
                    "input_tokens": 0,
                    "output_tokens": 0,
                    "cost": 0
                }
            by_operation[op]["count"] += 1
            by_operation[op]["input_tokens"] += log["input_tokens"]
            by_operation[op]["output_tokens"] += log["output_tokens"]
            by_operation[op]["cost"] += log["estimated_cost"]
        
        # Group by key
        by_key = {}
        for log in self.token_logs:
            key = log["key"]
            if key not in by_key:
                by_key[key] = {
                    "count": 0,
                    "input_tokens": 0,
                    "output_tokens": 0,
                    "cost": 0
                }
            by_key[key]["count"] += 1
            by_key[key]["input_tokens"] += log["input_tokens"]
            by_key[key]["output_tokens"] += log["output_tokens"]
            by_key[key]["cost"] += log["estimated_cost"]
        
        return {
            "total_operations": len(self.token_logs),
            "total_input_tokens": total_input,
            "total_output_tokens": total_output,
            "total_tokens": total_input + total_output,
            "total_cost": total_cost,
            "currency": "USD",
            "by_operation": by_operation,
            "by_key": by_key,
            "detailed_logs": self.token_logs
        }
    
    def save_logs(self, filepath: str):
        """Save token logs to JSON file"""
        summary = self.get_usage_summary()
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        print(f"✅ Token logs saved to {filepath}")
    
    def split_large_content(self, content: str, max_tokens: int = 15000, 
                           overlap_tokens: int = 500) -> List[str]:
        """
        Split large content into chunks that fit within token limits
        
        Args:
            content: Content to split
            max_tokens: Maximum tokens per chunk
            overlap_tokens: Number of overlapping tokens between chunks
            
        Returns:
            List of content chunks
        """
        # Use accurate splitting if tiktoken available
        if self.use_accurate_counting and self.encoding:
            tokens = self.encoding.encode(content)
            total_tokens = len(tokens)
            
            if total_tokens <= max_tokens:
                return [content]
            
            chunks = []
            start = 0
            
            while start < total_tokens:
                end = min(start + max_tokens, total_tokens)
                chunk_tokens = tokens[start:end]
                chunk_text = self.encoding.decode(chunk_tokens)
                chunks.append(chunk_text)
                
                # Move start position with overlap
                start = end - overlap_tokens
                
                if start >= total_tokens:
                    break
            
            return chunks
        
        # Fallback: Character-based splitting
        total_tokens = self.count_tokens(content)
        
        if total_tokens <= max_tokens:
            return [content]
        
        # Estimate characters per chunk based on token estimation
        chinese_ratio = sum(1 for c in content if '\u4e00' <= c <= '\u9fff') / len(content)
        chars_per_token = 1.8 if chinese_ratio > 0.3 else 4
        
        max_chars = int(max_tokens * chars_per_token)
        overlap_chars = int(overlap_tokens * chars_per_token)
        
        chunks = []
        start = 0
        content_len = len(content)
        
        while start < content_len:
            end = min(start + max_chars, content_len)
            chunk = content[start:end]
            chunks.append(chunk)
            
            start = end - overlap_chars
            if start >= content_len:
                break
        
        return chunks


def format_token_info(token_count: int, limit: int = 20000) -> str:
    """
    Format token information for display
    
    Args:
        token_count: Current token count
        limit: Token limit
        
    Returns:
        Formatted string with color indicators
    """
    percentage = (token_count / limit) * 100
    
    if percentage < 50:
        status = "✅"
    elif percentage < 80:
        status = "⚠️"
    else:
        status = "🔴"
    
    return f"{status} {token_count:,} / {limit:,} tokens ({percentage:.1f}%)"


def get_token_counter(model_name: str = "gpt-4") -> TokenCounter:
    """
    Factory function to get a TokenCounter instance
    
    Args:
        model_name: Model name for encoding
        
    Returns:
        TokenCounter instance
    """
    return TokenCounter(model_name)


# Example usage
if __name__ == "__main__":
    print("="*60)
    print("Token Counter Test")
    print("="*60)
    
    counter = TokenCounter("gpt-4")
    
    print(f"\nAccurate counting available: {counter.use_accurate_counting}")
    
    # Test English text
    english_text = "This is a sample text for token counting. " * 10
    token_count = counter.count_tokens(english_text)
    print(f"\nEnglish text ({len(english_text)} chars): {token_count} tokens")
    
    # Test Chinese text
    chinese_text = "这是一个用于计数令牌的示例文本。" * 10
    token_count_cn = counter.count_tokens(chinese_text)
    print(f"Chinese text ({len(chinese_text)} chars): {token_count_cn} tokens")
    
    # Check limit
    is_ok, count, message = counter.check_token_limit(english_text, limit=100)
    print(f"\n{message}")
    
    # Log usage
    counter.log_token_usage("test", "sample_key", token_count, 50, "agent1")
    
    # Get summary
    summary = counter.get_usage_summary()
    print(f"\nTotal cost: ${summary['total_cost']:.4f}")
    
    print("\n" + "="*60)
    if not counter.use_accurate_counting:
        print("ℹ️  Install tiktoken for accurate counting: pip install tiktoken")
    print("="*60)

