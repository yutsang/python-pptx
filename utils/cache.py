import json
import pickle
import hashlib
import os
import time
from pathlib import Path
from typing import Any, Dict, Optional, Callable
import pandas as pd
from functools import wraps
import streamlit as st
import gc
from datetime import datetime, timedelta

class CacheManager:
    """Comprehensive caching system for the application"""
    
    def __init__(self, cache_dir: str = "cache", max_cache_size: int = 100):
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(exist_ok=True)
        self.max_cache_size = max_cache_size
        self.memory_cache = {}
        self.cache_stats = {
            'hits': 0,
            'misses': 0,
            'size': 0
        }
        
        # Initialize cache subdirectories
        for subdir in ['config', 'excel', 'ai_responses', 'patterns']:
            (self.cache_dir / subdir).mkdir(exist_ok=True)
    
    def _get_cache_key(self, *args, **kwargs) -> str:
        """Generate a unique cache key from arguments"""
        key_data = str(args) + str(sorted(kwargs.items()))
        return hashlib.md5(key_data.encode()).hexdigest()
    
    def _get_file_hash(self, file_path: Path) -> str:
        """Get hash of file contents for cache invalidation"""
        if not file_path.exists():
            return ""
        
        with open(file_path, 'rb') as f:
            return hashlib.md5(f.read()).hexdigest()
    
    def _is_cache_valid(self, cache_file: Path, source_file: Optional[Path], ttl: int = 3600) -> bool:
        """Check if cache is still valid"""
        if not cache_file.exists():
            return False
        
        # Check TTL
        cache_time = cache_file.stat().st_mtime
        if time.time() - cache_time > ttl:
            return False
        
        # Check file hash if source file exists
        if source_file and source_file.exists():
            try:
                with open(cache_file, 'rb') as f:
                    cached_data = pickle.load(f)
                    if 'file_hash' in cached_data:
                        current_hash = self._get_file_hash(source_file)
                        return cached_data['file_hash'] == current_hash
            except:
                return False
        
        return True
    
    def get_cached_config(self, config_file: str) -> Optional[Dict]:
        """Get cached configuration with file hash validation"""
        cache_key = Path(config_file).stem
        cache_file = self.cache_dir / 'config' / f'{cache_key}.pkl'
        source_file = Path(config_file)
        
        if self._is_cache_valid(cache_file, source_file):
            try:
                with open(cache_file, 'rb') as f:
                    cached_data = pickle.load(f)
                    self.cache_stats['hits'] += 1
                    return cached_data['content']
            except:
                pass
        
        self.cache_stats['misses'] += 1
        return None
    
    def cache_config(self, config_file: str, content: Dict):
        """Cache configuration file with hash"""
        cache_key = Path(config_file).stem
        cache_file = self.cache_dir / 'config' / f'{cache_key}.pkl'
        source_file = Path(config_file)
        
        cache_data = {
            'content': content,
            'file_hash': self._get_file_hash(source_file) if source_file.exists() else "",
            'timestamp': time.time()
        }
        
        with open(cache_file, 'wb') as f:
            pickle.dump(cache_data, f)
    
    def get_cached_excel_data(self, file_path: str, sheet_name: str) -> Optional[pd.DataFrame]:
        """Get cached Excel data"""
        cache_key = self._get_cache_key(file_path, sheet_name)
        cache_file = self.cache_dir / 'excel' / f'{cache_key}.pkl'
        source_file = Path(file_path)
        
        if self._is_cache_valid(cache_file, source_file):
            try:
                with open(cache_file, 'rb') as f:
                    cached_data = pickle.load(f)
                    self.cache_stats['hits'] += 1
                    return cached_data['dataframe']
            except:
                pass
        
        self.cache_stats['misses'] += 1
        return None
    
    def cache_excel_data(self, file_path: str, sheet_name: str, dataframe: pd.DataFrame):
        """Cache Excel data with file hash"""
        cache_key = self._get_cache_key(file_path, sheet_name)
        cache_file = self.cache_dir / 'excel' / f'{cache_key}.pkl'
        source_file = Path(file_path)
        
        cache_data = {
            'dataframe': dataframe,
            'file_hash': self._get_file_hash(source_file) if source_file.exists() else "",
            'timestamp': time.time()
        }
        
        with open(cache_file, 'wb') as f:
            pickle.dump(cache_data, f)
    
    def get_cached_ai_response(self, query: str, system_prompt: str, context: str) -> Optional[str]:
        """Get cached AI response"""
        cache_key = self._get_cache_key(query, system_prompt, context)
        cache_file = self.cache_dir / 'ai_responses' / f'{cache_key}.pkl'
        
        if self._is_cache_valid(cache_file, None, ttl=86400):  # 24 hour TTL for AI responses
            try:
                with open(cache_file, 'rb') as f:
                    cached_data = pickle.load(f)
                    self.cache_stats['hits'] += 1
                    return cached_data['response']
            except:
                pass
        
        self.cache_stats['misses'] += 1
        return None
    
    def cache_ai_response(self, query: str, system_prompt: str, context: str, response: str):
        """Cache AI response"""
        cache_key = self._get_cache_key(query, system_prompt, context)
        cache_file = self.cache_dir / 'ai_responses' / f'{cache_key}.pkl'
        
        cache_data = {
            'response': response,
            'timestamp': time.time()
        }
        
        with open(cache_file, 'wb') as f:
            pickle.dump(cache_data, f)
    
    def get_cached_processed_excel(self, filename: str, entity_name: str, entity_suffixes: list) -> Optional[str]:
        """Get cached processed Excel markdown content"""
        cache_key = self._get_cache_key(filename, entity_name, entity_suffixes)
        cache_file = self.cache_dir / 'excel' / f'processed_{cache_key}.pkl'
        source_file = Path(filename)
        
        if self._is_cache_valid(cache_file, source_file):
            try:
                with open(cache_file, 'rb') as f:
                    cached_data = pickle.load(f)
                    self.cache_stats['hits'] += 1
                    return cached_data['markdown_content']
            except:
                pass
        
        self.cache_stats['misses'] += 1
        return None
    
    def cache_processed_excel(self, filename: str, entity_name: str, entity_suffixes: list, markdown_content: str):
        """Cache processed Excel markdown content"""
        cache_key = self._get_cache_key(filename, entity_name, entity_suffixes)
        cache_file = self.cache_dir / 'excel' / f'processed_{cache_key}.pkl'
        source_file = Path(filename)
        
        cache_data = {
            'markdown_content': markdown_content,
            'file_hash': self._get_file_hash(source_file) if source_file.exists() else "",
            'timestamp': time.time()
        }
        
        with open(cache_file, 'wb') as f:
            pickle.dump(cache_data, f)
    
    def get_cached_processed_excel_by_content(self, file_content_hash: str, original_filename: str, entity_name: str, entity_suffixes: list) -> Optional[str]:
        """Get cached processed Excel markdown content using file content hash"""
        cache_key = self._get_cache_key(original_filename, entity_name, entity_suffixes, file_content_hash)
        cache_file = self.cache_dir / 'excel' / f'processed_content_{cache_key}.pkl'
        
        if cache_file.exists():
            try:
                with open(cache_file, 'rb') as f:
                    cached_data = pickle.load(f)
                    # Check if content hash matches
                    if cached_data.get('file_content_hash') == file_content_hash:
                        self.cache_stats['hits'] += 1
                        return cached_data['markdown_content']
            except:
                pass
        
        self.cache_stats['misses'] += 1
        return None
    
    def cache_processed_excel_by_content(self, file_content_hash: str, original_filename: str, entity_name: str, entity_suffixes: list, markdown_content: str):
        """Cache processed Excel markdown content using file content hash"""
        cache_key = self._get_cache_key(original_filename, entity_name, entity_suffixes, file_content_hash)
        cache_file = self.cache_dir / 'excel' / f'processed_content_{cache_key}.pkl'
        
        cache_data = {
            'markdown_content': markdown_content,
            'file_content_hash': file_content_hash,
            'original_filename': original_filename,
            'timestamp': time.time()
        }
        
        with open(cache_file, 'wb') as f:
            pickle.dump(cache_data, f)
    
    def get_file_content_hash(self, file_content: bytes) -> str:
        """Get hash of file content from bytes"""
        return hashlib.md5(file_content).hexdigest()
    
    def clear_cache(self, cache_type: str = 'all'):
        """Clear cache of specified type"""
        if cache_type == 'all':
            import shutil
            if self.cache_dir.exists():
                shutil.rmtree(self.cache_dir)
                self.cache_dir.mkdir(exist_ok=True)
                for subdir in ['config', 'excel', 'ai_responses', 'patterns']:
                    (self.cache_dir / subdir).mkdir(exist_ok=True)
        else:
            cache_subdir = self.cache_dir / cache_type
            if cache_subdir.exists():
                import shutil
                shutil.rmtree(cache_subdir)
                cache_subdir.mkdir(exist_ok=True)
        
        self.memory_cache.clear()
        self.cache_stats = {'hits': 0, 'misses': 0, 'size': 0}
    
    def get_cache_stats(self) -> Dict:
        """Get cache statistics"""
        total_requests = self.cache_stats['hits'] + self.cache_stats['misses']
        hit_rate = (self.cache_stats['hits'] / total_requests * 100) if total_requests > 0 else 0
        
        return {
            'hits': self.cache_stats['hits'],
            'misses': self.cache_stats['misses'],
            'hit_rate': f"{hit_rate:.1f}%",
            'cache_size': len(self.memory_cache)
        }
    
    def cleanup_old_cache(self, days_old: int = 7):
        """Clean up cache files older than specified days"""
        cutoff_time = time.time() - (days_old * 24 * 3600)
        
        for cache_type in ['config', 'excel', 'ai_responses', 'patterns']:
            cache_subdir = self.cache_dir / cache_type
            if cache_subdir.exists():
                for cache_file in cache_subdir.glob('*.pkl'):
                    if cache_file.stat().st_mtime < cutoff_time:
                        cache_file.unlink()

# Global cache manager instance
_cache_manager = None

def get_cache_manager() -> CacheManager:
    """Get global cache manager instance"""
    global _cache_manager
    if _cache_manager is None:
        _cache_manager = CacheManager()
    return _cache_manager

def cached_function(ttl: int = 3600):
    """Decorator for caching function results"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            cache_manager = get_cache_manager()
            cache_key = cache_manager._get_cache_key(func.__name__, *args, **kwargs)
            
            # Check memory cache first
            if cache_key in cache_manager.memory_cache:
                cache_data = cache_manager.memory_cache[cache_key]
                if time.time() - cache_data['timestamp'] < ttl:
                    cache_manager.cache_stats['hits'] += 1
                    return cache_data['result']
            
            # Execute function
            result = func(*args, **kwargs)
            
            # Cache result
            cache_manager.memory_cache[cache_key] = {
                'result': result,
                'timestamp': time.time()
            }
            cache_manager.cache_stats['misses'] += 1
            
            # Limit memory cache size
            if len(cache_manager.memory_cache) > cache_manager.max_cache_size:
                # Remove oldest entries
                sorted_items = sorted(
                    cache_manager.memory_cache.items(),
                    key=lambda x: x[1]['timestamp']
                )
                for key, _ in sorted_items[:len(sorted_items)//2]:
                    del cache_manager.memory_cache[key]
                gc.collect()
            
            return result
        return wrapper
    return decorator

def streamlit_cache_manager():
    """Initialize Streamlit session state cache"""
    if 'cache_manager' not in st.session_state:
        st.session_state.cache_manager = get_cache_manager()
    
    if 'cached_configs' not in st.session_state:
        st.session_state.cached_configs = {}
    
    if 'cached_excel_data' not in st.session_state:
        st.session_state.cached_excel_data = {}
    
    return st.session_state.cache_manager

def optimize_memory():
    """Optimize memory usage"""
    gc.collect()
    
    # Clean up Streamlit session state if needed
    if hasattr(st, 'session_state'):
        # Remove old cached data
        for key in list(st.session_state.keys()):
            if key.startswith('temp_') or key.startswith('cached_'):
                if hasattr(st.session_state[key], '__sizeof__'):
                    # Remove large objects older than 1 hour
                    if hasattr(st.session_state, f'{key}_timestamp'):
                        timestamp = getattr(st.session_state, f'{key}_timestamp')
                        if time.time() - timestamp > 3600:
                            delattr(st.session_state, key)
                            if hasattr(st.session_state, f'{key}_timestamp'):
                                delattr(st.session_state, f'{key}_timestamp') 