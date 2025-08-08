#!/usr/bin/env python3
"""
Simple file-based caching system for Excel processing and AI results
Uses actual filenames instead of complex hashes for better reliability
"""

import json
import os
import hashlib
import time
from pathlib import Path
from typing import Optional, Dict, Any
import pickle

class SimpleCache:
    """Simple file-based caching system"""
    
    def __init__(self, cache_dir: str = "cache"):
        # Ensure cache directory is relative to project root
        project_root = Path(__file__).parent.parent  # Go up from utils/ to project root
        self.cache_dir = project_root / cache_dir
        self.cache_dir.mkdir(exist_ok=True)
        
        # Create subdirectories
        for subdir in ['excel', 'ai_results', 'config']:
            (self.cache_dir / subdir).mkdir(exist_ok=True)
    
    def _get_safe_filename(self, filename: str) -> str:
        """Convert filename to safe cache filename"""
        # Remove path and extension, keep only the base name
        base_name = Path(filename).stem
        # Replace any problematic characters
        safe_name = "".join(c for c in base_name if c.isalnum() or c in ('-', '_'))
        return safe_name
    
    def _get_file_hash(self, file_path: str) -> str:
        """Get hash of file contents"""
        try:
            with open(file_path, 'rb') as f:
                return hashlib.md5(f.read()).hexdigest()
        except:
            return ""
    
    def get_cached_excel_data(self, filename: str, entity_name: str, force_refresh: bool = False) -> Optional[str]:
        """Get cached Excel processing results"""
        if force_refresh:
            print(f"🔄 Force refresh requested for Excel data")
            return None
            
        safe_name = self._get_safe_filename(filename)
        cache_file = self.cache_dir / 'excel' / f'{safe_name}_{entity_name}.json'
        
        if cache_file.exists():
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # Check if source file has changed
                if os.path.exists(filename):
                    current_hash = self._get_file_hash(filename)
                    if data.get('file_hash') == current_hash:
                        # Check if cache is not too old (1 hour for Excel data)
                        if time.time() - data.get('timestamp', 0) < 3600:
                            print(f"📋 Using cached Excel data for {safe_name}")
                            return data.get('content')
                        else:
                            print(f"📋 Excel cache expired for {safe_name}, regenerating")
                    else:
                        print(f"📋 Excel file changed, regenerating cache for {safe_name}")
                else:
                    print(f"📋 Source file not found, using cached data for {safe_name}")
                    return data.get('content')
            except Exception as e:
                print(f"⚠️ Error reading cache: {e}")
        
        return None
    
    def cache_excel_data(self, filename: str, entity_name: str, content: str):
        """Cache Excel processing results"""
        safe_name = self._get_safe_filename(filename)
        cache_file = self.cache_dir / 'excel' / f'{safe_name}_{entity_name}.json'
        
        data = {
            'content': content,
            'file_hash': self._get_file_hash(filename) if os.path.exists(filename) else "",
            'timestamp': time.time(),
            'source_file': filename,
            'entity_name': entity_name
        }
        
        try:
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            print(f"💾 Cached Excel data for {safe_name}")
        except Exception as e:
            print(f"⚠️ Error caching Excel data: {e}")
    
    def get_cached_ai_result(self, key: str, entity_name: str, force_refresh: bool = False) -> Optional[str]:
        """Get cached AI processing results"""
        if force_refresh:
            print(f"🔄 Force refresh requested for AI result")
            return None
            
        safe_key = "".join(c for c in key if c.isalnum() or c in ('-', '_'))
        cache_file = self.cache_dir / 'ai_results' / f'{safe_key}_{entity_name}.json'
        
        if cache_file.exists():
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # Check if cache is still valid (1 hour instead of 24 hours)
                if time.time() - data.get('timestamp', 0) < 3600:
                    # Using cached AI result (silent for better UX)
                    return data.get('content')
                else:
                    # Cache expired, will regenerate (silent for better UX)
                    pass
            except Exception as e:
                print(f"⚠️ Error reading AI cache: {e}")
        
        return None
    
    def cache_ai_result(self, key: str, entity_name: str, content: str):
        """Cache AI processing results"""
        safe_key = "".join(c for c in key if c.isalnum() or c in ('-', '_'))
        cache_file = self.cache_dir / 'ai_results' / f'{safe_key}_{entity_name}.json'
        
        data = {
            'content': content,
            'timestamp': time.time(),
            'key': key,
            'entity_name': entity_name
        }
        
        try:
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            # Cached AI result (silent for better progress bar display)
        except Exception as e:
            print(f"⚠️ Error caching AI result: {e}")
    
    def clear_cache(self, cache_type: str = 'all'):
        """Clear cache of specified type"""
        if cache_type == 'all':
            import shutil
            if self.cache_dir.exists():
                shutil.rmtree(self.cache_dir)
                self.cache_dir.mkdir(exist_ok=True)
                for subdir in ['excel', 'ai_results', 'config']:
                    (self.cache_dir / subdir).mkdir(exist_ok=True)
                print("🗑️ Cleared all cache")
        else:
            cache_subdir = self.cache_dir / cache_type
            if cache_subdir.exists():
                import shutil
                shutil.rmtree(cache_subdir)
                cache_subdir.mkdir(exist_ok=True)
                print(f"🗑️ Cleared {cache_type} cache")
    
    def list_cache_files(self):
        """List all cache files"""
        print("📁 Cache files:")
        for subdir in ['excel', 'ai_results', 'config']:
            subdir_path = self.cache_dir / subdir
            if subdir_path.exists():
                files = list(subdir_path.glob('*'))
                if files:
                    print(f"  {subdir}/:")
                    for file in files:
                        print(f"    {file.name}")

# Global cache instance
_simple_cache = None

def get_simple_cache() -> SimpleCache:
    """Get global cache instance"""
    global _simple_cache
    if _simple_cache is None:
        _simple_cache = SimpleCache()
    return _simple_cache 