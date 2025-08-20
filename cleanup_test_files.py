#!/usr/bin/env python3
"""
Clean up all testing and debug files from the root directory.
"""

import os
import glob

def cleanup_test_files():
    """Clean up all testing and debug files."""
    
    print("🧹 CLEANING UP TEST AND DEBUG FILES")
    print("=" * 60)
    
    # Patterns for files to remove
    patterns_to_remove = [
        'test_*.py',
        'debug_*.py', 
        'analyze_*.py',
        'check_*.py',
        'clean_*.py',
        'cleanup_*.py',
        'deep_*.py',
        'find_*.py',
        'investigate_*.py',
        'new_*.py',
        'quick_*.py',
        'server_*.py',
        'verify_*.py',
        'final_*.py'
    ]
    
    # Specific files to remove
    specific_files = [
        'FILE_PLACEMENT_GUIDE.md',
        'DEEPSEEK_INTEGRATION_SUMMARY.md',
        'LOCAL_AI_SETUP.md'
    ]
    
    removed_count = 0
    
    print("🗑️  Removing test files...")
    
    # Remove files matching patterns
    for pattern in patterns_to_remove:
        files = glob.glob(pattern)
        for file in files:
            if os.path.exists(file):
                try:
                    os.remove(file)
                    print(f"   ✅ Removed: {file}")
                    removed_count += 1
                except Exception as e:
                    print(f"   ❌ Failed to remove {file}: {e}")
    
    # Remove specific files
    for file in specific_files:
        if os.path.exists(file):
            try:
                os.remove(file)
                print(f"   ✅ Removed: {file}")
                removed_count += 1
            except Exception as e:
                print(f"   ❌ Failed to remove {file}: {e}")
    
    # Remove __pycache__ directories
    pycache_dirs = []
    for root, dirs, files in os.walk('.'):
        for dir_name in dirs:
            if dir_name == '__pycache__':
                pycache_dirs.append(os.path.join(root, dir_name))
    
    print(f"\n🗑️  Removing __pycache__ directories...")
    for pycache_dir in pycache_dirs:
        try:
            import shutil
            shutil.rmtree(pycache_dir)
            print(f"   ✅ Removed: {pycache_dir}")
            removed_count += 1
        except Exception as e:
            print(f"   ❌ Failed to remove {pycache_dir}: {e}")
    
    print(f"\n📊 Cleanup Summary:")
    print(f"   🗑️  Files/directories removed: {removed_count}")
    
    # Show remaining files in root
    remaining_files = []
    for item in os.listdir('.'):
        if os.path.isfile(item) and not item.startswith('.'):
            remaining_files.append(item)
    
    print(f"   📋 Remaining files in root: {len(remaining_files)}")
    
    # Show essential files that should remain
    essential_files = [
        'fdd_app.py',
        'databook.xlsx', 
        'requirements.txt',
        'README.md',
        'LICENSE'
    ]
    
    print(f"\n📋 Essential files check:")
    for file in essential_files:
        if file in remaining_files:
            print(f"   ✅ {file}")
        else:
            print(f"   ❌ {file} (missing)")
    
    print(f"\n✅ Cleanup completed!")

if __name__ == "__main__":
    cleanup_test_files()
