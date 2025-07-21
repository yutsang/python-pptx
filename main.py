#!/usr/bin/env python3
"""
Main entry point for the Due Diligence Automation System.

This is the improved enterprise-grade version with hexagonal architecture.
The original implementation has been preserved in the old_ver/ directory.
"""

import os
import sys
from pathlib import Path

# Add src to Python path for imports
current_dir = Path(__file__).parent
src_dir = current_dir / "src"
sys.path.insert(0, str(src_dir))

def run_streamlit():
    """Run the Streamlit application."""
    print("🚀 Starting Due Diligence Automation System (Improved Architecture)")
    print("📊 Launching Streamlit interface...")
    print("=" * 60)
    
    # Set environment variables for the application
    os.environ.setdefault("STREAMLIT_SERVER_PORT", "8501")
    os.environ.setdefault("STREAMLIT_SERVER_ADDRESS", "localhost")
    
    # Check available applications in priority order
    full_app = current_dir / "streamlit_app_full.py"
    working_app = current_dir / "streamlit_app_working.py"
    demo_app = current_dir / "streamlit_app.py"
    new_streamlit_app = src_dir / "interfaces" / "web" / "streamlit_app.py"
    original_app = current_dir / "old_ver" / "app.py"
    
    if full_app.exists():
        print("🚀 Running full-featured application with AI and PowerPoint export...")
        os.system(f"streamlit run {full_app}")
    elif working_app.exists():
        print("✅ Running self-contained new architecture app...")
        os.system(f"streamlit run {working_app}")
    elif demo_app.exists():
        print("🎯 Running architecture demo app...")
        os.system(f"streamlit run {demo_app}")
    elif new_streamlit_app.exists():
        print("🆕 Running new architecture app...")
        os.system(f"streamlit run {new_streamlit_app}")
    elif original_app.exists():
        print("🔄 Falling back to original implementation...")
        os.system(f"streamlit run {original_app}")
    else:
        print("❌ No Streamlit application found!")
        print("💡 Available options:")
        print("   - streamlit run streamlit_app_full.py    (complete functionality)")
        print("   - streamlit run streamlit_app.py         (simplified version)")
        print("   - streamlit run old_ver/app.py           (original working version)")


def run_fastapi():
    """Run the FastAPI application."""
    print("🚀 Starting Due Diligence Automation API (FastAPI)")
    print("🌐 API will be available at http://localhost:8000")
    print("📖 API docs will be available at http://localhost:8000/docs")
    print("=" * 60)
    
    # This would run the FastAPI application when implemented
    fastapi_app = src_dir / "interfaces" / "web" / "fastapi_app.py"
    
    if fastapi_app.exists():
        os.system(f"uvicorn src.interfaces.web.fastapi_app:app --host 0.0.0.0 --port 8000 --reload")
    else:
        print("❌ FastAPI application not yet implemented!")
        print("💡 This is part of the new architecture that would be built in Phase 2")


def run_cli():
    """Run the CLI application."""
    print("🚀 Starting Due Diligence Automation CLI")
    print("=" * 60)
    
    # This would run the CLI application when implemented
    cli_app = src_dir / "interfaces" / "cli" / "main.py"
    
    if cli_app.exists():
        os.system(f"python {cli_app}")
    else:
        print("❌ CLI application not yet implemented!")
        print("💡 This is part of the new architecture")


def show_help():
    """Show help information."""
    print("""
🏗️  Due Diligence Automation System - Improved Architecture

USAGE:
    python main.py [COMMAND]

COMMANDS:
    streamlit, st     Launch Streamlit web interface (default)
    fastapi, api      Launch FastAPI REST API server
    cli               Launch command-line interface
    help, -h, --help  Show this help message

EXAMPLES:
    python main.py                 # Run Streamlit (default)
    python main.py streamlit       # Run Streamlit explicitly
    python main.py api             # Run FastAPI server
    python main.py cli             # Run CLI interface

ARCHITECTURE:
    This improved version uses hexagonal architecture with:
    - 📊 Domain layer: Core business logic
    - 🔧 Application layer: Use cases and DTOs
    - 🏗️ Infrastructure layer: External adapters
    - 🌐 Interface layer: UI and API adapters

MIGRATION STATUS:
    ✅ Domain entities implemented
    ✅ Repository interfaces defined
    ✅ Improved Streamlit UI structure
    ✅ PowerPoint export preserved
    ⏳ Infrastructure implementations (Phase 2)
    ⏳ FastAPI endpoints (Phase 2)
    ⏳ Database integration (Phase 3)

ORIGINAL VERSION:
    The original working implementation is preserved in old_ver/
    Run with: streamlit run old_ver/app.py
    """)


def main():
    """Main entry point."""
    # Get command line argument
    command = sys.argv[1] if len(sys.argv) > 1 else "streamlit"
    
    if command in ["help", "-h", "--help"]:
        show_help()
    elif command in ["streamlit", "st"]:
        run_streamlit()
    elif command in ["fastapi", "api"]:
        run_fastapi() 
    elif command in ["cli"]:
        run_cli()
    else:
        print(f"❌ Unknown command: {command}")
        print("💡 Run 'python main.py help' for usage information")
        sys.exit(1)


if __name__ == "__main__":
    main() 