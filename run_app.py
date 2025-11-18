#!/usr/bin/env python3
"""
Start script for the Inspection Form Application
"""

import sys
import os

def main():
    """Main entry point for the application."""
    # Add the src directory to Python path
    current_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.join(current_dir, 'src')
    sys.path.insert(0, src_dir)

    try:
        from main import main as app_main
        print("Starting Inspection Form Application...")
        print("The application now generates DOCX files from your template.")
        print("Make sure template.docx is in the root directory.")
        print("Application window size optimized for 14\" MacBook Air: 1100x800.")
        print("If you see a GUI window, the application has started successfully.")
        print("Close the window to exit the application.")
        app_main()
    except ImportError as e:
        print(f"Error importing main application: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error running application: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()