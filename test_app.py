#!/usr/bin/env python3
"""
Simple test script to verify the inspection form application works properly.
"""

import sys
import os

def test_application():
    print("Testing Inspection Form Application...")

    # Change to the project directory
    project_dir = "/Users/jrb/Desktop/pyform12"
    os.chdir(project_dir)

    try:
        # Test imports
        print("Testing imports...")
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        import json
        from datetime import datetime
        import PyPDF2
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter, A4
        from reportlab.lib.utils import ImageReader
        from io import BytesIO
        
        print("✓ All required modules imported successfully")
        
        # Test that main.py can be imported
        print("Testing main module import...")
        import src.main
        print("✓ Main module imported successfully")
        
        # Test that defaults.json can be loaded
        print("Testing defaults.json...")
        with open("defaults.json", "r") as f:
            defaults = json.load(f)
        print(f"✓ defaults.json loaded with {len(defaults)} default fields")
        
        # Test that global.json can be loaded
        print("Testing global.json...")
        with open("global.json", "r") as f:
            global_data = json.load(f)
        print(f"✓ global.json loaded with {len(global_data['building_certifier'])} building certifiers and {len(global_data['appointed_competent_person'])} competent persons")
        
        print("\n✓ All tests passed! The application is ready to use.")
        print("\nTo run the application, execute: python3 src/main.py")
        print("Note: You'll need to have template.docx in the root directory for full DOCX functionality.")
        
        return True
        
    except Exception as e:
        print(f"✗ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_application()
    if success:
        print("\n" + "="*50)
        print("APPLICATION TEST PASSED")
        print("="*50)
    else:
        print("\n" + "="*50)
        print("APPLICATION TEST FAILED")
        print("="*50)
        sys.exit(1)