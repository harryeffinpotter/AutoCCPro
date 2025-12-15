"""Wrapper to catch and display any startup errors"""
import sys
import traceback

try:
    import app_gui
except Exception as e:
    print("=" * 60)
    print("FATAL ERROR ON STARTUP:")
    print("=" * 60)
    traceback.print_exc()
    print("=" * 60)
    input("Press Enter to exit...")
    sys.exit(1)
