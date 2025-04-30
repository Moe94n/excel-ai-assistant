#!/usr/bin/env python3
"""
Excel AI Assistant - Main Application Entry Point

This application allows users to apply AI transformations to Excel and CSV data
using OpenAI's API.
"""

import os
import sys
import tkinter as tk

# Add the current directory to the path so we can import our modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Local application imports
from app.config import AppConfig
from app.ui.main_window import ExcelAIAssistantApp
from app.utils.theme_manager import ThemeManager


def main():
    """Main application entry point"""
    # Create the root window
    root = tk.Tk()
    root.title("Excel AI Assistant")
    root.geometry("1200x800")
    root.minsize(900, 600)

    # Set application icon if available
    try:
        if sys.platform == "darwin":  # macOS
            # macOS handles app icons differently
            pass
        elif sys.platform == "win32":  # Windows
            root.iconbitmap(os.path.join(os.path.dirname(__file__), "resources", "app_icon.ico"))
        else:  # Linux/Unix
            img = tk.PhotoImage(file=os.path.join(os.path.dirname(__file__), "resources", "app_icon.png"))
            root.tk.call('wm', 'iconphoto', root._w, img)
    except Exception:
        # If icon loading fails, continue without an icon
        pass

    # Load configuration
    config = AppConfig()

    # Setup theme manager
    theme_manager = ThemeManager(root)

    # Default to system theme if possible, otherwise use light theme
    system_theme = theme_manager.detect_system_theme()
    theme_manager.set_theme(config.get('theme', system_theme or 'light'))

    # Initialize the main application
    app = ExcelAIAssistantApp(root, config, theme_manager)

    # Start the application main loop
    root.mainloop()

    # Save configuration on exit
    config.save()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Display error message in a dialog if possible
        try:
            import traceback
            from tkinter import messagebox

            error_message = f"An unexpected error occurred: {str(e)}\n\n{traceback.format_exc()}"
            messagebox.showerror("Unexpected Error", error_message)
        except Exception as e:
            # If that fails, print to stderr
            import traceback

            print(f"An unexpected error occurred: {str(e)}", file=sys.stderr)
            traceback.print_exc()
        sys.exit(1)
