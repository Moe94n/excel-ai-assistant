"""
Cloud storage settings dialog for Excel AI Assistant.
Allows users to configure cloud storage integrations.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, Any, Optional

from app.config import AppConfig


class CloudStorageSettingsDialog:
    """Dialog for configuring cloud storage settings"""

    def __init__(self, parent: tk.Tk, config: AppConfig):
        """Initialize the cloud storage settings dialog"""
        self.parent = parent
        self.config = config
        self.result = None

        # Create dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Cloud Storage Settings")
        self.dialog.geometry("600x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._create_ui()

        # Center dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - self.dialog.winfo_width()) // 2
        y = (self.dialog.winfo_screenheight() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def _create_ui(self):
        """Create the dialog UI"""
        # Main frame
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text="Cloud Storage Configuration",
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))

        # Notebook for different services
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Google Drive tab
        google_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(google_frame, text="Google Drive")

        self._create_google_drive_tab(google_frame)

        # OneDrive tab (placeholder)
        onedrive_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(onedrive_frame, text="OneDrive")

        ttk.Label(onedrive_frame, text="OneDrive integration coming soon!",
                 font=("Arial", 12)).pack(pady=20)

        # Dropbox tab (placeholder)
        dropbox_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(dropbox_frame, text="Dropbox")

        ttk.Label(dropbox_frame, text="Dropbox integration coming soon!",
                 font=("Arial", 12)).pack(pady=20)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="Save", command=self._save_settings).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT)

    def _create_google_drive_tab(self, parent):
        """Create the Google Drive settings tab"""
        # Instructions
        instructions_frame = ttk.LabelFrame(parent, text="Setup Instructions", padding=10)
        instructions_frame.pack(fill=tk.X, pady=(0, 15))

        instructions_text = (
            "1. Go to https://console.cloud.google.com/\n"
            "2. Create a new project or select existing one\n"
            "3. Enable Google Drive API\n"
            "4. Create OAuth 2.0 Desktop Application credentials\n"
            "5. Download the client_secrets.json file\n"
            "6. Place it in a secure location on your computer\n"
            "7. Set the path below"
        )

        instructions_label = ttk.Label(instructions_frame, text=instructions_text, justify=tk.LEFT)
        instructions_label.pack(anchor=tk.W)

        # Settings frame
        settings_frame = ttk.LabelFrame(parent, text="Configuration", padding=10)
        settings_frame.pack(fill=tk.X, pady=(0, 10))

        # Client secrets path
        path_frame = ttk.Frame(settings_frame)
        path_frame.pack(fill=tk.X, pady=5)

        ttk.Label(path_frame, text="Client Secrets File:").pack(side=tk.LEFT, padx=(0, 10))

        self.client_secrets_var = tk.StringVar(value=self.config.get('google_client_secrets_path', ''))
        self.client_secrets_entry = ttk.Entry(path_frame, textvariable=self.client_secrets_var, width=40)
        self.client_secrets_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(path_frame, text="Browse...", command=self._browse_client_secrets).pack(side=tk.RIGHT)

        # Status indicator
        status_frame = ttk.Frame(settings_frame)
        status_frame.pack(fill=tk.X, pady=5)

        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT, padx=(0, 10))

        self.status_var = tk.StringVar(value="Not configured")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var)
        self.status_label.pack(side=tk.LEFT)

        # Test button
        ttk.Button(settings_frame, text="Test Connection", command=self._test_google_drive).pack(pady=10)

        # Update status on startup
        self._update_google_drive_status()

    def _browse_client_secrets(self):
        """Browse for client secrets file"""
        file_path = filedialog.askopenfilename(
            parent=self.dialog,
            title="Select Google Client Secrets File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if file_path:
            self.client_secrets_var.set(file_path)
            self._update_google_drive_status()

    def _update_google_drive_status(self):
        """Update Google Drive configuration status"""
        client_secrets_path = self.client_secrets_var.get()

        if not client_secrets_path:
            self.status_var.set("Not configured")
            self.status_label.config(foreground="red")
            return

        if not os.path.exists(client_secrets_path):
            self.status_var.set("File not found")
            self.status_label.config(foreground="red")
            return

        # Try to validate the JSON file
        try:
            import json
            with open(client_secrets_path, 'r') as f:
                data = json.load(f)

            if 'installed' in data or 'web' in data:
                self.status_var.set("Valid configuration file")
                self.status_label.config(foreground="green")
            else:
                self.status_var.set("Invalid file format")
                self.status_label.config(foreground="red")
        except Exception:
            self.status_var.set("Invalid JSON file")
            self.status_label.config(foreground="red")

    def _test_google_drive(self):
        """Test Google Drive connection"""
        from app.services.cloud import GoogleDriveService
        from app.utils.logger import setup_logger

        client_secrets_path = self.client_secrets_var.get()

        if not client_secrets_path:
            messagebox.showerror("Error", "Please select a client secrets file first.")
            return

        # Create a temporary service to test
        logger = setup_logger("TestGoogleDrive")
        test_service = GoogleDriveService(self.config, logger)

        # Show progress
        self.status_var.set("Testing connection...")
        self.status_label.config(foreground="blue")
        self.dialog.update()

        try:
            success = test_service.authenticate(client_secrets_path)

            if success:
                self.status_var.set("Connection successful!")
                self.status_label.config(foreground="green")
                messagebox.showinfo("Success", "Google Drive connection test successful!")
            else:
                self.status_var.set("Connection failed")
                self.status_label.config(foreground="red")
                messagebox.showerror("Error", "Google Drive connection test failed. Check your configuration.")

        except Exception as e:
            self.status_var.set("Test error")
            self.status_label.config(foreground="red")
            messagebox.showerror("Error", f"Test failed: {str(e)}")

    def _save_settings(self):
        """Save the cloud storage settings"""
        # Save Google Drive settings
        client_secrets_path = self.client_secrets_var.get()
        self.config.set('google_client_secrets_path', client_secrets_path)

        # Save configuration
        self.config.save()

        messagebox.showinfo("Settings Saved", "Cloud storage settings have been saved.")
        self.result = True
        self.dialog.destroy()

    def _on_cancel(self):
        """Cancel the dialog"""
        self.result = False
        self.dialog.destroy()

    def get_result(self) -> bool:
        """Get the dialog result"""
        self.parent.wait_window(self.dialog)
        return self.result or False