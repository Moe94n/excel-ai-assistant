"""
Cloud storage file browser dialog for Excel AI Assistant.
Allows users to browse and select files from cloud storage services.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, List
import threading


class CloudStorageDialog:
    """Dialog for browsing cloud storage files"""

    def __init__(self, parent: tk.Tk, service_name: str, service):
        """Initialize the cloud storage dialog"""
        self.parent = parent
        self.service_name = service_name
        self.service = service
        self.selected_file = None
        self.current_path = []
        self.files = []

        # Create dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Browse {service_name}")
        self.dialog.geometry("800x600")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._create_ui()
        self._load_files()

        # Center dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - self.dialog.winfo_width()) // 2
        y = (self.dialog.winfo_screenheight() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def _create_ui(self):
        """Create the dialog UI"""
        # Main frame
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

    def _safe_ui_update(self, callback, *args, **kwargs):
        """
        Safely update UI from a background thread.

        Args:
            callback: Function to call in the main thread
            *args: Arguments to pass to the callback
            **kwargs: Keyword arguments to pass to the callback
        """
        if not hasattr(self, 'dialog') or not self.dialog.winfo_exists():
            import logging
            logger = logging.getLogger(__name__)
            logger.warning("Cannot update UI: dialog does not exist")
            return

        try:
            self.dialog.after(0, lambda: self._execute_ui_callback(callback, *args, **kwargs))
        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.error(f"Failed to schedule UI update: {str(e)}")

    def _execute_ui_callback(self, callback, *args, **kwargs):
        """
        Execute UI callback with error handling.

        Args:
            callback: Function to execute
            *args: Arguments to pass to the callback
            **kwargs: Keyword arguments to pass to the callback
        """
        try:
            callback(*args, **kwargs)
        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.error(f"Error executing UI callback {callback.__name__}: {str(e)}")
            import traceback
            logger.error(f"UI callback traceback: {traceback.format_exc()}")

        # Toolbar
        toolbar_frame = ttk.Frame(main_frame)
        toolbar_frame.pack(fill=tk.X, pady=(0, 10))

        # Back button
        self.back_button = ttk.Button(toolbar_frame, text="← Back", command=self._go_back, state=tk.DISABLED)
        self.back_button.pack(side=tk.LEFT, padx=(0, 10))

        # Current path label
        self.path_label = ttk.Label(toolbar_frame, text="Root")
        self.path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Refresh button
        ttk.Button(toolbar_frame, text="Refresh", command=self._load_files).pack(side=tk.RIGHT)

        # File list
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Create treeview for files
        columns = ("name", "type", "size", "modified")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=20)

        # Configure columns
        self.tree.heading("name", text="Name")
        self.tree.heading("type", text="Type")
        self.tree.heading("size", text="Size")
        self.tree.heading("modified", text="Modified")

        self.tree.column("name", width=300)
        self.tree.column("type", width=100)
        self.tree.column("size", width=100)
        self.tree.column("modified", width=150)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind double-click
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Return>", self._on_select)

        # Loading label
        self.loading_label = ttk.Label(list_frame, text="Loading...")
        self.loading_label.pack(pady=20)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="Select", command=self._on_select).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT)

        # Status bar
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.pack(fill=tk.X, pady=(10, 0))

    def _load_files(self):
        """Load files from current directory"""
        import logging
        logger = logging.getLogger(__name__)
        logger.debug("Starting cloud storage file loading")

        self.loading_label.pack(pady=20)
        self.tree.pack_forget()
        self.status_label.config(text="Loading files...")

        def load_thread():
            thread_id = threading.get_ident()
            logger.debug(f"Cloud storage load thread started: {thread_id}")
            try:
                if self.current_path:
                    folder_id = self.current_path[-1]['id']
                    logger.debug(f"Loading files from folder: {folder_id}")
                    files = self.service.list_files(folder_id)
                else:
                    logger.debug("Loading files from root directory")
                    files = self.service.list_files()  # Root directory

                logger.debug(f"Cloud storage load completed, {len(files)} files found")
                self._safe_ui_update(self._update_file_list, files)
                logger.debug(f"Cloud storage UI update scheduled from thread {thread_id}")
            except Exception as e:
                logger.error(f"Exception in cloud storage load thread {thread_id}: {str(e)}")
                import traceback
                logger.error(f"Traceback: {traceback.format_exc()}")
                self._safe_ui_update(self._show_error, str(e))
                logger.debug(f"Cloud storage error UI update scheduled from thread {thread_id}")

        load_thread_obj = self._create_tracked_thread(load_thread, name="CloudStorageLoader", daemon=True)
        load_thread_obj.start()

    def _update_file_list(self, files: List[Dict[str, Any]]):
        """Update the file list in the UI"""
        import logging
        logger = logging.getLogger(__name__)
        thread_id = threading.get_ident()
        logger.debug(f"Cloud storage file list UI update called from thread {thread_id}")

        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.files = files
        logger.debug(f"Updating UI with {len(files)} cloud storage files")

        # Add files and folders
        for file_info in files:
            name = file_info.get('name', 'Unknown')
            file_type = "Folder" if file_info.get('mimeType') == 'application/vnd.google-apps.folder' else "File"
            size = self._format_size(file_info.get('size', 0))
            modified = self._format_date(file_info.get('modifiedTime', ''))

            # Insert item
            item_id = self.tree.insert("", tk.END, values=(name, file_type, size, modified))

            # Store file info
            self.tree.item(item_id, tags=(str(len(self.files) - 1),))

        # Hide loading label and show tree
        self.loading_label.pack_forget()
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Update path label
        if self.current_path:
            path_names = [item['name'] for item in self.current_path]
            self.path_label.config(text=" / ".join(path_names))
            self.back_button.config(state=tk.NORMAL)
        else:
            self.path_label.config(text="Root")
            self.back_button.config(state=tk.DISABLED)

        self.status_label.config(text=f"Found {len(files)} items")
        logger.debug("Cloud storage file list UI update completed")

    def _show_error(self, error: str):
        """Show error message"""
        self.loading_label.pack_forget()
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.status_label.config(text=f"Error: {error}")
        messagebox.showerror("Error", f"Failed to load files: {error}")

    def _on_double_click(self, event):
        """Handle double-click on item"""
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        item_tags = self.tree.item(item, 'tags')

        if item_tags:
            index = int(item_tags[0])
            file_info = self.files[index]

            if file_info.get('mimeType') == 'application/vnd.google-apps.folder':
                # Navigate into folder
                self.current_path.append(file_info)
                self._load_files()
            else:
                # Select file
                self.selected_file = file_info
                self._on_select()

    def _on_select(self, event=None):
        """Handle file selection"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Please select a file first.")
            return

        item = selection[0]
        item_tags = self.tree.item(item, 'tags')

        if item_tags:
            index = int(item_tags[0])
            file_info = self.files[index]

            if file_info.get('mimeType') == 'application/vnd.google-apps.folder':
                messagebox.showinfo("Folder Selected", "Please select a file, not a folder.")
                return

            self.selected_file = file_info
            self.dialog.destroy()

    def _go_back(self):
        """Go back to parent directory"""
        if self.current_path:
            self.current_path.pop()
            self._load_files()

    def _on_cancel(self):
        """Cancel dialog"""
        self.selected_file = None
        self.dialog.destroy()

    def get_selected_file(self) -> Optional[Dict[str, Any]]:
        """Get the selected file info"""
        self.parent.wait_window(self.dialog)
        return self.selected_file

    def _format_size(self, size_bytes: int) -> str:
        """Format file size"""
        if not size_bytes:
            return ""

        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return ".1f"
            size_bytes /= 1024.0
        return ".1f"

    def _format_date(self, date_str: str) -> str:
        """Format modification date"""
        if not date_str:
            return ""

        try:
            from datetime import datetime
            # Parse ISO format date
            dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            return dt.strftime("%Y-%m-%d %H:%M")
        except:
            return date_str