"""
Main Application Window
Provides tabbed interface for form selection
"""

import tkinter as tk
from tkinter import ttk

from .acknowledgment_form import AcknowledgmentFormFrame
from .transfer_form import TransferFormFrame


class MainWindow:
    """Main application window with tabbed interface"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Ajman University - Form Filler")
        self.root.geometry("800x700")
        self.root.minsize(750, 600)

        # Set window icon if available
        try:
            self.root.iconbitmap("Forms/Ajman-University-Logo.ico")
        except Exception:
            pass

        self._setup_styles()
        self._create_widgets()

    def _setup_styles(self):
        """Configure ttk styles"""
        style = ttk.Style()

        # Try to use a modern theme
        available_themes = style.theme_names()
        if 'clam' in available_themes:
            style.theme_use('clam')
        elif 'vista' in available_themes:
            style.theme_use('vista')

        # Configure colors
        style.configure("TFrame", background="#f5f5f5")
        style.configure("TLabel", background="#f5f5f5", font=("Helvetica", 10))
        style.configure("TButton", font=("Helvetica", 10))
        style.configure("TNotebook", background="#f5f5f5")
        style.configure("TNotebook.Tab", font=("Helvetica", 11, "bold"), padding=[20, 8])

        # Accent button style
        style.configure("Accent.TButton", font=("Helvetica", 12, "bold"))

    def _create_widgets(self):
        """Create main window widgets"""
        # Header
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", padx=10, pady=10)

        title_label = ttk.Label(
            header_frame,
            text="Ajman University - Main Store Forms",
            font=("Helvetica", 18, "bold")
        )
        title_label.pack()

        subtitle_label = ttk.Label(
            header_frame,
            text="Select a form type and fill in the details to generate a PDF",
            font=("Helvetica", 10)
        )
        subtitle_label.pack()

        # Notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Acknowledgment Form tab
        self.ack_frame = AcknowledgmentFormFrame(self.notebook)
        self.notebook.add(self.ack_frame, text="  Acknowledgment of Receipt  ")

        # Transfer Form tab
        self.transfer_frame = TransferFormFrame(self.notebook)
        self.notebook.add(self.transfer_frame, text="  Asset Transfer Form (ATF)  ")

        # Bottom button bar
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=10)

        clear_btn = ttk.Button(button_frame, text="Clear Form", command=self._clear_current_form)
        clear_btn.pack(side="left", padx=5)

        exit_btn = ttk.Button(button_frame, text="Exit", command=self.root.quit)
        exit_btn.pack(side="right", padx=5)

    def _clear_current_form(self):
        """Clear the currently active form"""
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:
            self.ack_frame.clear_form()
        else:
            self.transfer_frame.clear_form()

    def run(self):
        """Start the application"""
        self.root.mainloop()
