"""Simple GUI for Document Comparison."""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import subprocess
import tempfile


class DocumentCompareGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Compare")
        self.root.geometry("600x200")
        self.root.resizable(True, False)

        # File paths
        self.original_path = tk.StringVar()
        self.modified_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky="nsew")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Original file (base)
        ttk.Label(main_frame, text="Original (base):").grid(row=0, column=0, sticky="w", pady=(0, 10))
        original_entry = ttk.Entry(main_frame, textvariable=self.original_path, width=50)
        original_entry.grid(row=0, column=1, sticky="ew", padx=(10, 10), pady=(0, 10))
        ttk.Button(main_frame, text="Browse...", command=self.browse_original).grid(row=0, column=2, pady=(0, 10))

        # Modified file (changes)
        ttk.Label(main_frame, text="Modified (changes):").grid(row=1, column=0, sticky="w", pady=(0, 10))
        modified_entry = ttk.Entry(main_frame, textvariable=self.modified_path, width=50)
        modified_entry.grid(row=1, column=1, sticky="ew", padx=(10, 10), pady=(0, 10))
        ttk.Button(main_frame, text="Browse...", command=self.browse_modified).grid(row=1, column=2, pady=(0, 10))

        # Compare button
        self.compare_btn = ttk.Button(main_frame, text="Compare Documents", command=self.compare)
        self.compare_btn.grid(row=2, column=0, columnspan=3, pady=(20, 0))

        # Status label
        self.status_var = tk.StringVar(value="Select two documents to compare")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var, foreground="gray")
        self.status_label.grid(row=3, column=0, columnspan=3, pady=(10, 0))

    def browse_original(self):
        path = filedialog.askopenfilename(
            title="Select Original (Base) Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.original_path.set(path)

    def browse_modified(self):
        path = filedialog.askopenfilename(
            title="Select Modified (Changed) Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.modified_path.set(path)

    def compare(self):
        original = self.original_path.get()
        modified = self.modified_path.get()

        if not original or not modified:
            messagebox.showerror("Error", "Please select both documents")
            return

        if not os.path.exists(original):
            messagebox.showerror("Error", f"Original file not found:\n{original}")
            return

        if not os.path.exists(modified):
            messagebox.showerror("Error", f"Modified file not found:\n{modified}")
            return

        # Generate output path
        modified_dir = os.path.dirname(modified)
        modified_name = os.path.splitext(os.path.basename(modified))[0]
        output_path = os.path.join(modified_dir, f"{modified_name}_REDLINE.docx")

        # Update UI
        self.compare_btn.config(state="disabled")
        self.status_var.set("Comparing documents...")
        self.status_label.config(foreground="blue")
        self.root.update()

        # Run comparison
        script_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(script_dir, "compare_preserve_formatting.py")

        try:
            result = subprocess.run(
                [sys.executable, script_path, original, modified, output_path],
                capture_output=True,
                text=True,
                cwd=script_dir
            )

            if result.returncode != 0:
                messagebox.showerror("Error", f"Comparison failed:\n{result.stderr}")
                self.compare_btn.config(state="normal")
                self.status_var.set("Comparison failed")
                self.status_label.config(foreground="red")
                return

            # Open the output file
            os.startfile(output_path)

            # Close the app
            self.root.quit()

        except Exception as e:
            messagebox.showerror("Error", f"Error running comparison:\n{str(e)}")
            self.compare_btn.config(state="normal")
            self.status_var.set("Error occurred")
            self.status_label.config(foreground="red")


def main():
    root = tk.Tk()
    app = DocumentCompareGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
