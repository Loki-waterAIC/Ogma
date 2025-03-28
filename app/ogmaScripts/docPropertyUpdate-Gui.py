import threading
import tkinter as tk
from concurrent.futures import Future, ThreadPoolExecutor
from tkinter import filedialog, messagebox, ttk
import os, sys
import copy

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

from app.ogmaScripts.documentPropertyUpdateTool import document_properity_update_tool as run_scripts

FILE_TYPES: list[tuple[str, str]] = [("Docx files", "*.docx;"), ("All files", "*;")]
TITLE_NAME = "OGMA Mass Doc Property Update Tool GUI"


# Wrapper function to run scripts and show a GUI message
def run_scripts_gui(file_paths: list[str], properties: dict[str, str], print: bool) -> None:
    confirmation = messagebox.askyesno(
        title="Confirm Files",
        message=f"Are you sure you want to run scripts for the following files?\n\n{', '.join(file_paths)}",
    )
    if confirmation:
        run_scripts(doc_paths=file_paths, properties=properties, export_pdf=print)
        messagebox.showinfo("Finished", f"Finished running scripts for files: {file_paths}")
    else:
        messagebox.showinfo("Cancelled", "Script execution was cancelled.")


# Main application class
class GUIApp:
    def __init__(self, root:tk.Tk):
        self.root: tk.Tk = root
        self.root.title(string=TITLE_NAME)
        self.root.geometry(newGeometry="600x800")
        self.root.minsize(width=400,height=720)

        self.file_paths: list[str] = []
        self.checkboxes: list[tuple[tk.BooleanVar, tk.Checkbutton, tk.Label]] = []
        self.properties: dict[str, tk.Entry] = {}
        self.print: bool = True

        # MARK: Top buttons
        self.toggle_all_button = tk.Button(root, text="Toggle all", command=self.toggle_all)
        self.toggle_all_button.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.select_button = tk.Button(root, text="Select parent folder", command=self.select_parent_folder)
        self.select_button.grid(row=0, column=1, padx=10, pady=10, sticky="e")

        self.select_button = tk.Button(root, text="Select docx files", command=self.select_files)
        self.select_button.grid(row=0, column=2, padx=10, pady=10, sticky="e")

        # MARK: Text Box - File Frame
        self.tb_frame = tk.Frame(root)
        self.tb_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        self.text_box_canvas = tk.Canvas(self.tb_frame, bg="white")
        self.h_scrollbar = ttk.Scrollbar(self.tb_frame, orient="horizontal", command=self.text_box_canvas.xview)
        self.v_scrollbar = ttk.Scrollbar(self.tb_frame, orient="vertical", command=self.text_box_canvas.yview)
        self.scrollable_frame = tk.Frame(self.text_box_canvas, bg="white")
        self.text_box_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.text_box_canvas.configure(xscrollcommand=self.h_scrollbar.set, yscrollcommand=self.v_scrollbar.set)
        self.h_scrollbar.pack(side="bottom", fill="x")
        self.v_scrollbar.pack(side="right", fill="y")
        self.text_box_canvas.pack(side="left", fill="both", expand=True)

        # MARK: Bottom buttons
        self.remove_button = tk.Button(root, text="Remove", command=self.remove_files)
        self.remove_button.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.print_checkbox_var = tk.BooleanVar(value=self.print)
        self.print_checkbox = tk.Checkbutton(
            root, text="Export to PDF", variable=self.print_checkbox_var, command=self.toggle_print
        )
        self.print_checkbox.grid(row=2, column=1, padx=10, pady=10, sticky="e")

        self.run_button = tk.Button(root, text="Run all", command=self.run_all)
        self.run_button.grid(row=2, column=2, padx=10, pady=10, sticky="e")

        # MARK: User input
        # User Input Fields
        input_titles: list[str] = [
            "BOK ID",
            "Document Name",
            "Company Name",
            "Division",
            "Author",
            "Company Address",
            "Project Name",
            "Project Number",
            "End Customer",
            "Site Name",
            "File Name",
        ]
        # User Input - bottom text section
        self.ui_frame = tk.Frame(root)
        self.ui_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        self.ui_canvas = tk.Canvas(self.ui_frame, bg="white")
        self.ui_h_scrollbar = ttk.Scrollbar(self.ui_frame, orient="horizontal", command=self.ui_canvas.xview)
        self.ui_v_scrollbar = ttk.Scrollbar(self.ui_frame, orient="vertical", command=self.ui_canvas.yview)
        self.ui_scroll_frame = tk.Frame(self.ui_canvas, bg="white")
        self.ui_canvas.create_window((0, 0), window=self.ui_scroll_frame, anchor="nw")
        self.ui_canvas.configure(xscrollcommand=self.ui_h_scrollbar.set, yscrollcommand=self.ui_v_scrollbar.set)
        self.ui_h_scrollbar.pack(side="bottom", fill="x")
        self.ui_v_scrollbar.pack(side="right", fill="y")
        self.ui_canvas.pack(side="left", fill="both", expand=True)
        # fill out canvas
        for i, title in enumerate(input_titles):
            self.user_input = tk.Frame(self.ui_scroll_frame)
            self.user_input.grid(row=i, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

            tk.Label(self.user_input, text=title).grid(row=0, column=0, padx=10, pady=5, sticky="w")
            entry = tk.Entry(master=self.user_input)
            entry.grid(row=0, column=1, padx=10, pady=5, sticky="w", columnspan=2)
            self.properties.update({title: entry})

        root.grid_rowconfigure(1, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)

    # MARK: Logic
    def toggle_print(self):
        self.print = self.print_checkbox_var.get()

    def select_files(self):
        filetypes = FILE_TYPES
        selected_files = filedialog.askopenfilenames(title="Select files", filetypes=filetypes)
        for file_path in selected_files:
            if file_path not in self.file_paths:
                self.file_paths.append(file_path)
                self.add_file_to_text_box(file_path)

    def select_parent_folder(self):
        parent_folder = filedialog.askdirectory(title="Select parent folder", mustexist=True)
        for root, _, files in os.walk(parent_folder):
            for file in files:
                if file.endswith(".docx"):
                    file_path = os.path.join(root, file)
                    if file_path not in self.file_paths:
                        self.file_paths.append(file_path)
                        self.add_file_to_text_box(file_path)

    def add_file_to_text_box(self, file_path):
        var = tk.BooleanVar(value=True)
        checkbox = tk.Checkbutton(self.scrollable_frame, variable=var, bg="white")
        label = tk.Label(self.scrollable_frame, text=file_path, anchor="w", bg="white")
        self.checkboxes.append((var, checkbox, label))
        checkbox.grid(row=len(self.checkboxes) - 1, column=0, sticky="w")
        label.grid(row=len(self.checkboxes) - 1, column=1, sticky="w")

    def toggle_all(self):
        if not self.checkboxes:
            return
        new_state = not self.checkboxes[0][0].get()
        for var, _, _ in self.checkboxes:
            var.set(new_state)

    def remove_files(self):
        self.file_paths = [fp for i, fp in enumerate(self.file_paths) if not self.checkboxes[i][0].get()]
        self.checkboxes = [cb for cb in self.checkboxes if not cb[0].get()]

    def run_all(self):
        selected_files = [self.file_paths[i] for i, (var, _, _) in enumerate(self.checkboxes) if var.get()]
        if selected_files:
            run_scripts_gui(
                file_paths=selected_files, properties={k: v.get() for k, v in self.properties.items()}, print=self.print
            )
        else:
            messagebox.showwarning("No Files Selected", "Please select at least one file to run.")


if __name__ == "__main__":
    root = tk.Tk()
    app = GUIApp(root)
    root.mainloop()
