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
        # message=f"Are you sure you want to run scripts for the following files?\n\n{', '.join(file_paths)}",
        message=f"Are you sure you want to run scripts for the selected following files?",
    )
    if confirmation:
        run_scripts(doc_paths=file_paths, properties=properties, export_pdf=print)
        messagebox.showinfo("Finished", f"Finished running scripts on the selected files.")
    else:
        messagebox.showinfo("Cancelled", "Script execution was cancelled.")


# Main application class
class GUIApp:
    def __init__(self, root: tk.Tk):
        self.root: tk.Tk = root
        self.root.title(string=TITLE_NAME)
        self.root.geometry(newGeometry="600x800")
        self.root.minsize(width=400, height=720)

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
        # Text box with scrollbars
        self.text_box_frame = tk.Frame(root)
        self.text_box_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        # Canvas and scrollbars
        self.canvas = tk.Canvas(self.text_box_frame, bg="white")
        self.h_scrollbar = ttk.Scrollbar(self.text_box_frame, orient="horizontal", command=self.canvas.xview)
        self.v_scrollbar = ttk.Scrollbar(self.text_box_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")  # Set background to white
        # Scroll Bars
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        # Make canvas
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set, yscrollcommand=self.v_scrollbar.set)
        # Pack scrollbars and canvas
        self.h_scrollbar.pack(side="bottom", fill="x")
        self.v_scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        # Bind mouse wheel events to the canvas for both horizontal and vertical scrolling
        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)  # Vertical scrolling (Windows/macOS)
        self.canvas.bind_all("<Shift-MouseWheel>", self.on_horizontal_mouse_wheel)  # Horizontal scrolling (Windows/macOS)
        self.canvas.bind_all("<Button-4>", self.on_mouse_wheel)  # Vertical scrolling (Linux, up)
        self.canvas.bind_all("<Button-5>", self.on_mouse_wheel)  # Vertical scrolling (Linux, down)
        self.canvas.bind_all("<Shift-Button-4>", self.on_horizontal_mouse_wheel)  # Horizontal scrolling (Linux, left)
        self.canvas.bind_all("<Shift-Button-5>", self.on_horizontal_mouse_wheel)  # Horizontal scrolling (Linux, right)

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

        # MARK: bottom padding
        row_max = 3
        tk.Label(self.root, text="").grid(row=row_max, column=0, padx=10, pady=0, sticky="w")

        # MARK: User input
        # User Input Fields
        input_titles: list[str] = [
            "         BOK ID",
            "  Document Name",
            "   Company Name",
            "       Division",
            "         Author",
            "Company Address",
            "   Project Name",
            " Project Number",
            "   End Customer",
            "      Site Name",
            "      File Name",
        ]

        # User Input - bottom text section
        for i, title in enumerate(input_titles, start=4):
            self.user_input = tk.Frame(root)
            self.user_input.grid(row=i, column=0, padx=10, pady=2, sticky="w", columnspan=3)

            tk.Label(master=self.user_input, text=title, width=len(title), anchor="e").grid(
                row=0, column=0, padx=10, pady=5, sticky="e"
            )
            entry = tk.Entry(master=self.user_input, justify="left", width=55)
            entry.grid(row=0, column=1, padx=0, pady=5, sticky="we", columnspan=2)
            entry.grid_columnconfigure(index=0, weight=2)
            self.properties.update({title: entry})
            row_max = i

        # MARK: bottom padding
        row_max += 1
        tk.Label(self.root, text="").grid(row=row_max, column=0, padx=10, pady=0, sticky="w")

        root.grid_rowconfigure(index=1, weight=1)
        root.grid_columnconfigure(index=0, weight=1)
        root.grid_columnconfigure(index=1, weight=1)

    # MARK: Logic
    def toggle_print(self):
        self.print = self.print_checkbox_var.get()

    def select_files(self):
        # Open file dialog to select .insv files
        filetypes: list[tuple[str, str]] = FILE_TYPES
        selected_files: tuple[str, ...] | str = filedialog.askopenfilenames(title="Select files", filetypes=filetypes)

        # Add selected files to the list and update the text box
        for file_path in selected_files:
            if file_path not in self.file_paths:
                file_path = os.path.normpath(file_path)
                self.file_paths.append(file_path)
                self.add_file_to_text_box(file_path)

    def select_parent_folder(self) -> None:
        # Open file dialog to select folder
        parent_folder: str = filedialog.askdirectory(title="Select parent folder to find .docx files in.", mustexist=True)

        # Find correct files to add to file list
        # https://docs.python.org/3/library/os.html#os.walk
        # For dirs in path
        for root, dirs, files in os.walk(top=parent_folder, topdown=True):
            # For file in cur dir
            for file in files:
                # if filename ext is docx
                if file.endswith(".docx"):
                    # get file path
                    file_path: str = os.path.join(root, file)
                    file_path = os.path.normpath(file_path)
                    # add path to file list if not already there
                    if file_path not in self.file_paths:
                        self.file_paths.append(file_path)
                        self.add_file_to_text_box(file_path)

    def add_file_to_text_box(self, file_path) -> None:
        # Create a checkbox and label for the file path
        var = tk.BooleanVar(value=True)
        checkbox = tk.Checkbutton(self.scrollable_frame, variable=var, bg="white")  # Set background to white
        label = tk.Label(self.scrollable_frame, text=file_path, anchor="w", bg="white")  # Set background to white

        # Store the checkbox and its variable
        self.checkboxes.append((var, checkbox, label))

        # Add to the scrollable frame
        checkbox.grid(row=len(self.checkboxes) - 1, column=0, sticky="w")
        label.grid(row=len(self.checkboxes) - 1, column=1, sticky="w")

    def toggle_all(self) -> None:
        # Toggle all checkboxes
        if not self.checkboxes:
            return

        # Determine the new state based on the first checkbox
        new_state: bool = not self.checkboxes[0][0].get()

        for var, checkbox, label in self.checkboxes:
            var.set(new_state)

    def remove_files(self) -> None:
        # Remove all checked files
        remaining_files:list[str] = []
        remaining_checkboxes:list[tuple[tk.BooleanVar,tk.Checkbutton,tk.Label]] = []

        for i, (var, checkbox, label) in enumerate(self.checkboxes):
            if not var.get():
                remaining_files.append(self.file_paths[i])
                remaining_checkboxes.append((var, checkbox, label))

        # Update the file paths and checkboxes
        self.file_paths = remaining_files
        self.checkboxes = remaining_checkboxes

        # Clear the scrollable frame and re-add the remaining files
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        for var, checkbox, label in self.checkboxes:
            checkbox.grid(row=len(self.checkboxes), column=0, sticky="w")
            label.grid(row=len(self.checkboxes), column=1, sticky="w")

    def run_all(self) -> None:
        # Get the selected file paths and pass them to the run_scripts_gui function
        selected_files: list[str] = [self.file_paths[i] for i, (var, _, _) in enumerate(self.checkboxes) if var.get()]
        if selected_files:
            run_scripts_gui(
                file_paths=selected_files, properties={k.strip(): v.get() for k, v in self.properties.items()}, print=self.print
            )
        else:
            messagebox.showwarning(title="No Files Selected", message="Please select at least one file to run.")

    def on_mouse_wheel(self, event:tk.Event) -> None:
        # Handle vertical scrolling
        if event.delta:  # Windows and macOS
            self.canvas.yview_scroll(-1 * (event.delta // 120), "units")
        elif event.num == 4:  # Linux (up)
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:  # Linux (down)
            self.canvas.yview_scroll(1, "units")

    def on_horizontal_mouse_wheel(self, event:tk.Event) -> None:
        # Handle horizontal scrolling
        if event.delta:  # Windows and macOS
            self.canvas.xview_scroll(-1 * (event.delta // 120), "units")
        elif event.num == 4:  # Linux (left)
            self.canvas.xview_scroll(-1, "units")
        elif event.num == 5:  # Linux (right)
            self.canvas.xview_scroll(1, "units")


if __name__ == "__main__":
    root = tk.Tk()
    app = GUIApp(root)
    root.mainloop()
