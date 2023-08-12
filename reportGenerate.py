import tkinter as tk
from tkinter import ttk, filedialog
import tkinter.messagebox as messagebox
from image_replace import create_name_dict, fully_update_docx
import zipfile
import shutil
import os
from PIL import Image

class ReportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Report Generator")

        self.create_widgets()

    def create_widgets(self):
        # Create a Notebook widget
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        # Add Input Tab
        input_frame = ttk.Frame(self.notebook)
        self.notebook.add(input_frame, text="Input")

        # Select the folder containing new images
        folder_label = tk.Label(input_frame, text="Select the folder containing new images:")
        folder_label.grid(row=0, column=0, sticky="e")

        self.folder_var = tk.StringVar()
        folder_entry = tk.Entry(input_frame, textvariable=self.folder_var)
        folder_entry.grid(row=0, column=1)

        folder_button = tk.Button(input_frame, text="Browse", command=self.browse_folder)
        folder_button.grid(row=0, column=2)

        # Select the template report (.docx)
        template_label = tk.Label(input_frame, text="Select the template report (.docx):")
        template_label.grid(row=1, column=0, sticky="e")

        self.template_var = tk.StringVar()
        template_entry = tk.Entry(input_frame, textvariable=self.template_var)
        template_entry.grid(row=1, column=1)

        template_button = tk.Button(input_frame, text="Browse", command=self.browse_template)
        template_button.grid(row=1, column=2)

        # Select the txt file with image names
        image_names_label = tk.Label(input_frame, text="Select the txt file with image names:")
        image_names_label.grid(row=2, column=0, sticky="e")

        self.image_names_var = tk.StringVar()
        image_names_entry = tk.Entry(input_frame, textvariable=self.image_names_var)
        image_names_entry.grid(row=2, column=1)

        image_names_button = tk.Button(input_frame, text="Browse", command=self.browse_image_names)
        image_names_button.grid(row=2, column=2)

        # Enter the output report name (without extension)
        output_name_label = tk.Label(input_frame, text="Output report name (without extension):")
        output_name_label.grid(row=3, column=0, sticky="e")

        self.output_name_var = tk.StringVar()
        output_name_entry = tk.Entry(input_frame, textvariable=self.output_name_var)
        output_name_entry.grid(row=3, column=1)

        # Select the output folder for generated reports
        output_folder_label = tk.Label(input_frame, text="Select the output folder:")
        output_folder_label.grid(row=4, column=0, sticky="e")

        self.output_folder_var = tk.StringVar()
        output_folder_entry = tk.Entry(input_frame, textvariable=self.output_folder_var)
        output_folder_entry.grid(row=4, column=1)

        output_folder_button = tk.Button(input_frame, text="Browse", command=self.browse_output_folder)
        output_folder_button.grid(row=4, column=2)

        # Add Help Tab
        help_frame = ttk.Frame(self.notebook)
        self.notebook.add(help_frame, text="Help")

        help_text = """
        - Folder containing new images: Select the folder where your new images are located.
        - Template report (.docx): Choose the template report in .docx format.
        - Txt file with image names: Select a text file containing the list of image names.
        - Output report name: Enter the desired name for the generated report (without extension).
        - Output folder: Select the folder where the generated reports will be saved.
        """
        help_label = tk.Label(help_frame, text=help_text, justify="left")
        help_label.grid(row=0, column=0, padx=10, pady=10)

        # Generate Report button
        run_button = tk.Button(self.root, text="Generate Report", command=self.generate_reports)
        run_button.grid(row=1, column=1, columnspan=2, pady=10)

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        self.folder_var.set(folder_path)

    def browse_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        self.template_var.set(file_path)

    def browse_image_names(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        self.image_names_var.set(file_path)
        
    def browse_output_folder(self):
        folder_path = filedialog.askdirectory()
        self.output_folder_var.set(folder_path)

    def generate_reports(self):
        # Get input values
        folder_path = self.folder_var.get()
        template_path = self.template_var.get()
        image_names_path = self.image_names_var.get()
        output_name = self.output_name_var.get()
        output_folder = self.output_folder_var.get()

        # Validate inputs
        if not folder_path:
            tk.messagebox.showerror("Error", "Please select the folder containing new images.")
            return
        if not template_path:
            tk.messagebox.showerror("Error", "Please select the template report (.docx).")
            return
        if not image_names_path:
            tk.messagebox.showerror("Error", "Please select the txt files that show to and from image names.")
            return
        if not output_name:
            tk.messagebox.showerror("Error", "Please enter the output report name.")
            return
        if not output_folder:
            tk.messagebox.showerror("Error", "Please select the output folder.")
            return
        if not os.path.exists(folder_path):
            tk.messagebox.showerror("Error", "Please provide a valid path to the folder containing new images.")
            return
        if not os.path.exists(template_path):
            tk.messagebox.showerror("Error", "Please provide a valid template report (.docx).")
            return
        if not template_path.endswith('.docx'):
            tk.messagebox.showerror("Error", "Please provide word (.docx) file.")
            return
        if not os.path.exists(image_names_path):
            tk.messagebox.showerror("Error", "Please select a valid txt file that show to and from image names.")
            return
        if not os.path.exists(output_folder):
            tk.messagebox.showerror("Error", "Please provide a valid output folder.")
            return

        # Now you can proceed with generating the reports
        print(f"Generating report '{output_name}' in folder '{output_folder}'...")
        # Call your custom report generation function here
        rename_dict = create_name_dict(image_names_path)
    
        fully_update_docx(output_name, template_path, folder_path, output_folder, rename_dict)


if __name__ == "__main__":
    root = tk.Tk()
    app = ReportGeneratorApp(root)
    root.mainloop()
