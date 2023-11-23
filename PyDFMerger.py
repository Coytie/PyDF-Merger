import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkFont
import subprocess
import sys

class PDFMergerGUI:
    def __init__(self, master):
        self.master = master
        master.title("Brenden's PDF Merger")
        master.geometry("450x300")  # Set width and height
        master.resizable(False, False)  # Lock the size
        master.configure(bg='white')

        self.input_files = []
        self.last_saved_label = tk.StringVar()

        # Create and set up menu bar
        self.menu_bar = tk.Menu(master)
        self.master.config(menu=self.menu_bar)  # Corrected line

        # Create "File" menu
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="Exit", command=master.destroy)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)

        # Create "About" menu
        self.about_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.about_menu.add_command(label="About", command=self.show_about)
        self.menu_bar.add_cascade(label="About", menu=self.about_menu)

        # Create and set up widgets
        self.label = tk.Label(master, text="Brenden's Simple PDF Merger!", bg='white')
        self.label.pack()
        self.label = tk.Label(master, text="Select PDFs to merge:", bg='white')
        self.label.pack()

        # Create a frame to hold the Listbox, scrollbar, and buttons
        main_frame = tk.Frame(master, bg='white')
        main_frame.pack()

        # Create a horizontal scrollbar
        self.xscrollbar = tk.Scrollbar(main_frame, orient=tk.HORIZONTAL)
        self.xscrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Create a Listbox with horizontal scrolling
        self.listbox = tk.Listbox(main_frame, selectmode=tk.MULTIPLE, width=50, xscrollcommand=self.xscrollbar.set, bg='white')
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Attach the scrollbar to the Listbox
        self.xscrollbar.config(command=self.listbox.xview)

        # Create a frame to hold the buttons
        button_frame = tk.Frame(main_frame, bg='white')
        button_frame.pack(side=tk.LEFT, padx=5)

        # "Add PDF" button
        self.add_button = tk.Button(button_frame, text=" + ", command=self.add_pdf, width=5)
        self.add_button.pack(pady=5)

        # "Move Up" button
        self.move_up_button = tk.Button(button_frame, text=" ↑ ", command=self.move_up, width=5)
        self.move_up_button.pack(pady=5)

        # "Move Down" button
        self.move_down_button = tk.Button(button_frame, text=" ↓ ", command=self.move_down, width=5)
        self.move_down_button.pack(pady=5)

        # "Remove" button
        self.remove_button = tk.Button(button_frame, text=" - ", command=self.remove_pdf, width=5)
        self.remove_button.pack(pady=5)

        # "Merge PDFs" button
        self.merge_button = tk.Button(button_frame, text="Merge", command=self.merge_pdfs, width=5)
        self.merge_button.pack(pady=5)

        # Create a label to display the last saved path
        self.last_saved_path_label = tk.Label(master, textvariable=self.last_saved_label, wraplength=450, bg='white')
        self.last_saved_path_label.pack()

    def add_pdf(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if file_paths:
            self.input_files.extend(file_paths)
            for file_path in file_paths:
                self.listbox.insert(tk.END, file_path)

    def remove_pdf(self):
        selected_indices = self.listbox.curselection()
        if selected_indices:
            for index in reversed(selected_indices):
                self.input_files.pop(index)
                self.listbox.delete(index)

    def move_up(self):
        selected_indices = self.listbox.curselection()
        if selected_indices:
            for index in selected_indices:
                if index > 0:
                    self.input_files[index], self.input_files[index - 1] = (
                        self.input_files[index - 1],
                        self.input_files[index],
                    )
                    self.listbox.delete(index)
                    self.listbox.insert(index - 1, self.input_files[index - 1])
                    self.listbox.select_clear(index)
                    self.listbox.select_set(index - 1)
                    self.listbox.activate(index - 1)

    def move_down(self):
        selected_indices = self.listbox.curselection()
        if selected_indices:
            for index in reversed(selected_indices):
                if index < self.listbox.size() - 1:
                    self.input_files[index], self.input_files[index + 1] = (
                        self.input_files[index + 1],
                        self.input_files[index],
                    )
                    self.listbox.delete(index)
                    self.listbox.insert(index + 1, self.input_files[index + 1])
                    self.listbox.select_clear(index)
                    self.listbox.select_set(index + 1)
                    self.listbox.activate(index + 1)

    def merge_pdfs(self):
        if not self.input_files:
            messagebox.showerror("Error", "Please add at least one input PDF file.")
            return

        try:
            # Ask user for the output file location and name
            output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

            if not output_pdf:
                # User canceled the save dialog
                return

            # Create a PDF merger object
            pdf_merger = PyPDF2.PdfMerger()

            # Add PDFs to the merger
            for pdf in self.input_files:
                pdf_merger.append(pdf)

            # Write the merged PDF to the output file
            with open(output_pdf, 'wb') as output_file:
                pdf_merger.write(output_file)

            # Update the label with the last saved path
            self.last_saved_label.set(f'Last file saved to: {output_pdf}')

            messagebox.showinfo("Success", f'Merged PDFs successfully. Output saved to: {output_pdf}')

            # Open the merged PDF using the default PDF viewer
            if sys.platform == 'win32':
                subprocess.Popen(["start", "", output_pdf], shell=True)
            elif sys.platform == 'darwin':
                subprocess.Popen(["open", output_pdf], shell=True)
            elif sys.platform.startswith('linux'):
                subprocess.Popen(["xdg-open", output_pdf], shell=True)

        except Exception as e:
            messagebox.showerror("Error", f'Error merging PDFs: {e}')

    def show_about(self):
        about_text = "Brenden's PDF Merger\nVersion 1.0\n\n© 2023 Brenden IT"
        messagebox.showinfo("About", about_text)

if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap('c:/path/to/icon.ico')
    app = PDFMergerGUI(root)
    root.mainloop()
