import os
import sys
import subprocess
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
import pyperclip


class SimpleTerm:
    def __init__(self, root, file_path=None):
        self.root = root
        self.file_path = file_path
        self.df = None
        self.results = []
        self.current_index = 0
        self.current_search_term = ""

        if self.file_path is None or not self.file_path.endswith('.xlsx'):
            self.select_excel_file()
        else:
            self.load_excel()
            self.setup_gui()

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", ".xlsx .xls")],
            title="Select an Excel file as the termbase"
        )
        if file_path:
            self.file_path = file_path
            self.load_excel()
            if self.df is not None:
                self.setup_gui()
        else:
           sys.exit()

    def load_excel(self):
        try:
            self.df = pd.read_excel(self.file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file:\n{str(e)}")

    def find_equivalent(self, term):
        results = []
        if self.df is not None:
            for index, row in self.df.iterrows():
                if row['Source Term'].lower() == term.lower():
                    notes = row['Notes']
                    if pd.isna(notes):
                        notes = ""  # Replace NaN with an empty string
                    results.append({
                        'target_term': row['Target Term'],
                        'notes': notes
                    })
        return results

    def search_term(self, event=None):
        term = self.entry.get().strip()

        if term and self.df is not None:
            try:
                self.results = self.find_equivalent(term)
                self.current_index = 0
                self.update_display()
                self.current_search_term = term  # Store current search term for Add Term dialog
            except Exception as e:
                print(f"Error occurred: {str(e)}")  # Debug print
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        else:
            messagebox.showwarning("Input Error", "Please enter a term to search or load the Excel file first.")

    def update_display(self):
        if self.results:
            result = self.results[self.current_index]
            self.result_label.config(text=result['target_term'])

            notes_text = result['notes'] if result['notes'] else ""
            self.notes_text.config(state=tk.NORMAL)
            self.notes_text.delete(1.0, tk.END)
            self.notes_text.insert(tk.END, notes_text)
            self.notes_text.config(state=tk.DISABLED)

            self.result_label.config(fg='blue' if len(self.results) > 1 else 'black')
        else:
            self.result_label.config(text='Term not found.')
            self.notes_text.config(state=tk.NORMAL)
            self.notes_text.delete(1.0, tk.END)
            self.notes_text.config(state=tk.DISABLED)

        self.root.update_idletasks()
        self.root.geometry(f"{self.root.winfo_width()}x{self.root.winfo_height()}")

    def setup_gui(self):
        if self.file_path:
            file_name = os.path.basename(self.file_path)
            self.root.title(f"SimpleTerm: {file_name}")

        self.root.attributes('-topmost', True)
        self.root.configure(bg='#f0f0f0')

        input_frame = tk.Frame(self.root, padx=10, pady=10, bg='#f0f0f0')
        input_frame.pack(padx=10, pady=10, fill=tk.X)
        input_frame.pack_propagate(False)

        self.entry = tk.Entry(input_frame, font=('Arial', 14), relief=tk.FLAT, width=20)
        self.entry.grid(row=0, column=0, padx=(0, 10), sticky='ew')

        self.result_label = tk.Label(input_frame, text='', font=('Arial', 14), anchor='w', bg='#f0f0f0')
        self.result_label.grid(row=0, column=1, padx=(10, 0), sticky='ew')

        input_frame.grid_columnconfigure(0, weight=0)

        result_frame = tk.Frame(self.root, padx=10, pady=0, bg='#f0f0f0')
        result_frame.pack(padx=10, pady=0, fill=tk.BOTH, expand=True)
        self.notes_text = tk.Text(result_frame, wrap=tk.WORD, font=('Arial', 12), state=tk.DISABLED, bg='#f0f0f0', relief=tk.FLAT)
        self.notes_text.pack(fill=tk.BOTH, expand=True)
        result_frame.pack_propagate(False)

        self.entry.bind('<Return>', self.search_term)
        self.root.bind('<Control-n>', lambda event: self.open_add_term_dialog())
        self.root.bind('<Control-o>', lambda event: self.open_excel_file())
        self.root.bind('<Right>', self.navigate_results)
        self.root.bind('<Left>', self.navigate_results)
        self.root.bind('<Tab>', self.navigate_results)
        self.root.bind('<Control-c>', self.copy_result_term)
        self.root.bind('<Control-plus>', self.increase_font_size)
        self.root.bind('<Control-minus>', self.decrease_font_size)
        self.root.bind('<Control-Shift-plus>', self.increase_notes_font_size)
        self.root.bind('<Control-Shift-minus>', self.decrease_notes_font_size)
        self.root.bind('<F5>', self.refresh_excel)
        self.root.bind('<F3>', self.display_help)
        self.root.bind('<F2>', self.change_excel_file)

        self.entry.focus_set()
        self.root.after(100, self.entry.focus_force)

    def navigate_results(self, event):
        if self.results:
            if event.keysym == 'Right' or event.keysym == 'Tab':
                self.current_index = (self.current_index + 1) % len(self.results)
            elif event.keysym == 'Left':
                self.current_index = (self.current_index - 1) % len(self.results)
            self.update_display()

    def copy_result_term(self, event):
        if self.results:
            pyperclip.copy(self.results[self.current_index]['target_term'])
            self.result_label.config(bg='green')
            self.root.after(500, lambda: self.result_label.config(bg=self.root.cget('bg')))

    def increase_font_size(self, event=None):
        current_font_size = int(self.entry.cget('font').split()[1])
        new_font_size = current_font_size + 1
        self.entry.config(font=('Arial', new_font_size))
        self.result_label.config(font=('Arial', new_font_size))

    def decrease_font_size(self, event=None):
        current_font_size = int(self.entry.cget('font').split()[1])
        new_font_size = max(current_font_size - 1, 8)
        self.entry.config(font=('Arial', new_font_size))
        self.result_label.config(font=('Arial', new_font_size))

    def increase_notes_font_size(self, event=None):
        current_font_size = int(self.notes_text.cget('font').split()[1])
        new_font_size = current_font_size + 1
        self.notes_text.config(font=('Arial', new_font_size))

    def decrease_notes_font_size(self, event=None):
        current_font_size = int(self.notes_text.cget('font').split()[1])
        new_font_size = max(current_font_size - 1, 8)
        self.notes_text.config(font=('Arial', new_font_size))

    def refresh_excel(self, event=None):
        self.load_excel()
        if self.df is not None:
            file_name = os.path.basename(self.file_path) if self.file_path else ""
            self.root.title(f"SimpleTerm: {file_name}")
            self.root.after(2000, lambda: self.root.title(f"SimpleTerm: {file_name}"))
            self.root.after(0, lambda: self.root.title("Updated!"))  # Display "Updated!" briefly

    def change_excel_file(self, event=None):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", ".xlsx .xls")],
            title="Select a new Excel file"
        )
        if file_path:
            self.file_path = file_path
            self.load_excel()
            if self.df is not None:
                file_name = os.path.basename(self.file_path)
                self.root.title(f"SimpleTerm: {file_name}")
                self.results = []
                self.update_display()
                self.root.after(2000, lambda: self.root.title(f"SimpleTerm: {file_name}"))
                self.root.after(0, lambda: self.root.title("File Changed!"))  # Display "File Changed!" briefly

    def display_help(self, event=None):
        help_text = (
            "Shortcuts\n"
            "Tab Key, Left and Right arrows: Navigate between results\n"
            "Ctrl+C: Copy the displayed result\n"
            "Ctrl+N: Add a new term\n"
            "Ctrl+O: Open the used Excel file\n"
            "Ctrl++/-: Increase or decrease the font size\n"
            "Ctrl+Shift++/-: Increase or decrease the font size of notes\n"
            "F2: Change the Excel file used\n"
            "F3: Open this help message\n"
            "F5: Refresh the Excel database\n"
        )
        messagebox.showinfo("Help", help_text)

    def open_add_term_dialog(self, event=None):
        new_term_dialog = tk.Toplevel(self.root)
        new_term_dialog.title("Add New Term")
        new_term_dialog.configure(bg='#f0f0f0')
        new_term_dialog.focus()

        # Get the position and dimensions of the main window
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()

        # Define dimensions for the new dialog
        dialog_width = 400
        dialog_height = 100

        # Calculate position for the new dialog to center it on the main window
        x = main_x + (main_width // 2) - (dialog_width // 2)
        y = main_y + (main_height // 2) - (dialog_height // 2)

        # Set the geometry for the new dialog
        new_term_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        source_label = tk.Label(new_term_dialog, text="Source:", font=('Arial', 12))
        source_label.grid(row=0, column=0, padx=10, pady=5, sticky='e')
        source_entry = tk.Entry(new_term_dialog, font=('Arial', 12), relief=tk.FLAT)
        source_entry.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        source_entry.insert(0, self.current_search_term)  # Pre-fill with current search term

        target_label = tk.Label(new_term_dialog, text="Target:", font=('Arial', 12))
        target_label.grid(row=1, column=0, padx=10, pady=5, sticky='e')
        target_entry = tk.Entry(new_term_dialog, font=('Arial', 12), relief=tk.FLAT)
        target_entry.grid(row=1, column=1, padx=10, pady=5, sticky='w')

        notes_label = tk.Label(new_term_dialog, text="Notes (optional):", font=('Arial', 12))
        notes_label.grid(row=2, column=0, padx=10, pady=5, sticky='ne')
        notes_entry = tk.Entry(new_term_dialog, font=('Arial', 12), relief=tk.FLAT)
        notes_entry.grid(row=2, column=1, padx=10, pady=5, sticky='we')

        # Set focus to the source_entry field
        source_entry.focus()

        def save_term():
            source_term = source_entry.get().strip()
            target_term = target_entry.get().strip()
            notes = notes_entry.get().strip()

            if source_term and target_term:
                new_data = pd.DataFrame({
                    'Source Term': [source_term],
                    'Target Term': [target_term],
                    'Notes': [notes if notes else pd.NA]
                })

                self.df = pd.concat([self.df, new_data], ignore_index=True)

                try:
                    self.df.to_excel(self.file_path, index=False)
                    messagebox.showinfo("Success", "New term added successfully.")
                    new_term_dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Error", f"Error saving to Excel file. Please make sure the file is not being used.\n{str(e)}")
            else:
                messagebox.showwarning("Missing Fields", "Please enter Source Term and Target Term.")

        new_term_dialog = tk.Toplevel(self.root)
        new_term_dialog.title("Add New Term")

        dialog_width = 400
        dialog_height = 100

        screen_width = new_term_dialog.winfo_screenwidth()
        screen_height = new_term_dialog.winfo_screenheight()

        x = (screen_width // 2) - (dialog_width // 2)
        y = (screen_height // 2) - (dialog_height // 2)

        new_term_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        tk.Label(new_term_dialog, text="Source Term:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        source_entry = tk.Entry(new_term_dialog, width=40, relief=tk.FLAT)
        source_entry.grid(row=0, column=1, padx=10, pady=5)
        source_entry.insert(0, self.current_search_term)

        tk.Label(new_term_dialog, text="Target Term:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        target_entry = tk.Entry(new_term_dialog, width=40, relief=tk.FLAT)
        target_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(new_term_dialog, text="Notes:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
        notes_entry = tk.Entry(new_term_dialog, width=40, relief=tk.FLAT)
        notes_entry.grid(row=2, column=1, padx=10, pady=5)

        new_term_dialog.bind('<Return>', lambda event: save_term())
        new_term_dialog.bind('<Escape>', lambda event: new_term_dialog.destroy())

        new_term_dialog.transient(self.root)
        new_term_dialog.grab_set()
        new_term_dialog.focus_set()
        source_entry.focus()

        self.root.wait_window(new_term_dialog)

    def open_excel_file(self):
        if self.file_path:
            try:
                subprocess.Popen([self.file_path], shell=True)
            except Exception as e:
                messagebox.showerror("Error", f"Error opening Excel file:\n{str(e)}")
        else:
            messagebox.showwarning("No Excel File", "No Excel file is currently loaded.")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleTerm(root)
    root.mainloop()