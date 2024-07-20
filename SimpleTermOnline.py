import os
import sys
import subprocess
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
import pyperclip
import gspread
from oauth2client.service_account import ServiceAccountCredentials


class SimpleTermOnline:
    def __init__(self, root, sheet_id=None):
        self.root = root
        self.sheet_id = sheet_id or '1bl5koCGUwI_qKnSVCO10e1VJLAS5V8kb6j52tKtl6Dk'
        self.sheet = None
        self.df = None
        self.results = []
        self.current_index = 0
        self.current_search_term = ""

        self.authenticate_google_sheets()
        self.load_sheet()
        self.setup_gui()

    def authenticate_google_sheets(self):
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name('simpletermonline-5f4cf21da8ae.json', scope)
            client = gspread.authorize(creds)
            self.sheet = client.open_by_key(self.sheet_id).sheet1
        except Exception as e:
            messagebox.showerror("Error", f"Error authenticating with Google Sheets:\n{str(e)}")
            sys.exit()

    def load_sheet(self):
        try:
            data = self.sheet.get_all_records()
            self.df = pd.DataFrame(data)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Google Sheet data:\n{str(e)}")
    
    def refresh_sheet(self, event=None):
        try:
            self.load_sheet()
            
            if self.df is not None:
                self.results = [] 
                file_name = os.path.basename(self.file_path) if self.file_path else "Google Sheet"
                self.root.title(f"SimpleTerm Online")
                self.root.after(2000, lambda: self.root.title(f"SimpleTerm Online"))
                self.root.after(0, lambda: self.root.title("Updated!"))
            else:
                messagebox.showwarning("Update Error", "Failed to refresh data from Google Sheets.")
        except Exception as e:
            messagebox.showerror("Error", f"Error refreshing data from Google Sheets:\n{str(e)}")


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
            messagebox.showwarning("Input Error", "Please enter a term to search or load the Google Sheet first.")

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
        if self.sheet_id:
            self.root.title(f"SimpleTerm Online")

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
        self.root.bind('<Control-o>', lambda event: self.open_google_sheet())
        self.root.bind('<Right>', self.navigate_results)
        self.root.bind('<Left>', self.navigate_results)
        self.root.bind('<Tab>', self.navigate_results)
        self.root.bind('<Control-c>', self.copy_result_term)
        self.root.bind('<Control-plus>', self.increase_font_size)
        self.root.bind('<Control-minus>', self.decrease_font_size)
        self.root.bind('<Control-Shift-plus>', self.increase_notes_font_size)
        self.root.bind('<Control-Shift-minus>', self.decrease_notes_font_size)
        self.root.bind('<F5>', self.refresh_google_sheet)
        self.root.bind('<F3>', self.display_help)
        self.root.bind('<F2>', self.change_google_sheet)

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

    def refresh_google_sheet(self, event=None):
        self.load_google_sheet()
        if self.df is not None:
            self.root.title(f"SimpleTerm Online")
            self.root.after(2000, lambda: self.root.title(f"SimpleTerm Online"))
            self.root.after(0, lambda: self.root.title("Updated!"))  # Display "Updated!" briefly

    def change_google_sheet(self, event=None):
        sheet_id = simpledialog.askstring("Input", "Enter new Google Sheet ID:")
        if sheet_id:
            self.sheet_id = sheet_id
            self.load_google_sheet()
            if self.df is not None:
                self.root.title(f"SimpleTerm Online")
                self.results = []
                self.update_display()
                self.root.after(2000, lambda: self.root.title(f"SimpleTerm Online"))
                self.root.after(0, lambda: self.root.title("Sheet Changed!"))  # Display "Sheet Changed!" briefly

    def display_help(self, event=None):
        help_text = (
            "Shortcuts\n"
            "Tab Key, Left and Right arrows: Navigate between results\n"
            "Ctrl+C: Copy the displayed result\n"
            "Ctrl+N: Add a new term\n"
            "Ctrl+O: Open the used Google Sheet\n"
            "Ctrl++/-: Increase or decrease the font size\n"
            "Ctrl+Shift++/-: Increase or decrease the notes font size\n"
            "F5: Refresh the Google Sheet\n"
            "F3: Display this help\n"
            "F2: Change the Google Sheet ID\n"
        )
        messagebox.showinfo("Help", help_text)


    def open_add_term_dialog(self, event=None):
        def save_term():
            source_term = source_entry.get().strip()
            target_term = target_entry.get().strip()
            notes = notes_entry.get().strip()

            if source_term and target_term:
                new_data = {
                    'Source Term': source_term,
                    'Target Term': target_term,
                    'Notes': notes if notes else ''
                }

                try:
                    self.sheet.append_row([source_term, target_term, new_data['Notes']])
                    messagebox.showinfo("Success", "New term added successfully.")
                    new_term_dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Error", f"Error saving to Google Sheet:\n{str(e)}")
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
        def open_google_sheet(self, event=None):
            url = f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/edit"
            if sys.platform.startswith('win'):
                os.startfile(url)
            elif sys.platform.startswith('darwin'):
                subprocess.call(['open', url])
            else:
                subprocess.call(['xdg-open', url])


if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleTermOnline(root)
    root.mainloop()
