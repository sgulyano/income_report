import tkinter as tk
from tkinter import ttk, filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import sys

from src import generate_report, program_code


class ExcelFileSelector(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Income Reconciliation 2566")
        bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
        self.iconbitmap(os.path.abspath(os.path.join(bundle_dir, "icon.ico")))
        # self.iconbitmap("icon.ico")
        self.files = []
        
        container = tk.Frame(self)
        container.pack()
        ttk.Label(container, text="หลักสูตร").pack(side=tk.LEFT, expand=True)
        self.combo = ttk.Combobox(container,
            state="readonly",
            values=list(program_code.keys())
        )
        self.combo.pack(side=tk.LEFT, expand=True)

        # Create listbox
        self.listbox = tk.Listbox(self, width=100, height=20, selectmode=tk.MULTIPLE)
        self.listbox.pack(fill=tk.BOTH, expand=True)
        
        # Create add and remove buttons
        add_button = ttk.Button(self, text="Add File", command=self.add_file, style='Blue.TButton')
        add_button.pack(side=tk.LEFT, expand=True)
        remove_button = ttk.Button(self, text="Remove File", command=self.remove_file, style='Red.TButton')
        remove_button.pack(side=tk.LEFT, expand=True)
        report_button = ttk.Button(self, text="Get Report", command=self.process_file, style='Green.TButton')
        report_button.pack(side=tk.LEFT, expand=True)

        # Add drag and drop functionality to the entry widget
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.add_file)

        # Define custom styles using CSS-like syntax
        self.style = ttk.Style()
        self.style.configure('Blue.TButton', font='Arial 10 bold', background='#007bff', foreground='#000', padding=10, borderwidth=0)
        self.style.map('Blue.TButton', background=[('active', '#0069d9')])
        self.style.configure('Red.TButton', font='Arial 10 bold', background='#d92d0f', foreground='#000', padding=10, borderwidth=0)
        self.style.map('Red.TButton', background=[('active', '#ba321a')])
        self.style.configure('Green.TButton', font='Arial 10 bold', background='#28a745', foreground='#000', padding=10, borderwidth=0)
        self.style.map('Green.TButton', background=[('active', '#218838')])


    def add_file(self, event=None):
        # Open file dialog to select Excel files
        if not event:
            filetypes = (("Excel files", "*.xlsx"), ("All files", "*.*"))
            selected_files = filedialog.askopenfilenames(filetypes=filetypes)
        else:
            selected_files = self.tk.splitlist(event.data)
        
        # Add selected files to listbox
        for file in selected_files:
            if file not in self.files:
                self.files.append(file)
                filename = os.path.basename(file)
                self.listbox.insert(tk.END, filename)
                
    def remove_file(self):
        # Remove selected files from listbox and files list
        selection = self.listbox.curselection()
        for index in reversed(selection):
            self.listbox.delete(index)
            del self.files[index]
    
    def process_file(self):
        if self.combo.get() == '':
            tk.messagebox.showerror(title="Error", message="Please select a programme first.")
            return
        if len(self.files) == 0:
            tk.messagebox.showerror(title="Error", message="Please select a file first.")
            return
        
        for f in self.files:
            if not os.path.isfile(f):
                tk.messagebox.showerror(title="Error", message="Please select a valid file.")
                return
        output_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")])
        
        if output_path:
            # try:
            code, campus, prog_name = program_code[self.combo.get()]
            result = generate_report(self.files, output_path, code, campus, prog_name)
            if result == -1:
                tk.messagebox.showerror(title="Error", message="Error in generating report.")
                return
            elif result == -2:
                tk.messagebox.showerror(title="Error", message="No records for the selected program.")
                return
            # except Exception as e:
            #     print(e)
            #     tk.messagebox.showerror(title="Error", message="Error in generating report. Contact Aj.Sarun.")
            #     return
            tk.messagebox.showinfo(title="Success", message=f"File processed and output saved to {output_path}.")
        return

if __name__ == "__main__":
    app = ExcelFileSelector()
    app.mainloop()