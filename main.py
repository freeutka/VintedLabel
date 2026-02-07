import tkinter as tk
from tkinter import ttk, messagebox
import win32print
from datetime import datetime

class VintedLabel:
    def __init__(self, root):
        self.root = root
        self.root.title("Vinted Label Creator")
        self.root.geometry("850x650")
        self.root.configure(bg="#1e1e1e")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TCombobox", 
                        fieldbackground="#2d2d2d", 
                        background="#2d2d2d", 
                        foreground="white",
                        bordercolor="#1e1e1e",
                        darkcolor="#1e1e1e",
                        lightcolor="#1e1e1e",
                        arrowcolor="white") 
        
        style.map("TCombobox", 
                  fieldbackground=[('readonly', '#2d2d2d')],
                  foreground=[('readonly', 'white')],
                  selectbackground=[('readonly', '#007acc')], 
                  selectforeground=[('readonly', 'white')])
        
        self.root.option_add("*TCombobox*Listbox.background", "#2d2d2d")
        self.root.option_add("*TCombobox*Listbox.foreground", "white")
        self.root.option_add("*TCombobox*Listbox.selectBackground", "#007acc")
        self.root.option_add("*TCombobox*Listbox.selectForeground", "white")
        self.root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 11))
        
        self.vars = {}
        self.sn_vars = [tk.StringVar() for _ in range(4)]
        self.setup_ui()
        self.update_preview()

    def setup_ui(self):
        input_frame = tk.Frame(self.root, bg="#1e1e1e", padx=20, pady=20)
        input_frame.pack(side="left", fill="both", expand=True)

        self.create_label(input_frame, "ITEM:")
        self.vars['item_var'] = tk.StringVar(value="")
        tk.Entry(input_frame, textvariable=self.vars['item_var'], font=("Segoe UI", 11), 
                 bg="#2d2d2d", fg="white", borderwidth=0).pack(fill="x", pady=(0, 10), ipady=4)

        self.create_label(input_frame, "COND:")
        cond_options = ["New item with tags", "New item without tags", "Very good", "Good", "Satisfactory", "Not fully functional"]
        self.vars['cond_var'] = tk.StringVar(value="Very good")
        cond_combo = ttk.Combobox(input_frame, textvariable=self.vars['cond_var'], 
                                  values=cond_options, state="readonly", font=("Segoe UI", 11))
        cond_combo.pack(fill="x", pady=(0, 10), ipady=4)
        cond_combo.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        self.create_label(input_frame, "S/N (Last 4):")
        sn_frame = tk.Frame(input_frame, bg="#1e1e1e")
        sn_frame.pack(fill="x", pady=(0, 10))
        
        self.sn_entries = []
        for i in range(4):
            e = tk.Entry(sn_frame, textvariable=self.sn_vars[i], width=3, font=("Consolas", 14, "bold"), 
                         justify="center", bg="#2d2d2d", fg="#007acc", insertbackground="white", borderwidth=0)
            e.pack(side="left", padx=(0, 10))
            e.bind("<KeyRelease>", lambda event, idx=i: self.on_sn_key(event, idx))
            self.sn_entries.append(e)
            self.sn_vars[i].trace_add("write", lambda *args: self.update_preview())

        other_fields = [
            ("STATUS:", "status_var", ""),
            ("NOTES:", "notes_var", ""),
            ("HAND:", "hand_var", "")
        ]

        for label_text, var_name, default in other_fields:
            self.create_label(input_frame, label_text)
            self.vars[var_name] = tk.StringVar(value=default)
            entry = tk.Entry(input_frame, textvariable=self.vars[var_name], font=("Segoe UI", 11), 
                             bg="#2d2d2d", fg="white", borderwidth=0)
            entry.pack(fill="x", pady=(0, 10), ipady=4)
            self.vars[var_name].trace_add("write", lambda *args: self.update_preview())

        self.print_btn = tk.Button(input_frame, text="PRINT", command=self.send_to_printer, 
                                   bg="#007acc", fg="white", font=("Segoe UI", 12, "bold"), 
                                   relief="flat", cursor="hand2", pady=10)
        self.print_btn.pack(fill="x", side="bottom", pady=10)

        preview_container = tk.Frame(self.root, bg="#252526", padx=20, pady=20)
        preview_container.pack(side="right", fill="both")
        tk.Label(preview_container, text="LIVE PREVIEW", bg="#252526", fg="#95a5a6").pack(pady=(0, 10))
        
        self.preview_text = tk.Text(preview_container, font=("Consolas", 10), width=32, height=28, 
                                    bg="white", fg="black", padx=10, pady=10, relief="flat")
        self.preview_text.pack()

    def create_label(self, parent, text):
        tk.Label(parent, text=text, bg="#1e1e1e", fg="#888888", font=("Consolas", 10, "bold")).pack(anchor="w")

    def on_sn_key(self, event, idx):
        if len(self.sn_vars[idx].get()) > 0:
            if idx < 3:
                self.sn_entries[idx+1].focus_set()
        if event.keysym == "BackSpace" and len(self.sn_vars[idx].get()) == 0:
            if idx > 0:
                self.sn_entries[idx-1].focus_set()

    def generate_receipt(self):
        date_now = datetime.now().strftime("%d/%m/%Y %H:%M")
        sn_full = "".join([v.get() for v in self.sn_vars]).upper()
        
        receipt = [
            "================================",
            f"ITEM: {self.vars['item_var'].get().upper()}",
            f"COND: {self.vars['cond_var'].get()}",
            f"S/N (Lats 4): {sn_full}",
            f"STAT: {self.vars['status_var'].get()}",
            f"NOTE: {self.vars['notes_var'].get()}",
            "--------------------------------",
            f"PACKAGED ON: {date_now}",
            f"HAND: {self.vars['hand_var'].get()}",
            "================================",
            "      THANK YOU FOR ORDER!      ",
            "  If any issues - message me!   ",
            "   Please leave a 5* review!    ",
            "================================"
        ]
        return "\n".join(receipt)

    def update_preview(self):
        self.preview_text.config(state="normal")
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", self.generate_receipt())
        self.preview_text.config(state="disabled")

    def send_to_printer(self):
        try:
            printer_name = win32print.GetDefaultPrinter()
            hPrinter = win32print.OpenPrinter(printer_name)
            win32print.StartDocPrinter(hPrinter, 1, ("VintedLabel", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            
            final_content = self.generate_receipt() + "\n\n\n\n"
            win32print.WritePrinter(hPrinter, final_content.encode('cp866'))
            
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            win32print.ClosePrinter(hPrinter)
            
            self.print_btn.config(text="SENT TO PRINTER!", bg="#4ec9b0")
            self.root.after(2000, lambda: self.print_btn.config(text="PRINT", bg="#007acc"))
        except Exception as e:
            messagebox.showerror("Printer Error", f"Error:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = VintedLabel(root)
    root.mainloop()
