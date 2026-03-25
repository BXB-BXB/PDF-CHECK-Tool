import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor v10 - Bogdan Bahrim")
        self.root.geometry("1100x800")
        self.root.configure(bg="#f0f2f5")

        self.full_results = []
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0)

        # --- Control Panel (Ca la V10) ---
        top = tk.Frame(root, bg="#2c3e50", pady=10)
        top.pack(fill=tk.X)
        
        tk.Button(top, text="1. Load Excel", command=self.load_excel).pack(side=tk.LEFT, padx=10)
        tk.Button(top, text="2. Load PDF", command=self.load_pdf).pack(side=tk.LEFT, padx=10)
        
        tk.Label(top, text="Skip Pgs:", fg="white", bg="#2c3e50").pack(side=tk.LEFT, padx=(20,0))
        self.exclude_entry = tk.Entry(top, width=8); self.exclude_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(top, text="Color", command=self.pick_color).pack(side=tk.LEFT, padx=10)
        
        self.run_btn = tk.Button(top, text="START AUDIT", command=self.start_thread, bg="#27ae60", fg="white", state=tk.DISABLED)
        self.run_btn.pack(side=tk.RIGHT, padx=10)

        # --- Functionalitate Nume + Sufix (Imbunatatire) ---
        export_bar = tk.Frame(root, bg="#dfe6e9", pady=5)
        export_bar.pack(fill=tk.X)
        tk.Label(export_bar, text="Nume Fisier:", bg="#dfe6e9").pack(side=tk.LEFT, padx=5)
        self.base_name_var = tk.StringVar(value="Audit_Result")
        tk.Entry(export_bar, textvariable=self.base_name_var, width=25).pack(side=tk.LEFT, padx=5)
        tk.Label(export_bar, text="+ Suffix:", bg="#dfe6e9").pack(side=tk.LEFT, padx=5)
        self.suffix_var = tk.StringVar(value="_v01")
        tk.Entry(export_bar, textvariable=self.suffix_var, width=15).pack(side=tk.LEFT, padx=5)

        # --- Tabel Rezultate ---
        self.tree = ttk.Treeview(root, columns=("Sheet", "Tag", "P/N", "BOM", "Found", "Verdict"), show='headings')
        for c in ("Sheet", "Tag", "P/N", "BOM", "Found", "Verdict"): 
            self.tree.heading(c, text=c); self.tree.column(c, width=100, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.prog = ttk.Progressbar(root, mode='determinate'); self.prog.pack(fill=tk.X, padx=10, pady=5)

    def pick_color(self):
        c = colorchooser.askcolor()[0]
        if c: self.highlight_color = (c[0]/255, c[1]/255, c[2]/255)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path: return
        xl = pd.ExcelFile(path)
        sheets = [s for s in xl.sheet_names if "PIPI" in s]
        pop = tk.Toplevel(self.root); lb = tk.Listbox(pop, selectmode="multiple", width=40); [lb.insert(tk.END, s) for s in sheets]; lb.pack(padx=10, pady=10)
        def confirm():
            self.full_results = []
            for i in lb.curselection():
                sn = lb.get(i); df = pd.read_excel(path, sheet_name=sn, header=1)
                for _, r in df.iterrows():
                    tag = str(r.iloc[0]).strip()
                    if tag and tag.lower() != "nan" and "TOTAL" not in tag.upper():
                        self.full_results.append({"sheet": sn, "tag": tag, "pn": str(r.iloc[1]), "target": int(r.iloc[3]) if pd.notnull(r.iloc[3]) else 1, "hits": 0, "verdict": "Pending"})
            self.run_btn.config(state=tk.NORMAL) if self.pdf_path else None; pop.destroy()
        tk.Button(pop, text="OK", command=confirm).pack(pady=5)

    def load_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if self.pdf_path and self.full_results: self.run_btn.config(state=tk.NORMAL)

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        try:
            doc = fitz.open(self.pdf_path)
            excl = set()
            for p in self.exclude_entry.get().split(","):
                try:
                    if "-" in p:
                        s, e = map(int, p.split("-"))
                        for i in range(s, e+1): excl.add(i-1)
                    elif p: excl.add(int(p)-1)
                except: pass

            self.prog["maximum"] = len(self.full_results)
            for i, item in enumerate(self.full_results):
                key = item["tag"] if item["tag"] != "-" else item["pn"]
                count = 0
                if key != "-" and key.lower() != "nan":
                    for p_idx in range(len(doc)):
                        if p_idx in excl: continue
                        page = doc[p_idx]
                        matches = page.search_for(key)
                        if matches:
                            count += len(matches)
                            for rect in matches:
                                try: # SCUTUL PENTRU CODE 4
                                    annot = page.add_highlight_annot(rect)
                                    annot.set_colors(stroke=self.highlight_color); annot.update()
                                except: continue 
                
                item["hits"] = count
                item["verdict"] = "OK" if count == item["target"] else "ERR"
                self.prog["value"] = i + 1; self.root.update_idletasks()

            out_folder = os.path.dirname(self.pdf_path)
            final_name = f"{self.base_name_var.get()}{self.suffix_var.get()}"
            doc.save(os.path.join(out_folder, f"{final_name}.pdf"))
            pd.DataFrame(self.full_results).to_excel(os.path.join(out_folder, f"{final_name}.xlsx"), index=False)
            messagebox.showinfo("Gata", f"Salvat: {final_name}")
            
        except Exception as e: messagebox.showerror("Eroare", str(e))
        finally: self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk(); app = AuditApp(root); root.mainloop()
