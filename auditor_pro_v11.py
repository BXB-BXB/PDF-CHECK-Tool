import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class Auditor11:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor 11 - BOM Validator | Bogdan Bahrim")
        self.root.geometry("1100x850")
        self.root.configure(bg="#f4f4f4")

        self.data_model = []
        self.pdf_source = ""
        self.highlight_color = (1, 1, 0)

        # --- Interfata Control ---
        top = tk.Frame(root, bg="#2c3e50", pady=15)
        top.pack(fill=tk.X)
        
        btn_cfg = {"width": 15, "bg": "#ecf0f1", "font": ("Arial", 9, "bold")}
        tk.Button(top, text="1. Load Excel", command=self.load_excel, **btn_cfg).pack(side=tk.LEFT, padx=10)
        tk.Button(top, text="2. Load PDF", command=self.load_pdf, **btn_cfg).pack(side=tk.LEFT, padx=10)
        
        tk.Label(top, text="Skip Pgs:", fg="white", bg="#2c3e50").pack(side=tk.LEFT, padx=(20,0))
        self.skip_entry = tk.Entry(top, width=10); self.skip_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(top, text="Color", command=self.pick_color).pack(side=tk.LEFT, padx=10)
        
        self.run_btn = tk.Button(top, text="START AUDIT", command=self.start_thread, bg="#27ae60", fg="white", state=tk.DISABLED, width=15, font=("Arial", 9, "bold"))
        self.run_btn.pack(side=tk.RIGHT, padx=10)

        # --- Sistem Nume Export ---
        export_frame = tk.Frame(root, bg="#dee2e6", pady=8)
        export_frame.pack(fill=tk.X)
        
        tk.Label(export_frame, text="Nume Fișier:", bg="#dee2e6").pack(side=tk.LEFT, padx=10)
        self.base_name = tk.Entry(export_frame, width=25); self.base_name.insert(0, "Audit_Drawing"); self.base_name.pack(side=tk.LEFT)
        
        tk.Label(export_frame, text="+ Suffix:", bg="#dee2e6").pack(side=tk.LEFT, padx=10)
        self.suffix_name = tk.Entry(export_frame, width=15); self.suffix_name.insert(0, "_Rev0"); self.suffix_name.pack(side=tk.LEFT)

        # --- Tabel Rezultate ---
        self.tree = ttk.Treeview(root, columns=("S", "T", "P", "Q", "F", "V"), show='headings')
        for c, h in zip(("S", "T", "P", "Q", "F", "V"), ("Sheet", "Tag", "P/N", "QTY BOM", "Found", "Verdict")):
            self.tree.heading(c, text=h); self.tree.column(c, width=100, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.pbar = ttk.Progressbar(root, mode='determinate'); self.pbar.pack(fill=tk.X, padx=10, pady=5)

    def pick_color(self):
        c = colorchooser.askcolor()[0]
        if c: self.highlight_color = (c[0]/255, c[1]/255, c[2]/255)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path: return
        xl = pd.ExcelFile(path)
        sheets = [s for s in xl.sheet_names if "PIPI" in s]
        
        pop = tk.Toplevel(self.root); lb = tk.Listbox(pop, selectmode="multiple", width=50); [lb.insert(tk.END, s) for s in sheets]; lb.pack(padx=10, pady=10)
        def confirm():
            self.data_model = []
            for i in lb.curselection():
                sn = lb.get(i); df = pd.read_excel(path, sheet_name=sn, header=1)
                for _, r in df.iterrows():
                    tag = str(r.iloc[0]).strip()
                    if tag and tag.lower() != "nan" and "TOTAL" not in tag.upper():
                        self.data_model.append({"sheet": sn, "tag": tag, "pn": str(r.iloc[1]).strip(), "qty": int(r.iloc[3]) if pd.notnull(r.iloc[3]) else 1, "hits": 0, "verdict": "Pending"})
            self.refresh(); self.run_btn.config(state=tk.NORMAL) if self.pdf_source else None; pop.destroy()
        tk.Button(pop, text="Confirm", command=confirm).pack(pady=5)

    def load_pdf(self):
        self.pdf_source = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if self.pdf_source and self.data_model: self.run_btn.config(state=tk.NORMAL)

    def refresh(self):
        [self.tree.delete(i) for i in self.tree.get_children()]
        [self.tree.insert("", "end", values=(d["sheet"], d["tag"], d["pn"], d["qty"], d["hits"], d["verdict"])) for d in self.data_model]

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        try:
            doc = fitz.open(self.pdf_source)
            skips = self.get_skips()
            self.pbar["maximum"] = len(self.data_model)

            for i, item in enumerate(self.data_model):
                # Logica veche: cauta Tag, daca e '-' cauta P/N
                key = item["tag"] if item["tag"] != "-" else item["pn"]
                count = 0
                if key != "-" and key.lower() != "nan":
                    for p_idx in range(len(doc)):
                        if p_idx in skips: continue
                        page = doc[p_idx]
                        matches = page.search_for(key)
                        if matches:
                            count += len(matches)
                            for rect in matches:
                                try: # Protectie Code 4
                                    annot = page.add_highlight_annot(rect)
                                    annot.set_colors(stroke=self.highlight_color); annot.update()
                                except: continue
                
                item["hits"] = count
                item["verdict"] = "✅ OK" if count == item["qty"] else f"❌ {count}/{item['qty']}"
                self.pbar["value"] = i + 1; self.root.update_idletasks()

            out_dir = os.path.dirname(self.pdf_source)
            f_name = f"{self.base_name.get()}{self.suffix_name.get()}"
            doc.save(os.path.join(out_dir, f"{f_name}.pdf"))
            pd.DataFrame(self.data_model).to_excel(os.path.join(out_dir, f"{f_name}.xlsx"), index=False)
            messagebox.showinfo("Gata", f"Audit terminat: {f_name}"); self.refresh()
        except Exception as e: messagebox.showerror("Eroare", str(e))
        finally: self.run_btn.config(state=tk.NORMAL)

    def get_skips(self):
        s = set()
        for p in self.skip_entry.get().split(","):
            try:
                if "-" in p:
                    a, b = map(int, p.split("-"))
                    for i in range(a, b+1): s.add(i-1)
                elif p: s.add(int(p)-1)
            except: pass
        return s

if __name__ == "__main__":
    root = tk.Tk(); Auditor11(root); root.mainloop()
