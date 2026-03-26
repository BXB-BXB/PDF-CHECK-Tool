import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditAppBB:
    def __init__(self, root):
        self.root = root
        self.root.title("Audit PDF Pro BB - Bogdan Bahrim")
        self.root.geometry("1300x850")
        self.root.configure(bg="#f8f9fa")

        self.full_results = []
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0) 

        # --- PANEL CONTROL ---
        top = tk.Frame(root, bg="#2c3e50", pady=15, padx=20)
        top.pack(fill=tk.X)
        
        tk.Button(top, text="📁 1. Load Excel", command=self.load_excel, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(top, text="📄 2. Load PDF", command=self.load_pdf, width=15).pack(side=tk.LEFT, padx=5)
        
        tk.Label(top, text="Skip Pgs:", fg="white", bg="#2c3e50").pack(side=tk.LEFT, padx=(15,0))
        self.exclude_entry = tk.Entry(top, width=10); self.exclude_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(top, text="🎨 Color", command=self.pick_color).pack(side=tk.LEFT, padx=5)

        self.run_btn = tk.Button(top, text="⚡ START AUDIT", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), width=15)
        self.run_btn.pack(side=tk.RIGHT, padx=5)

        # --- CUSTOM NAMING ---
        name_frame = tk.Frame(root, bg="#dfe6e9", pady=10)
        name_frame.pack(fill=tk.X)
        
        tk.Label(name_frame, text="Nume Fișier Bazal:", bg="#dfe6e9", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=10)
        self.base_name = tk.Entry(name_frame, width=30); self.base_name.insert(0, "Audit_Drawing"); self.base_name.pack(side=tk.LEFT)
        
        tk.Label(name_frame, text=" + Extra/Suffix:", bg="#dfe6e9", font=("Arial", 9)).pack(side=tk.LEFT, padx=10)
        self.suffix_name = tk.Entry(name_frame, width=20); self.suffix_name.insert(0, "_REV01"); self.suffix_name.pack(side=tk.LEFT)

        # --- TABEL ---
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        cols = ("Sheet", "Identifier", "Description", "QTY_BOM", "Found", "Verdict", "Pages")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        cw = {"Sheet": 90, "Identifier": 180, "Description": 300, "QTY_BOM": 80, "Found": 80, "Verdict": 100, "Pages": 150}
        for col, width in cw.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=tk.W if "Desc" in col or "Iden" in col else tk.CENTER, width=width)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y); self.tree.configure(yscrollcommand=vsb.set)

        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=1200, mode='determinate')
        self.progress.pack(pady=10)

    def pick_color(self):
        color = colorchooser.askcolor(title="Select Color")[0]
        if color: self.highlight_color = (color[0]/255, color[1]/255, color[2]/255)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsm *.xlsx")])
        if not path: return
        try:
            xl = pd.ExcelFile(path)
            sheets = [s for s in xl.sheet_names if "PIPI" in s]
            if not sheets:
                messagebox.showwarning("Atentie", "Nu am gasit niciun sheet care sa contina 'PIPI'!")
                return
                
            pop = tk.Toplevel(self.root)
            pop.title("Selecteaza Sheet-uri")
            pop.grab_set() # Face fereastra modala
            lb = tk.Listbox(pop, selectmode="multiple", width=50, height=10)
            for s in sheets: lb.insert(tk.END, s)
            lb.pack(padx=20, pady=10)

            def confirm():
                try:
                    self.full_results = []
                    selection = lb.curselection()
                    if not selection:
                        messagebox.showwarning("Atentie", "Selecteaza cel putin un sheet!")
                        return

                    for i in selection:
                        sn = lb.get(i)
                        # Citim Excel-ul - incercam header=1 (linia 2)
                        df = pd.read_excel(path, sheet_name=sn, header=1)
                        
                        if df.empty: continue

                        for _, row in df.iterrows():
                            # Verificam daca prima coloana are date
                            val = str(row.iloc[0]).strip()
                            if val and val != "nan" and "TOTAL" not in val.upper():
                                desc = str(row.iloc[2]) if len(row) > 2 else "-"
                                qty = int(row.iloc[3]) if len(row) > 3 and pd.notnull(row.iloc[3]) else 1
                                self.full_results.append({
                                    "sheet": sn, "term": val, "desc": desc, "target": qty, 
                                    "hits": 0, "pages": [], "verdict": "Pending"
                                })
                    
                    if self.full_results:
                        self.refresh_table()
                        if self.pdf_path: self.run_btn.config(state=tk.NORMAL)
                        pop.destroy()
                    else:
                        messagebox.showerror("Eroare", "Nu am extras date. Verifica formatul tabelului!")
                except Exception as e:
                    messagebox.showerror("Eroare la procesare", str(e))

            tk.Button(pop, text="OK (Incarca)", command=confirm, width=20, bg="#2ecc71").pack(pady=10)
        except Exception as e:
            messagebox.showerror("Eroare Excel", str(e))

    def load_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if self.pdf_path and self.full_results: self.run_btn.config(state=tk.NORMAL)

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results:
            self.tree.insert("", "end", values=(item["sheet"], item["term"], item["desc"], item["target"], item["hits"], item["verdict"], ", ".join(map(str, sorted(list(set(item["pages"])))))))

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        try:
            doc = fitz.open(self.pdf_path)
            excluded = set()
            raw = self.exclude_entry.get().replace(" ", "")
            if raw:
                for p in raw.split(","):
                    try:
                        if "-" in p:
                            s, e = map(int, p.split("-")); [excluded.add(x-1) for x in range(s, e+1)]
                        else: excluded.add(int(p)-1)
                    except: pass

            self.progress["maximum"] = len(self.full_results)
            for i, item in enumerate(self.full_results):
                count, pgs = 0, []
                for p_idx in range(len(doc)):
                    if p_idx in excluded: continue
                    page = doc[p_idx]
                    m = page.search_for(item["term"])
                    if m:
                        count += len(m); pgs.append(p_idx+1)
                        for r in m:
                            annot = page.add_highlight_annot(r)
                            annot.set_colors(stroke=self.highlight_color)
                            annot.update()
                
                item["hits"], item["pages"] = count, pgs
                item["verdict"] = "✅ MATCH" if count == item["target"] else f"❌ {count}/{item['target']}"
                self.progress["value"] = i+1
                self.root.update_idletasks()
                if i % 3 == 0: self.refresh_table()

            folder = os.path.dirname(self.pdf_path)
            fname = f"{self.base_name.get()}{self.suffix_name.get()}"
            doc.save(os.path.join(folder, f"{fname}.pdf"))
            pd.DataFrame(self.full_results).to_excel(os.path.join(folder, f"{fname}.xlsx"), index=False)
            messagebox.showinfo("Gata", f"Salvat: {fname}")
        except Exception as e:
            messagebox.showerror("Eroare", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk(); AuditAppBB(root); root.mainloop()
