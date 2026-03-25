import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor Pro v11 - Advanced Equipment & Component Validator")
        self.root.geometry("1300x850")
        self.root.configure(bg="#f0f2f5")

        self.full_results = []
        self.excel_path = ""
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0) 

        # --- UI LAYOUT ---
        top_frame = tk.Frame(root, bg="#1a2a3a", pady=15, padx=20)
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="📁 Load Excel", command=self.load_excel, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="📄 Load PDF", command=self.load_pdf, width=15).pack(side=tk.LEFT, padx=5)
        
        # Filtru Pagini
        tk.Label(top_frame, text="Skip:", fg="white", bg="#1a2a3a").pack(side=tk.LEFT, padx=(15,0))
        self.exclude_entry = tk.Entry(top_frame, width=8)
        self.exclude_entry.pack(side=tk.LEFT, padx=5)

        self.run_btn = tk.Button(top_frame, text="⚡ RUN AUDIT", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#27ae60", fg="white", font=("Arial", 10, "bold"))
        self.run_btn.pack(side=tk.RIGHT, padx=5)

        # --- TABEL REZULTATE CU COLOANE EXTRA ---
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        # Am adăugat coloanele Description și Part_No
        cols = ("Sheet", "Identifier", "Part_No", "Description", "QTY", "Found", "Verdict")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        
        column_configs = {
            "Sheet": 100, "Identifier": 150, "Part_No": 150, 
            "Description": 250, "QTY": 60, "Found": 60, "Verdict": 120
        }
        
        for col, width in column_configs.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=tk.W if "Desc" in col else tk.CENTER, width=width)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)

        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=1200, mode='determinate')
        self.progress.pack(pady=10)
        self.status_var = tk.StringVar(value="Selectați fișierele pentru a începe.")
        tk.Label(root, textvariable=self.status_var, font=("Arial", 10)).pack()

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if not path: return
        self.excel_path = path
        
        xl = pd.ExcelFile(path)
        pipi_sheets = [s for s in xl.sheet_names if "PIPI" in s]
        
        popup = tk.Toplevel(self.root)
        popup.title("Select BOM Sheets")
        lb = tk.Listbox(popup, selectmode="multiple", width=50, height=10)
        for s in pipi_sheets: lb.insert(tk.END, s)
        lb.pack(padx=20, pady=10)

        def confirm():
            self.full_results = []
            for i in lb.curselection():
                sheet_name = lb.get(i)
                df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=1)
                
                # Mapping coloane (Adaptabil la macro-ul tău)
                cols = df.columns.tolist()
                id_col = cols[0] # Tag sau Spool No
                pn_col = next((c for c in cols if "PART" in str(c).upper() or "P/N" in str(c).upper()), None)
                desc_col = next((c for c in cols if "DESC" in str(c).upper()), None)
                qty_col = next((c for c in cols if "QTY" in str(c).upper()), None)

                for _, row in df.iterrows():
                    val = str(row[id_col]).strip()
                    if val and val != "nan" and "TOTAL" not in val.upper():
                        self.full_results.append({
                            "sheet": sheet_name,
                            "term": val,
                            "part_no": str(row[pn_col]) if pn_col else "-",
                            "desc": str(row[desc_col])[:50] if desc_col else "-", # Limităm lungimea descrierii
                            "target": int(row[qty_col]) if qty_col and pd.notnull(row[qty_col]) else 1,
                            "hits": 0, "pages": [], "verdict": "Pending"
                        })
            self.refresh_table()
            if self.pdf_path: self.run_btn.config(state=tk.NORMAL)
            popup.destroy()

        tk.Button(popup, text="Confirmă Selecția", command=confirm).pack(pady=10)

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results:
            self.tree.insert("", "end", values=(
                item["sheet"], item["term"], item["part_no"], 
                item["desc"], item["target"], item["hits"], item["verdict"]
            ))

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path = path
            if self.full_results: self.run_btn.config(state=tk.NORMAL)

    def parse_exclusions(self):
        excluded = set()
        raw = self.exclude_entry.get().replace(" ", "")
        if not raw: return excluded
        try:
            for part in raw.split(","):
                if "-" in part:
                    s, e = map(int, part.split("-"))
                    for p in range(s, e + 1): excluded.add(p - 1)
                else: excluded.add(int(part) - 1)
        except: pass
        return excluded

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        try:
            doc = fitz.open(self.pdf_path)
            excluded = self.parse_exclusions()
            self.progress["maximum"] = len(self.full_results)
            
            for i, item in enumerate(self.full_results):
                self.status_var.set(f"Căutare: {item['term']} | {item['part_no']}")
                count = 0
                pages = []
                
                # Căutăm în PDF după Identifier (Tag sau Spool)
                for p_idx in range(len(doc)):
                    if p_idx in excluded: continue
                    page = doc[p_idx]
                    matches = page.search_for(item["term"])
                    
                    if matches:
                        count += len(matches)
                        pages.append(p_idx + 1)
                        for rect in matches:
                            annot = page.add_highlight_annot(rect)
                            # Adăugăm un comentariu în PDF cu descrierea piesei
                            annot.set_info(content=f"P/N: {item['part_no']}\nDesc: {item['desc']}")
                            annot.set_colors(stroke=self.highlight_color)
                            annot.update()
                
                item["hits"] = count
                item["pages"] = pages
                item["verdict"] = "✅ MATCH" if count == item["target"] else f"❌ ERR ({count}/{item['target']})"

                self.progress["value"] = i + 1
                if i % 5 == 0: self.refresh_table()
                self.root.update_idletasks()

            out_base = os.path.splitext(self.pdf_path)[0]
            doc.save(out_base + "_AUDITED.pdf")
            pd.DataFrame(self.full_results).to_excel(out_base + "_Full_Report.xlsx", index=False)
            
            messagebox.showinfo("Gata", "Audit Finalizat!\nRaportul conține Tag, P/N și Descriere.")
            
        except Exception as e:
            messagebox.showerror("Eroare", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
