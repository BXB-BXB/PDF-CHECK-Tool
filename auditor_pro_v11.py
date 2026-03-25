import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor Pro v11 - Stable Fallback & Description")
        self.root.geometry("1350x900")
        self.root.configure(bg="#f4f7f6")

        self.full_results = []
        self.excel_path = ""
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0) # Galben default

        # --- UI SUPERIOR (v10 Style) ---
        top_frame = tk.Frame(root, bg="#1a2a3a", pady=15, padx=20)
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="📁 1. Load Excel", command=self.load_excel, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="📄 2. Load PDF", command=self.load_pdf, width=15).pack(side=tk.LEFT, padx=5)
        
        tk.Label(top_frame, text="Skip Pgs:", fg="white", bg="#1a2a3a").pack(side=tk.LEFT, padx=(15,0))
        self.exclude_entry = tk.Entry(top_frame, width=10)
        self.exclude_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(top_frame, text="🎨 Color", command=self.pick_color).pack(side=tk.LEFT, padx=5)

        self.run_btn = tk.Button(top_frame, text="⚡ START AUDIT", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), width=15)
        self.run_btn.pack(side=tk.RIGHT, padx=5)

        # --- TABEL REZULTATE CU DESCRIPTION ---
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        cols = ("Sheet", "Identifier", "Part_No", "Description", "QTY_BOM", "Found", "Verdict", "Pages")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        
        cw = {"Sheet": 100, "Identifier": 150, "Part_No": 150, "Description": 250, "QTY_BOM": 70, "Found": 70, "Verdict": 100, "Pages": 150}
        for col, width in cw.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=tk.W if "Desc" in col else tk.CENTER, width=width)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)

        # --- EXPORT SETTINGS ---
        export_frame = tk.LabelFrame(root, text=" 💾 Export Names ", bg="#f4f7f6", padx=20, pady=10)
        export_frame.pack(fill=tk.X, padx=20, pady=10)
        self.pdf_name_var = tk.StringVar(value="Audit_Drawing_Result")
        tk.Entry(export_frame, textvariable=self.pdf_name_var, width=40).grid(row=0, column=1, padx=10)
        tk.Label(export_frame, text="Nume PDF:", bg="#f4f7f6").grid(row=0, column=0)
        
        self.xlsx_name_var = tk.StringVar(value="Audit_Final_Report")
        tk.Entry(export_frame, textvariable=self.xlsx_name_var, width=40).grid(row=1, column=1, padx=10)
        tk.Label(export_frame, text="Nume Raport:", bg="#f4f7f6").grid(row=1, column=0)

        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=1200, mode='determinate')
        self.progress.pack(pady=10)
        self.status_var = tk.StringVar(value="Gata.")
        tk.Label(root, textvariable=self.status_var).pack()

    def pick_color(self):
        color = colorchooser.askcolor(title="Culoare")
        if color[0]: self.highlight_color = (color[0][0]/255, color[0][1]/255, color[0][2]/255)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsm *.xlsx")])
        if not path: return
        self.excel_path = path
        xl = pd.ExcelFile(path)
        sheets = [s for s in xl.sheet_names if "PIPI" in s]
        
        popup = tk.Toplevel(self.root)
        lb = tk.Listbox(popup, selectmode="multiple", width=60, height=12)
        for s in sheets: lb.insert(tk.END, s)
        lb.pack(padx=20, pady=10)

        def confirm():
            self.full_results = []
            for i in lb.curselection():
                sheet_name = lb.get(i)
                df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=1)
                cols = df.columns.tolist()
                id_col = cols[0]
                pn_col = next((c for c in cols if "PART" in str(c).upper() or "P/N" in str(c).upper()), None)
                desc_col = next((c for c in cols if "DESC" in str(c).upper()), None)
                qty_col = next((c for c in cols if "QTY" in str(c).upper()), None)

                for _, row in df.iterrows():
                    tag_val = str(row[id_col]).strip()
                    pn_val = str(row[pn_col]).strip() if pn_col else "-"
                    
                    if tag_val and tag_val != "nan" and "TOTAL" not in tag_val.upper():
                        # Logica de selectare a termenului de cautat
                        search_term = pn_val if tag_val == "-" and pn_val != "-" else tag_val
                        
                        self.full_results.append({
                            "sheet": sheet_name,
                            "identifier": tag_val,
                            "term_to_search": search_term, # Asta cautam efectiv in PDF
                            "part_no": pn_val,
                            "desc": str(row[desc_col]) if desc_col else "-",
                            "target": int(row[qty_col]) if qty_col and pd.notnull(row[qty_col]) else 1,
                            "hits": 0, "pages": [], "verdict": "Pending"
                        })
            self.refresh_table()
            if self.pdf_path: self.run_btn.config(state=tk.NORMAL)
            popup.destroy()
        tk.Button(popup, text="Confirm", command=confirm).pack(pady=10)

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path:
            self.pdf_path = path
            if self.full_results: self.run_btn.config(state=tk.NORMAL)

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results:
            self.tree.insert("", "end", values=(item["sheet"], item["identifier"], item["part_no"], 
                                               item["desc"], item["target"], item["hits"], 
                                               item["verdict"], ", ".join(map(str, sorted(list(set(item["pages"])))))))

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        try:
            doc = fitz.open(self.pdf_path)
            excluded = set()
            raw = self.exclude_entry.get().replace(" ", "")
            try:
                for p in raw.split(","):
                    if "-" in p:
                        s, e = map(int, p.split("-"))
                        for i in range(s, e+1): excluded.add(i-1)
                    elif p: excluded.add(int(p)-1)
            except: pass

            self.progress["maximum"] = len(self.full_results)
            for i, item in enumerate(self.full_results):
                count, pgs = 0, []
                # Folosim term_to_search (care e P/N daca Tag-ul e "-")
                term = item["term_to_search"]
                
                if term != "-":
                    for p_idx in range(len(doc)):
                        if p_idx in excluded: continue
                        matches = doc[p_idx].search_for(term)
                        if matches:
                            count += len(matches)
                            pgs.append(p_idx + 1)
                            for rect in matches:
                                annot = doc[p_idx].add_highlight_annot(rect)
                                annot.set_colors(stroke=self.highlight_color)
                                annot.update()
                
                item["hits"], item["pages"] = count, pgs
                item["verdict"] = "✅ MATCH" if count == item["target"] else f"❌ ERR ({count}/{item['target']})"
                self.progress["value"] = i + 1
                if i % 3 == 0: self.refresh_table()
                self.root.update_idletasks()

            folder = os.path.dirname(self.pdf_path)
            doc.save(os.path.join(folder, f"{self.pdf_name_var.get()}.pdf"))
            # Scoatem coloana term_to_search din raportul final pentru curatenie
            df_final = pd.DataFrame(self.full_results).drop(columns=['term_to_search'])
            df_final.to_excel(os.path.join(folder, f"{self.xlsx_name_var.get()}.xlsx"), index=False)
            messagebox.showinfo("Gata", "Audit finalizat cu succes!")
        except Exception as e: messagebox.showerror("Eroare", str(e))
        finally: self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
