import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor Pro v10 - Full Technical Validator")
        self.root.geometry("1250x850")
        self.root.configure(bg="#f8f9fa")

        self.full_results = []
        self.excel_path = ""
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0) # Galben default (RGB normalizat 0-1)

        # --- UI LAYOUT ---
        top_frame = tk.Frame(root, bg="#2c3e50", pady=15, padx=20)
        top_frame.pack(fill=tk.X)
        
        # Butoane principale
        tk.Button(top_frame, text="📁 1. Load Excel (PIPI)", command=self.load_excel, width=20, bg="#34495e", fg="white").pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="📄 2. Load PDF Drawings", command=self.load_pdf, width=20, bg="#34495e", fg="white").pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="🎨 Color", command=self.pick_color, bg="#34495e", fg="white").pack(side=tk.LEFT, padx=5)
        
        # Filtru Pagini (Restaurat)
        filter_frame = tk.LabelFrame(top_frame, text=" Skip Pages (ex: 1, 3-5) ", fg="white", bg="#2c3e50")
        filter_frame.pack(side=tk.LEFT, padx=15)
        self.exclude_entry = tk.Entry(filter_frame, width=15)
        self.exclude_entry.pack(padx=5, pady=2)

        self.run_btn = tk.Button(top_frame, text="⚡ START AUDIT", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#27ae60", fg="white", width=18, font=("Arial", 10, "bold"))
        self.run_btn.pack(side=tk.RIGHT, padx=5)

        # --- TABEL REZULTATE ---
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        cols = ("Sheet", "Identifier", "QTY_BOM", "Found_PDF", "Verdict", "Pages")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        
        for col in cols:
            self.tree.heading(col, text=col.replace("_", " "))
            self.tree.column(col, anchor=tk.CENTER, width=100)
        
        self.tree.column("Identifier", width=250, anchor=tk.W)
        self.tree.column("Pages", width=250)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)

        # --- PROGRESS & STATUS ---
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=1100, mode='determinate')
        self.progress.pack(pady=10)
        self.status_var = tk.StringVar(value="Waiting for files...")
        tk.Label(root, textvariable=self.status_var, bg="#f8f9fa", font=("Arial", 10, "italic")).pack()

    def pick_color(self):
        color = colorchooser.askcolor(title="Select Highlight Color")
        if color[0]:
            # Convertim din 0-255 in 0-1 pentru PyMuPDF
            self.highlight_color = (color[0][0]/255, color[0][1]/255, color[0][2]/255)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if not path: return
        self.excel_path = path
        
        xl = pd.ExcelFile(path)
        # Filtram foile care contin PIPI (cum ai in macro)
        pipi_sheets = [s for s in xl.sheet_names if "PIPI" in s]
        
        popup = tk.Toplevel(self.root)
        popup.title("Select Multiple Sheets")
        tk.Label(popup, text="Select sheets to verify (Hold Ctrl to select multiple):").pack(pady=10)
        
        lb = tk.Listbox(popup, selectmode="multiple", width=50, height=10)
        for s in pipi_sheets: lb.insert(tk.END, s)
        lb.pack(padx=20, pady=5)

        def confirm():
            self.full_results = []
            for i in lb.curselection():
                sheet_name = lb.get(i)
                df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=1)
                
                # Identificare automata coloane
                id_col = df.columns[0] # SPOOL NO sau VSY P/N
                qty_col = next((c for c in df.columns if "QTY" in str(c).upper()), None)
                
                for _, row in df.iterrows():
                    val = str(row[id_col]).strip()
                    if val and val != "nan" and "TOTAL" not in val.upper():
                        target = int(row[qty_col]) if qty_col and pd.notnull(row[qty_col]) else 1
                        self.full_results.append({
                            "sheet": sheet_name, "term": val, "target": target,
                            "hits": 0, "pages": [], "verdict": "Pending"
                        })
            self.refresh_table()
            if self.pdf_path: self.run_btn.config(state=tk.NORMAL)
            popup.destroy()

        tk.Button(popup, text="Add Selected Sheets", command=confirm).pack(pady=10)

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
                else:
                    excluded.add(int(part) - 1)
        except: pass
        return excluded

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results:
            self.tree.insert("", "end", values=(
                item["sheet"], item["term"], item["target"], 
                item["hits"], item["verdict"], 
                ", ".join(map(str, sorted(list(set(item["pages"])))))
            ))

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        try:
            doc = fitz.open(self.pdf_path)
            excluded = self.parse_exclusions()
            self.progress["maximum"] = len(self.full_results)
            
            for i, item in enumerate(self.full_results):
                self.status_var.set(f"Auditing: {item['term']} ({item['sheet']})")
                count = 0
                found_pages = []
                
                for p_idx in range(len(doc)):
                    if p_idx in excluded: continue
                    page = doc[p_idx]
                    matches = page.search_for(item["term"])
                    
                    if matches:
                        count += len(matches)
                        found_pages.append(p_idx + 1)
                        for rect in matches:
                            annot = page.add_highlight_annot(rect)
                            annot.set_colors(stroke=self.highlight_color)
                            annot.update()
                
                item["hits"] = count
                item["pages"] = found_pages
                # Verdict Logic
                if count == 0: item["verdict"] = "❌ MISSING"
                elif count == item["target"]: item["verdict"] = "✅ MATCH"
                else: item["verdict"] = f"⚠️ DIFF ({count}/{item['target']})"

                self.progress["value"] = i + 1
                if i % 5 == 0: self.refresh_table()
                self.root.update_idletasks()

            # Save Results
            out_pdf = os.path.splitext(self.pdf_path)[0] + "_VALIDATED.pdf"
            doc.save(out_pdf)
            pd.DataFrame(self.full_results).to_excel(out_pdf.replace(".pdf", "_Report.xlsx"), index=False)
            
            self.status_var.set("Audit Complete!")
            messagebox.showinfo("Success", f"Audit finished!\nPDF: {out_pdf}\nReport: Excel created.")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
