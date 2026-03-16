import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor Pro v7 - Verificare Etichetare & QTY")
        self.root.geometry("1200x800")
        self.root.configure(bg="#f4f4f9")

        self.full_results = []
        self.excel_path = ""
        self.pdf_path = ""
        self.df_source = None
        self.highlight_color = (1, 1, 0) 

        # --- UI Layout ---
        top_frame = tk.Frame(root, bg="#ffffff", pady=15, padx=20, relief=tk.RAISED, borderwidth=1)
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="📁 1. Load PIPI Excel", command=self.load_excel_macro, width=18, bg="#e3f2fd").pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="📄 2. Load Drawing PDF", command=self.load_pdf, width=18).pack(side=tk.LEFT, padx=5)
        
        exclude_frame = tk.LabelFrame(top_frame, text=" Skip Pages ", bg="#ffffff")
        exclude_frame.pack(side=tk.LEFT, padx=20)
        self.exclude_entry = tk.Entry(exclude_frame, width=12)
        self.exclude_entry.pack(padx=5, pady=2)

        self.run_btn = tk.Button(top_frame, text="⚡ Run Validation", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#2e7d32", fg="white", width=18, font=('Arial', 10, 'bold'))
        self.run_btn.pack(side=tk.RIGHT, padx=5)

        # --- Table Previewer ---
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        # Am adăugat coloana "Target QTY" și "Verdict"
        cols = ("Term", "Target_QTY", "Hits", "Status", "Verdict")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        
        self.tree.heading("Term", text="P/N / Spool Name")
        self.tree.heading("Target_QTY", text="QTY (Excel)")
        self.tree.heading("Hits", text="Found (PDF)")
        self.tree.heading("Status", text="Pages")
        self.tree.heading("Verdict", text="Match Status")
        
        self.tree.column("Target_QTY", width=100, anchor=tk.CENTER)
        self.tree.column("Hits", width=100, anchor=tk.CENTER)
        self.tree.column("Verdict", width=150, anchor=tk.CENTER)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- Progress Bar ---
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=1000, mode='determinate')
        self.progress.pack(pady=10)

    def load_excel_macro(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if not path: return
        self.excel_path = path
        
        xl = pd.ExcelFile(path)
        sheets = [s for s in xl.sheet_names if "PIPI" in s]
        
        popup = tk.Toplevel(self.root)
        popup.title("Select BOM Sheet")
        tk.Label(popup, text="Alege lista pentru verificare:").pack(padx=20, pady=10)
        combo = ttk.Combobox(popup, values=sheets, state="readonly", width=30)
        combo.pack(padx=20, pady=5)
        combo.current(0)

        def on_select():
            sheet_name = combo.get()
            # Citim datele ignorând titlul de pe rândul 1
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=1)
            
            self.full_results = []
            # Identificăm coloana de QTY (Macro-ul tău o numește de obicei 'QTY')
            qty_col = next((c for c in df.columns if "QTY" in str(c).upper()), None)
            search_col = df.columns[0]
            
            for _, row in df.iterrows():
                term = str(row[search_col]).strip()
                if term and term != "nan":
                    target_qty = int(row[qty_col]) if qty_col and pd.notnull(row[qty_col]) else 1
                    self.full_results.append({
                        "term": term,
                        "target_qty": target_qty,
                        "hits": 0,
                        "status": "Waiting...",
                        "verdict": "Pending"
                    })
            
            self.refresh_table()
            self.check_ready()
            popup.destroy()

        tk.Button(popup, text="Load Data", command=on_select).pack(pady=10)

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results:
            self.tree.insert("", "end", values=(
                item["term"], 
                item["target_qty"], 
                item["hits"], 
                item["status"], 
                item["verdict"]
            ))

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.check_ready()

    def check_ready(self):
        if self.excel_path and self.pdf_path:
            self.run_btn.config(state=tk.NORMAL)

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process_audit, daemon=True).start()

    def process_audit(self):
        try:
            doc = fitz.open(self.pdf_path)
            excluded = self.parse_exclusions()
            total = len(self.full_results)
            self.progress["maximum"] = total
            
            for i, item in enumerate(self.full_results):
                term = item["term"]
                pages_found = []
                count = 0
                
                for p_num in range(len(doc)):
                    if p_num in excluded: continue
                    page = doc[p_num]
                    hits = page.search_for(term)
                    if hits:
                        count += len(hits)
                        for rect in hits:
                            annot = page.add_highlight_annot(rect)
                            annot.set_colors(stroke=(1, 0, 0) if count != item["target_qty"] else (0, 0.8, 0))
                            annot.update()
                        pages_found.append(str(p_num + 1))
                
                item["hits"] = count
                item["status"] = ", ".join(list(set(pages_found))) if pages_found else "MISSING"
                
                # --- LOGICA DE VERDICT ---
                if count == 0:
                    item["verdict"] = "❌ NOT LABELED"
                elif count == item["target_qty"]:
                    item["verdict"] = "✅ MATCH"
                elif count < item["target_qty"]:
                    item["verdict"] = f"⚠️ UNDER ({count}/{item['target_qty']})"
                else:
                    item["verdict"] = f"❗ OVER ({count}/{item['target_qty']})"

                self.progress["value"] = i + 1
                if i % 5 == 0: self.refresh_table()
                self.root.update_idletasks()

            # Salvare Raport Detaliat
            base = os.path.splitext(self.pdf_path)[0]
            doc.save(f"{base}_Validation_Drawing.pdf")
            pd.DataFrame(self.full_results).to_excel(f"{base}_Validation_Report.xlsx", index=False)
            
            messagebox.showinfo("Done", "Validare finalizată!\nVerifică coloana 'Verdict' pentru neconcordanțe.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

    def parse_exclusions(self):
        # ... (aceeași funcție ca în v6)
        return set()

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
