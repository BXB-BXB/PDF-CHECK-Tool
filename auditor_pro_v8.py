import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor Pro v9 - Multi-Sheet Technical Validator")
        self.root.geometry("1200x850")
        
        self.full_results = []
        self.excel_path = ""
        self.pdf_path = ""
        self.selected_sheets = []

        # --- UI Setup ---
        top_frame = tk.Frame(root, bg="#2c3e50", pady=20)
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="1. Încarcă Excel (BOM)", command=self.load_excel, width=25).pack(side=tk.LEFT, padx=10)
        tk.Button(top_frame, text="2. Încarcă PDF (Drawings)", command=self.load_pdf, width=25).pack(side=tk.LEFT, padx=10)
        
        self.run_btn = tk.Button(top_frame, text="⚡ Start Audit", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#27ae60", fg="white", width=20, font=("Arial", 10, "bold"))
        self.run_btn.pack(side=tk.RIGHT, padx=20)

        # --- Tabel Rezultate ---
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        cols = ("Sheet", "Identifier", "Target_QTY", "Found_PDF", "Verdict", "Pages")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        
        for col in cols:
            self.tree.heading(col, text=col.replace("_", " "))
            self.tree.column(col, anchor=tk.CENTER, width=120)
        
        self.tree.column("Identifier", width=250, anchor=tk.W)
        self.tree.column("Pages", width=200)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=1100, mode='determinate')
        self.progress.pack(pady=10)
        
        self.status_var = tk.StringVar(value="Gata pentru încărcare.")
        tk.Label(root, textvariable=self.status_var).pack()

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if not path: return
        self.excel_path = path
        
        xl = pd.ExcelFile(path)
        sheets = [s for s in xl.sheet_names if "PIPI" in s]
        
        # Fereastră de selecție multiplă
        popup = tk.Toplevel(self.root)
        popup.title("Selectează Foile pentru Audit")
        tk.Label(popup, text="Selectează listele pe care vrei să le verifici simultan:").pack(padx=20, pady=10)
        
        lb = tk.Listbox(popup, selectmode="multiple", width=50, height=10)
        for s in sheets: lb.insert(tk.END, s)
        lb.pack(padx=20, pady=5)

        def confirm_selection():
            selected_indices = lb.curselection()
            self.full_results = []
            
            for i in selected_indices:
                sheet_name = lb.get(i)
                # Citim foaia (header=1 pentru a sări peste titlul pus de macro)
                df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=1)
                
                # Detectăm automat coloanele
                id_col = df.columns[0] # De obicei Spool No sau VSY P/N
                qty_col = next((c for c in df.columns if "QTY" in str(c).upper()), None)
                
                for _, row in df.iterrows():
                    term = str(row[id_col]).strip()
                    if term and term != "nan" and "TOTAL" not in term.upper():
                        self.full_results.append({
                            "sheet": sheet_name,
                            "term": term,
                            "target": int(row[qty_col]) if qty_col and pd.notnull(row[qty_col]) else 1,
                            "hits": 0,
                            "pages": [],
                            "verdict": "Pending"
                        })
            
            self.refresh_table()
            if self.pdf_path: self.run_btn.config(state=tk.NORMAL)
            popup.destroy()

        tk.Button(popup, text="Încarcă Foile Selectate", command=confirm_selection).pack(pady=10)

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results:
            self.tree.insert("", "end", values=(
                item["sheet"],
                item["term"], 
                item["target"], 
                item["hits"], 
                item["verdict"],
                ", ".join(map(str, sorted(list(set(item["pages"])))))
            ))

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path = path
            if self.full_results: self.run_btn.config(state=tk.NORMAL)

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process_audit, daemon=True).start()

    def process_audit(self):
        try:
            doc = fitz.open(self.pdf_path)
            self.progress["maximum"] = len(self.full_results)
            
            for i, item in enumerate(self.full_results):
                self.status_var.set(f"Căutare {item['term']} din {item['sheet']}...")
                count = 0
                pages = []
                
                for p_num in range(len(doc)):
                    page = doc[p_num]
                    matches = page.search_for(item["term"])
                    if matches:
                        count += len(matches)
                        pages.append(p_num + 1)
                        # Culori diferite în funcție de foaie
                        color = (0, 0.5, 1) if "SP_LIST" in item["sheet"] else (1, 0.5, 0)
                        for rect in matches:
                            annot = page.add_highlight_annot(rect)
                            annot.set_colors(stroke=color)
                            annot.update()
                
                item["hits"] = count
                item["pages"] = pages
                
                # Logică Verdict
                if count == 0: item["verdict"] = "❌ Missing"
                elif count == item["target"]: item["verdict"] = "✅ Match"
                else: item["verdict"] = f"⚠️ {count}/{item['target']}"

                self.progress["value"] = i + 1
                if i % 5 == 0: self.refresh_table()
                self.root.update_idletasks()

            output_name = os.path.splitext(self.pdf_path)[0] + "_AUDITED_FULL.pdf"
            doc.save(output_name)
            messagebox.showinfo("Succes", f"Audit complet!\nRezultat salvat în: {output_name}")
            
        except Exception as e:
            messagebox.showerror("Eroare", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
