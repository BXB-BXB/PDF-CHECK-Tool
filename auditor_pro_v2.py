import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Auditor Pro v2")
        self.root.geometry("900x650")
        self.root.configure(bg="#f4f4f9")

        # Data Storage
        self.full_results = [] # List of dicts for filtering
        self.excel_path = ""
        self.pdf_path = ""
        self.df = None

        # --- Top Control Panel ---
        top_frame = tk.Frame(root, bg="#ffffff", pady=15, padx=20, relief=tk.RAISED, borderwidth=1)
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="📁 Load Excel", command=self.load_excel, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="📄 Load PDF", command=self.load_pdf, width=12).pack(side=tk.LEFT, padx=5)
        
        self.run_btn = tk.Button(top_frame, text="⚡ Run Audit", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#2e7d32", fg="white", width=12, font=('Arial', 9, 'bold'))
        self.run_btn.pack(side=tk.LEFT, padx=20)

        # --- Filter Section ---
        tk.Label(top_frame, text="Filter View:", bg="#ffffff").pack(side=tk.LEFT, padx=5)
        self.filter_var = tk.StringVar(value="Show All")
        self.filter_menu = ttk.Combobox(top_frame, textvariable=self.filter_var, 
                                        values=["Show All", "Found Only", "Not Found Only"], state="readonly", width=15)
        self.filter_menu.pack(side=tk.LEFT, padx=5)
        self.filter_menu.bind("<<ComboboxSelected>>", self.apply_filter)

        # --- Table Previewer ---
        self.tree_frame = tk.Frame(root, bg="#f4f4f9")
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=("Term", "Status"), show='headings')
        self.tree.heading("Term", text="Excel Search Term")
        self.tree.heading("Status", text="Audit Status / Pages")
        self.tree.column("Term", width=350)
        self.tree.column("Status", width=450)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scroller = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        scroller.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scroller.set)

        # --- Bottom Status & Progress ---
        self.progress_label = tk.Label(root, text="Ready", bg="#f4f4f9", font=('Arial', 9))
        self.progress_label.pack()
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=700, mode='determinate')
        self.progress.pack(pady=10)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.excel_path = path
            self.df = pd.read_excel(path)
            self.full_results = []
            col = self.df.columns[0]
            for val in self.df[col].dropna():
                self.full_results.append({"term": str(val).strip(), "status": "Waiting..."})
            self.apply_filter()
            self.progress_label.config(text=f"Loaded Excel: {os.path.basename(path)}")
            self.check_ready()

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.progress_label.config(text=f"Loaded PDF: {os.path.basename(path)}")
            self.check_ready()

    def check_ready(self):
        if self.excel_path and self.pdf_path:
            self.run_btn.config(state=tk.NORMAL)

    def apply_filter(self, event=None):
        # Clear tree
        for i in self.tree.get_children(): self.tree.delete(i)
        
        f = self.filter_var.get()
        for item in self.full_results:
            status = item["status"]
            if f == "Show All":
                self.tree.insert("", "end", values=(item["term"], status))
            elif f == "Found Only" and status not in ["Waiting...", "Not Found"]:
                self.tree.insert("", "end", values=(item["term"], status))
            elif f == "Not Found Only" and status == "Not Found":
                self.tree.insert("", "end", values=(item["term"], status))

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process_audit, daemon=True).start()

    def process_audit(self):
        try:
            doc = fitz.open(self.pdf_path)
            total = len(self.full_results)
            self.progress["maximum"] = total
            
            for i, item in enumerate(self.full_results):
                term = item["term"]
                pages = []
                
                for p_num in range(len(doc)):
                    page = doc[p_num]
                    hits = page.search_for(term)
                    if hits:
                        for rect in hits:
                            page.add_highlight_annot(rect).update()
                        pages.append(str(p_num + 1))
                
                # Update data
                item["status"] = ", ".join(list(set(pages))) if pages else "Not Found"
                
                # Update UI
                self.progress["value"] = i + 1
                self.progress_label.config(text=f"Searching: {i+1} / {total}")
                if i % 5 == 0: # Refresh table every 5 items to keep it smooth
                    self.apply_filter()
                self.root.update_idletasks()

            # Final Save
            out_dir = os.path.dirname(self.pdf_path)
            doc.save(os.path.join(out_dir, "AUDITED_DOCUMENT.pdf"))
            
            # Export updated Excel
            final_df = pd.DataFrame(self.full_results)
            final_df.to_excel(os.path.join(out_dir, "AUDITED_RESULTS.xlsx"), index=False)
            
            self.apply_filter()
            messagebox.showinfo("Finished", "Audit Complete! Checked all items.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
