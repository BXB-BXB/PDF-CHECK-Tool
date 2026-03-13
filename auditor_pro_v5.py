import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os
import re

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Auditor Pro v5")
        self.root.geometry("1100x750")
        self.root.configure(bg="#f4f4f9")

        # Data & Settings
        self.full_results = []
        self.excel_path = ""
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0) 

        # --- UI Layout ---
        top_frame = tk.Frame(root, bg="#ffffff", pady=15, padx=20, relief=tk.RAISED, borderwidth=1)
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="📁 Load Excel", command=self.load_excel, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="📄 Load PDF", command=self.load_pdf, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="🎨 Color", command=self.pick_color, width=8).pack(side=tk.LEFT, padx=5)
        
        # --- Exclusion Section ---
        exclude_frame = tk.LabelFrame(top_frame, text=" Skip Pages (e.g. 1, 5, 10-15) ", bg="#ffffff", padx=10)
        exclude_frame.pack(side=tk.LEFT, padx=20)
        self.exclude_entry = tk.Entry(exclude_frame, width=20)
        self.exclude_entry.pack(pady=2)

        self.run_btn = tk.Button(top_frame, text="⚡ Run Audit", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#2e7d32", fg="white", width=12, font=('Arial', 9, 'bold'))
        self.run_btn.pack(side=tk.RIGHT, padx=5)

        # --- Table Previewer ---
        self.tree_frame = tk.Frame(root, bg="#f4f4f9")
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=("Term", "Hits", "Status"), show='headings')
        self.tree.heading("Term", text="Excel Search Term")
        self.tree.heading("Hits", text="Total Count")
        self.tree.heading("Status", text="Found on Pages")
        self.tree.column("Term", width=300)
        self.tree.column("Hits", width=100, anchor=tk.CENTER)
        self.tree.column("Status", width=400)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scroller = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        scroller.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scroller.set)

        # --- Progress Bar ---
        self.progress_label = tk.Label(root, text="Ready", bg="#f4f4f9")
        self.progress_label.pack()
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=900, mode='determinate')
        self.progress.pack(pady=10)

    def parse_exclusions(self):
        """Converts user string '1, 3, 5-7' into a set of 0-based page indices {0, 2, 4, 5, 6}"""
        excluded = set()
        raw_text = self.exclude_entry.get().replace(" ", "")
        if not raw_text: return excluded
        
        parts = raw_text.split(",")
        for part in parts:
            if "-" in part:
                try:
                    start, end = map(int, part.split("-"))
                    for p in range(start, end + 1):
                        excluded.add(p - 1)
                except: pass
            else:
                try:
                    excluded.add(int(part) - 1)
                except: pass
        return excluded

    def pick_color(self):
        color = colorchooser.askcolor(title="Select Highlight Color")
        if color[0]:
            self.highlight_color = (color[0][0]/255, color[0][1]/255, color[0][2]/255)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.excel_path = path
            df = pd.read_excel(path)
            self.full_results = []
            col = df.columns[0]
            for val in df[col].dropna():
                self.full_results.append({"term": str(val).strip(), "hits": 0, "status": "Waiting..."})
            self.refresh_table()
            self.check_ready()

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.check_ready()

    def check_ready(self):
        if self.excel_path and self.pdf_path:
            self.run_btn.config(state=tk.NORMAL)

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results[:100]:
            self.tree.insert("", "end", values=(item["term"], item["hits"], item["status"]))

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process_audit, daemon=True).start()

    def process_audit(self):
        try:
            doc = fitz.open(self.pdf_path)
            excluded_pages = self.parse_exclusions()
            total_items = len(self.full_results)
            self.progress["maximum"] = total_items
            
            for i, item in enumerate(self.full_results):
                term = item["term"]
                pages_found = []
                count = 0
                
                for p_num in range(len(doc)):
                    if p_num in excluded_pages:
                        continue # Skip this page
                        
                    page = doc[p_num]
                    hits = page.search_for(term)
                    if hits:
                        count += len(hits)
                        for rect in hits:
                            annot = page.add_highlight_annot(rect)
                            annot.set_colors(stroke=self.highlight_color)
                            annot.update()
                        pages_found.append(str(p_num + 1))
                
                item["hits"] = count
                item["status"] = ", ".join(list(set(pages_found))) if pages_found else "Not Found"

                self.progress["value"] = i + 1
                if i % 10 == 0 or i == total_items - 1:
                    self.refresh_table()
                self.root.update_idletasks()

            # --- Save Logic ---
            base_pdf = os.path.splitext(self.pdf_path)[0]
            out_pdf = f"{base_pdf}_Check.pdf"
            out_xlsx = f"{base_pdf}_Check.xlsx"

            if os.path.exists(out_pdf) or os.path.exists(out_xlsx):
                if not messagebox.askyesno("Overwrite?", "Files ending in '_Check' already exist. Overwrite?"):
                    out_pdf = filedialog.asksaveasfilename(defaultextension=".pdf")
                    out_xlsx = out_pdf.replace(".pdf", ".xlsx")

            doc.save(out_pdf)
            pd.DataFrame(self.full_results).to_excel(out_xlsx, index=False)
            messagebox.showinfo("Success", f"Audit Complete!\nSkipped pages: {list(map(lambda x: x+1, excluded_pages))}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
