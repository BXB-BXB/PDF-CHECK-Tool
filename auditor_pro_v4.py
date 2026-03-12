import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Auditor Pro v4")
        self.root.geometry("1000x700")
        self.root.configure(bg="#f4f4f9")

        # Data & Settings
        self.full_results = []
        self.excel_path = ""
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0) # Default Yellow

        # --- UI Layout ---
        top_frame = tk.Frame(root, bg="#ffffff", pady=15, padx=20, relief=tk.RAISED, borderwidth=1)
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="📁 Load Excel", command=self.load_excel, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="📄 Load PDF", command=self.load_pdf, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="🎨 Color", command=self.pick_color, width=8).pack(side=tk.LEFT, padx=5)
        
        self.run_btn = tk.Button(top_frame, text="⚡ Run Audit", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#2e7d32", fg="white", width=12, font=('Arial', 9, 'bold'))
        self.run_btn.pack(side=tk.LEFT, padx=20)

        # --- Table Previewer ---
        self.tree_frame = tk.Frame(root, bg="#f4f4f9")
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        # Added "Hits" column to the table
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
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=800, mode='determinate')
        self.progress.pack(pady=10)

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
        for item in self.full_results[:100]: # Preview limit for speed
            self.tree.insert("", "end", values=(item["term"], item["hits"], item["status"]))

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process_audit, daemon=True).start()

    def process_audit(self):
        try:
            doc = fitz.open(self.pdf_path)
            total_items = len(self.full_results)
            self.progress["maximum"] = total_items
            
            for i, item in enumerate(self.full_results):
                term = item["term"]
                pages = []
                count = 0
                
                for p_num in range(len(doc)):
                    page = doc[p_num]
                    hits = page.search_for(term)
                    if hits:
                        count += len(hits) # Count every single occurrence
                        for rect in hits:
                            annot = page.add_highlight_annot(rect)
                            annot.set_colors(stroke=self.highlight_color)
                            annot.update()
                        pages.append(str(p_num + 1))
                
                item["hits"] = count
                item["status"] = ", ".join(list(set(pages))) if pages else "Not Found"

                # Update UI periodically
                self.progress["value"] = i + 1
                if i % 10 == 0 or i == total_items - 1:
                    self.refresh_table()
                self.root.update_idletasks()

            # --- Save Logic (Check & Ask) ---
            base_pdf = os.path.splitext(self.pdf_path)[0]
            out_pdf = f"{base_pdf}_Check.pdf"
            out_xlsx = f"{base_pdf}_Check.xlsx"

            if os.path.exists(out_pdf) or os.path.exists(out_xlsx):
                if not messagebox.askyesno("Overwrite?", f"Files ending in '_Check' already exist. Overwrite?"):
                    out_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="Audit_Manual_Save.pdf")
                    out_xlsx = out_pdf.replace(".pdf", ".xlsx")

            doc.save(out_pdf)
            
            # Export Final Excel
            final_df = pd.DataFrame(self.full_results)
            final_df.to_excel(out_xlsx, index=False)

            messagebox.showinfo("Success", f"Audit Complete!\nResults saved to '_Check' files.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
