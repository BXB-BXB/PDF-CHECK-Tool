import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Auditor Pro v3")
        self.root.geometry("950x700")
        self.root.configure(bg="#f4f4f9")

        # Data & Settings
        self.full_results = []
        self.excel_path = ""
        self.pdf_path = ""
        self.df = None
        self.highlight_color = (1, 1, 0) # Default Yellow (RGB 0-1 scale)

        # --- Top Control Panel ---
        top_frame = tk.Frame(root, bg="#ffffff", pady=15, padx=20, relief=tk.RAISED, borderwidth=1)
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="📁 Load Excel", command=self.load_excel, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="📄 Load PDF", command=self.load_pdf, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="🎨 Color", command=self.pick_color, width=8).pack(side=tk.LEFT, padx=5)
        
        self.run_btn = tk.Button(top_frame, text="⚡ Run Audit", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#2e7d32", fg="white", width=12, font=('Arial', 9, 'bold'))
        self.run_btn.pack(side=tk.LEFT, padx=20)

        # Filter View
        tk.Label(top_frame, text="View:", bg="#ffffff").pack(side=tk.LEFT, padx=5)
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

        # --- Progress Bar ---
        self.progress_label = tk.Label(root, text="Ready", bg="#f4f4f9")
        self.progress_label.pack()
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=700, mode='determinate')
        self.progress.pack(pady=10)

    def pick_color(self):
        color = colorchooser.askcolor(title="Select Highlight Color")
        if color[0]: # RGB tuple from 0-255
            self.highlight_color = (color[0][0]/255, color[0][1]/255, color[0][2]/255)

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
            self.check_ready()

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.check_ready()

    def check_ready(self):
        if self.excel_path and self.pdf_path:
            self.run_btn.config(state=tk.NORMAL)

    def apply_filter(self, event=None):
        for i in self.tree.get_children(): self.tree.delete(i)
        f = self.filter_var.get()
        for item in self.full_results:
            status = item["status"]
            if f == "Show All" or (f == "Found Only" and status not in ["Waiting...", "Not Found"]) or (f == "Not Found Only" and status == "Not Found"):
                self.tree.insert("", "end", values=(item["term"], status))

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process_audit, daemon=True).start()

    def process_audit(self):
        try:
            doc = fitz.open(self.pdf_path)
            found_count = 0
            
            for i, item in enumerate(self.full_results):
                term = item["term"]
                pages = []
                for p_num in range(len(doc)):
                    page = doc[p_num]
                    hits = page.search_for(term)
                    if hits:
                        for rect in hits:
                            annot = page.add_highlight_annot(rect)
                            annot.set_colors(stroke=self.highlight_color)
                            annot.update()
                        pages.append(str(p_num + 1))
                
                if pages:
                    item["status"] = ", ".join(list(set(pages)))
                    found_count += 1
                else:
                    item["status"] = "Not Found"

                self.progress["value"] = i + 1
                self.root.update_idletasks()

            # --- Results Summary ---
            total = len(self.full_results)
            missing_count = total - found_count
            
            # --- Save Logic (Check & Ask) ---
            base_pdf = os.path.splitext(self.pdf_path)[0]
            out_pdf = f"{base_pdf}_Check.pdf"
            out_xlsx = f"{base_pdf}_Check.xlsx"

            if os.path.exists(out_pdf) or os.path.exists(out_xlsx):
                if not messagebox.askyesno("Overwrite?", f"Files ending in '_Check' already exist. Overwrite?"):
                    out_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="Audit_Manual_Save.pdf")
                    out_xlsx = out_pdf.replace(".pdf", ".xlsx")

            doc.save(out_pdf)
            
            # Export Excel with Summary
            final_df = pd.DataFrame(self.full_results)
            summary_df = pd.DataFrame([{"term": "SUMMARY:", "status": f"Found: {found_count} | Missing: {missing_count}"}])
            final_df = pd.concat([summary_df, final_df], ignore_index=True)
            final_df.to_excel(out_xlsx, index=False)

            messagebox.showinfo("Success", f"Done!\nFound: {found_count}\nMissing: {missing_count}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
