import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class Auditor11:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor 11 - BOM Validator | Publisher: Bogdan Bahrim")
        self.root.geometry("1100x800")
        self.root.configure(bg="#eceff1")

        self.data_model = []
        self.pdf_source = ""
        self.color = (1, 1, 0) # Highlight galben

        # --- Zona de Comenzi (Top) ---
        cmd_frame = tk.Frame(root, bg="#263238", pady=15)
        cmd_frame.pack(fill=tk.X)

        btn_style = {"width": 15, "relief": tk.FLAT, "font": ("Segoe UI", 9, "bold")}
        
        tk.Button(cmd_frame, text="📁 1. BAZA EXCEL", command=self.import_excel, **btn_style).pack(side=tk.LEFT, padx=10)
        tk.Button(cmd_frame, text="📄 2. DESEN PDF", command=self.import_pdf, **btn_style).pack(side=tk.LEFT, padx=10)
        
        tk.Label(cmd_frame, text="Sari Pagini:", fg="white", bg="#263238").pack(side=tk.LEFT, padx=(20, 0))
        self.skip_pgs = tk.Entry(cmd_frame, width=10)
        self.skip_pgs.pack(side=tk.LEFT, padx=5)

        self.go_btn = tk.Button(cmd_frame, text="🚀 START AUDIT", command=self.run_engine, 
                                bg="#00c853", fg="white", state=tk.DISABLED, **btn_style)
        self.go_btn.pack(side=tk.RIGHT, padx=10)

        # --- Zona de Nume Export (Custom Suffix) ---
        naming_frame = tk.Frame(root, bg="#cfd8dc", pady=10)
        naming_frame.pack(fill=tk.X)

        tk.Label(naming_frame, text="Nume Fișier:", bg="#cfd8dc").pack(side=tk.LEFT, padx=10)
        self.fn_base = tk.Entry(naming_frame, width=25); self.fn_base.insert(0, "Raport_Audit")
        self.fn_base.pack(side=tk.LEFT, padx=5)

        tk.Label(naming_frame, text="+ Suffix:", bg="#cfd8dc").pack(side=tk.LEFT, padx=10)
        self.fn_suffix = tk.Entry(naming_frame, width=15); self.fn_suffix.insert(0, "_v11")
