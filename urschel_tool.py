import fitz  # PyMuPDF
import pandas as pd
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import datetime

# File to store the session data
CONFIG_FILE = "urschel_tool_config.json"

class UrschelPDFTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Urschel PDF Generator")
        self.root.state('zoomed')
        self.root.configure(bg="#F4F6F8")

        # Nagarkot Brand Colors
        self.c_primary = "#1F3F6E"
        self.c_bg = "#F4F6F8"
        self.c_white = "#FFFFFF"
        self.c_text = "#1E1E1E"
        self.c_muted = "#6B7280"
        self.c_border = "#E5E7EB"

        # Variables
        self.excel_path = tk.StringVar()
        self.pdf_paths = []  # List to store multiple PDF paths
        self.num_models = tk.StringVar(value="10")

        # Mapping for priority search
        self.mapping = {
            "1700": ["66", "67", "10", "11", "64", "62"],
            "CC": ["22", "23", "10", "37", "11", "34"],
            "DCA": ["43", "42", "10", "11"],
            "ETRS": ["52", "55", "10", "11"],
            "OV": ["27", "29", "10", "11"],
            "RA": ["10", "18", "19"]
        }

        self.create_widgets()
        self.load_last_session()

    def create_widgets(self):
        # 1. HEADER (Dynamic Height)
        header_frame = tk.Frame(self.root, bg=self.c_white, height=60)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False) # Keep height

        # Logo on the left inside header
        try:
            import sys
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            logo_path = os.path.join(base_path, "logo.png")

            # We must set logo height to exactly 20 units (pixels) proportionally
            img = Image.open(logo_path)
            ratio = 20 / img.height
            new_width = int(img.width * ratio)
            img = img.resize((new_width, 20), Image.Resampling.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)
            
            logo_label = tk.Label(header_frame, image=self.logo_img, bg=self.c_white)
            logo_label.pack(side=tk.LEFT, padx=30, pady=20)
        except Exception as e:
            print("Logo load error:", e)
            # Fallback if no logo
            tk.Label(header_frame, text="[LOGO]", bg=self.c_white, fg=self.c_primary, font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=30, pady=20)

        # Tool name strictly absolute centered
        title_label = tk.Label(header_frame, text="Urschel PDF Generator", bg=self.c_white, fg=self.c_primary, font=("Arial", 18, "bold"))
        title_label.place(relx=0.5, rely=0.3, anchor="center")
        
        subtitle_label = tk.Label(header_frame, text="Automated Catalogue Extraction", bg=self.c_white, fg=self.c_muted, font=("Arial", 10))
        subtitle_label.place(relx=0.5, rely=0.7, anchor="center")

        # Border under header
        tk.Frame(self.root, bg=self.c_border, height=1).pack(fill=tk.X, side=tk.TOP)

        # 2. BODY
        body_frame = tk.Frame(self.root, bg=self.c_bg)
        body_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=50)

        # Center Container
        center_container = tk.Frame(body_frame, bg=self.c_white, padx=40, pady=40, relief=tk.FLAT)
        center_container.pack(expand=True, fill=tk.Y, pady=20)
        
        # Style configurations
        lbl_font = ("Arial", 10, "bold")
        entry_font = ("Arial", 11)
        btn_font = ("Arial", 10, "bold")

        # Step 1
        tk.Label(center_container, text="Step 1: Select Urschel Excel File", font=lbl_font, bg=self.c_white, fg=self.c_text).pack(anchor="w", pady=(0, 5))
        excel_frame = tk.Frame(center_container, bg=self.c_white)
        excel_frame.pack(fill=tk.X, pady=(0, 20))
        tk.Entry(excel_frame, textvariable=self.excel_path, width=40, font=entry_font, relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=(0, 10), ipady=4)
        tk.Button(excel_frame, text="Browse Excel...", command=self.browse_excel, bg=self.c_white, fg=self.c_primary, font=btn_font, relief=tk.SOLID, bd=1, cursor="hand2").pack(side=tk.LEFT, ipadx=10, ipady=2)

        # Step 2
        tk.Label(center_container, text="Step 2: Select Catalogue PDF Files", font=lbl_font, bg=self.c_white, fg=self.c_text).pack(anchor="w", pady=(0, 5))
        pdf_frame = tk.Frame(center_container, bg=self.c_white)
        pdf_frame.pack(fill=tk.X, pady=(0, 20))
        tk.Button(pdf_frame, text="Select PDF Catalogues...", command=self.browse_pdfs, bg=self.c_white, fg=self.c_primary, font=btn_font, relief=tk.SOLID, bd=1, cursor="hand2").pack(side=tk.LEFT, ipadx=10, ipady=2)
        self.files_label = tk.Label(pdf_frame, text="No files selected", bg=self.c_white, fg=self.c_muted, font=("Arial", 10))
        self.files_label.pack(side=tk.LEFT, padx=15)

        # Step 3
        tk.Label(center_container, text="Step 3: Number of models to fetch", font=lbl_font, bg=self.c_white, fg=self.c_text).pack(anchor="w", pady=(0, 5))
        tk.Entry(center_container, textvariable=self.num_models, width=15, font=entry_font, relief=tk.SOLID, bd=1).pack(anchor="w", ipady=4, pady=(0, 30))

        # Action Area
        tk.Button(center_container, text="GENERATE PDF", bg=self.c_primary, fg=self.c_white, 
                  activebackground="#2A528F", activeforeground="white",
                  font=("Arial", 12, "bold"), command=self.run_process, cursor="hand2", relief=tk.FLAT).pack(fill=tk.X, ipady=10)

        # Inline Status Area
        self.status_label = tk.Label(center_container, text="", bg=self.c_white, fg=self.c_primary, font=("Arial", 10, "italic"))
        self.status_label.pack(pady=(15, 0))
        self.progress_label = tk.Label(center_container, text="", bg=self.c_white, fg=self.c_text, font=("Arial", 10, "bold"))
        self.progress_label.pack()

        # Logs Area
        log_frame = tk.Frame(center_container, bg=self.c_white)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 9, "bold"), background=self.c_bg)
        style.configure("Treeview", font=("Arial", 9))
        style.map("Treeview", background=[("selected", self.c_primary)])
        
        columns = ("Model", "Status", "Source PDF", "Page(s)")
        self.log_table = ttk.Treeview(log_frame, columns=columns, show="headings", height=8)
        self.log_table.heading("Model", text="Model")
        self.log_table.heading("Status", text="Status")
        self.log_table.heading("Source PDF", text="Source PDF")
        self.log_table.heading("Page(s)", text="Page(s)")
        
        self.log_table.column("Model", width=120, anchor=tk.W)
        self.log_table.column("Status", width=80, anchor=tk.CENTER)
        self.log_table.column("Source PDF", width=250, anchor=tk.W)
        self.log_table.column("Page(s)", width=80, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_table.yview)
        self.log_table.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 3. FOOTER
        footer_frame = tk.Frame(self.root, bg=self.c_bg)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(footer_frame, text="Nagarkot Forwarders Pvt. Ltd. ©", font=("Arial", 8), bg=self.c_bg, fg=self.c_muted).pack(side=tk.LEFT, padx=20, pady=10)

    def save_session(self):
        """Saves current file selections to a local JSON file."""
        data = {
            "excel_path": self.excel_path.get(),
            "pdf_paths": self.pdf_paths
        }
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(data, f)
        except:
            pass

    def load_last_session(self):
        """Loads files from the previous session if they exist."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    data = json.load(f)
                    self.excel_path.set(data.get("excel_path", ""))
                    saved_paths = data.get("pdf_paths", [])
                    self.pdf_paths = [f for f in saved_paths if os.path.exists(f)]
                    
                    if self.pdf_paths:
                        self.files_label.config(text=f"{len(self.pdf_paths)} files auto-loaded", fg="blue")
            except:
                pass

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path: self.excel_path.set(path)

    def browse_pdfs(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.pdf_paths = list(paths)
            self.files_label.config(text=f"{len(self.pdf_paths)} files selected", fg="black")
            self.save_session()

    def run_process(self):
        if not self.excel_path.get() or not self.pdf_paths:
            messagebox.showerror("Error", "Please select the Excel file and at least one PDF.")
            return

        try:
            target_count = int(self.num_models.get())
            df = pd.read_excel(self.excel_path.get())
            # Clean "Model number" column
            model_list = df['Model'].dropna().astype(str).str.strip().tolist()
            
            output_pdf = fitz.open()
            found_count = 0

            # --- Progress UI ---
            self.status_label.config(text="Preparing PDFs...")
            self.progress_label.config(text=f"Found: 0 / {target_count} models")
            
            # Clear previous logs
            for item in self.log_table.get_children():
                self.log_table.delete(item)
                
            self.root.update()
            # -------------------

            # Pre-open documents and extract text to avoid repeated parsing overhead
            opened_docs = {}
            doc_texts = {}
            for path in self.pdf_paths:
                if os.path.exists(path):
                    try:
                        self.status_label.config(text=f"Reading catalogue: {os.path.basename(path)[:30]}...")
                        self.root.update()
                        
                        doc = fitz.open(path)
                        opened_docs[path] = doc
                        # Extract and lower all text on every page into memory for instant string matching
                        doc_texts[path] = [page.get_text("text").lower() for page in doc]
                    except Exception as e:
                        print(f"Skipping {path}: {e}")

            self.status_label.config(text="Extracting models...")
            self.root.update()

            for model in model_list:
                if found_count >= target_count: break
                
                # Priority Mapping Logic
                priority_pdfs = []
                for key, prefixes in self.mapping.items():
                    if any(model.startswith(pre) for pre in prefixes):
                        priority_pdfs.extend([f for f in self.pdf_paths if key.lower() in os.path.basename(f).lower()])
                
                priority_pdfs = list(dict.fromkeys(priority_pdfs))
                search_queue = priority_pdfs + [f for f in self.pdf_paths if f not in priority_pdfs]
                
                model_found = False
                model_lower = model.lower()

                for pdf_full_path in search_queue:
                    if model_found: break
                    if pdf_full_path not in opened_docs: continue
                    
                    doc = opened_docs[pdf_full_path]
                    texts = doc_texts[pdf_full_path]
                    
                    for page_num in range(len(doc)):
                        # Quick string check before doing expensive PDF vector extraction
                        if model_lower in texts[page_num]:
                            page = doc[page_num]
                            text_rects = page.search_for(model)
                            
                            if text_rects:
                                # Dual-page logic (Current and Previous)
                                start_page = max(0, page_num - 1)
                                output_pdf.insert_pdf(doc, from_page=start_page, to_page=page_num)
                                
                                model_found = True
                                found_count += 1
                                
                                self.progress_label.config(text=f"Found: {found_count} / {target_count} models")
                                self.log_table.insert("", tk.END, values=(model, "Found", os.path.basename(doc.name), f"{start_page + 1}-{page_num + 1}"))
                                self.root.update()
                                break
                                
                if not model_found:
                    self.log_table.insert("", tk.END, values=(model, "Not Found", "-", "-"))
                    self.root.update()

            # Close all pre-opened documents
            for doc in opened_docs.values():
                doc.close()

            self.status_label.config(text="Finished.")
            self.root.update()

            # Save Output
            if len(output_pdf) > 0:
                # Save next to the Excel file
                save_dir = os.path.dirname(self.excel_path.get())
                excel_name = os.path.splitext(os.path.basename(self.excel_path.get()))[0]
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
                filename = f"{excel_name}_Catalogue_{timestamp}.pdf"
                save_path = os.path.join(save_dir, filename)
                output_pdf.save(save_path)
                output_pdf.close()
                self.status_label.config(text=f"Finished. Saved at:\n{save_path}")
                messagebox.showinfo("Success", f"Found {found_count} models.\nSaved to: {save_path}")
            else:
                messagebox.showwarning("Not Found", "No matching models were found.")

        except Exception as e:
            self.status_label.config(text="Error occurred.")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UrschelPDFTool(root)
    root.mainloop()