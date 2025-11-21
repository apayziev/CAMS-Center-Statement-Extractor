import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
import threading
import sys
import os
import io
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# Import the extraction function directly
try:
    from extract_CamsCenterStatement_transactions import extract_transactions, extract_header, extract_statement_period
    from openpyxl import load_workbook
    import pandas as pd
    extract_available = True
except ImportError:
    extract_available = False


class ExcelScraperGUI:
    def __init__(self, root):
        self.root = root
        root.title("CAMS Center Statement Extractor")
        root.geometry("700x600")
        root.resizable(False, False)

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill="both", expand=True)

        # ========== INPUT SECTION ==========
        ttk.Label(frame, text="Select Excel File", font=("Arial", 14, "bold")).pack(pady=5)

        self.input_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.input_path_var, width=70,
                  bootstyle="info").pack(pady=3)

        ttk.Button(frame, text="Browse Excel", bootstyle="primary-outline",
                   width=20, command=self.browse_file).pack(pady=5)

        ttk.Button(frame, text="Extract Transactions", bootstyle="success",
                   padding=8, width=22, command=self.start_conversion).pack(pady=10)

        self.progress = ttk.Progressbar(frame, mode="indeterminate", bootstyle="info")
        self.progress.pack(fill="x", pady=5)

        # ========== LOGS ==========
        ttk.Label(frame, text="Logs", font=("Arial", 13, "bold")).pack(pady=4)

        # Shrink log height to keep output path visible
        self.log_area = ScrolledText(frame, width=82, height=10, font=("Consolas", 10))
        self.log_area.pack(pady=3)

        # ========== OUTPUT PATH SECTION ==========
        card = ttk.Labelframe(frame, text="Generated Excel Path",
                              padding=12, bootstyle="info")
        card.pack(fill="x", pady=10)

        self.output_var = tk.StringVar()
        ttk.Entry(card, textvariable=self.output_var, width=85,
                  bootstyle="success").pack(pady=3)


    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.input_path_var.set(file_path)

    def log(self, text):
        self.log_area.insert("end", text + "\n")
        self.log_area.see("end")

    def start_conversion(self):
        excel_path = self.input_path_var.get()
        if not excel_path:
            self.log("❌ Please select an Excel file.")
            return
        
        if not extract_available:
            self.log("❌ Required modules not available. Please install dependencies.")
            return

        self.progress.start(10)
        self.log("▶ Starting extraction...\n")

        threading.Thread(target=self.run_script, args=(excel_path,), daemon=True).start()

    def run_script(self, excel_path):
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        
        try:
            self.log(f"Processing file: {os.path.basename(excel_path)}")
            
            # Load workbook to extract header and statement period
            wb = load_workbook(excel_path)
            ws = wb['CamsCenterStatement']
            
            # Extract header and statement period for filename
            header = extract_header(ws)
            statement_period = extract_statement_period(ws)
            
            # Save output in the program's directory (where the script/executable is located)
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                program_dir = os.path.dirname(sys.executable)
            else:
                # Running as script
                program_dir = os.path.dirname(os.path.abspath(__file__))
            
            output_path = os.path.join(program_dir, f"{header}_{statement_period}.xlsx")
            
            self.log(f"Extracting transactions...")
            
            # Temporarily modify FILE_PATH for the extraction
            import extract_CamsCenterStatement_transactions as extractor
            original_path = extractor.FILE_PATH
            extractor.FILE_PATH = excel_path
            
            # Extract transactions
            transactions = extract_transactions()
            
            # Restore original path
            extractor.FILE_PATH = original_path
            
            self.log(f"Creating DataFrame with {len(transactions)} transactions...")
            
            df = pd.DataFrame(transactions, columns=[
                'Date', 'Transaction Date', 'Transaction Time', 'Client Name', 'Ref ID1',
                'Product - Service', 'Qty', 'Unit Price', 'Subtotal', '% Discount',
                'Discount Amount', 'Total', 'Subtotal (Summary)', 'Tax', 'Total Due Center'
            ])
            
            # Save to Excel
            df.to_excel(output_path, index=False)
            
            self.log(f"\n✅ Total transactions extracted: {len(df)}")
            self.log(f"✅ Output saved to: {os.path.basename(output_path)}")
            
            # Update output path display
            self.output_var.set(output_path)
            
            self.progress.stop()
            self.log("\n✅ Extraction completed successfully!")

        except Exception as e:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.progress.stop()
            self.log(f"\n❌ Error: {str(e)}")
            import traceback
            self.log(traceback.format_exc())


# ----- RUN -----
if __name__ == "__main__":
    app = ttk.Window(themename="flatly")
    GUI = ExcelScraperGUI(app)
    app.mainloop()
