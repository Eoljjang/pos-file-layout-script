import os
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Tkinter modules for the UI
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ==========================================
# CONFIGURATION LOADER ENGINE
# ==========================================
def load_layout_configs():
    """Loads parsing schemas from local external JSON configs."""
    configs = {}
    files = {
        "B01": "configs/b01_config.json",
        "B02": "configs/b02_config.json",
        "P01": "configs/p01_config.json"
    }
    
    for key, filename in files.items():
        if not os.path.exists(filename):
            raise FileNotFoundError(
                f"Missing critical layout specification schema: '{filename}'.\n"
                f"Ensure it resides in the application execution path."
            )
        with open(filename, "r", encoding="utf-8") as f:
            configs[key] = json.load(f)
            
    return configs["B01"], configs["B02"], configs["P01"]

try:
    file_layout_dict_b01, file_layout_dict_b02, file_layout_dict_p01 = load_layout_configs()
except Exception as err:
    # Fail-fast safeguard window prior to standard Tk initialization loop execution context
    root_err = tk.Tk()
    root_err.withdraw()
    messagebox.showerror("Initialization Error", str(err))
    exit(1)


# ==========================================
# PARSING ENGINE LOGIC
# ==========================================
def retrieve_specific_entry(line, layout_dict, entry_name):
    if entry_name not in layout_dict:
        raise ValueError(f"Entry name '{entry_name}' not found in layout dictionary.")
    offset = layout_dict[entry_name]["offset"]
    length = layout_dict[entry_name]["length"]
    val = line[offset:offset + length].strip()
    
    # Quick data conversion for clean spreadsheets
    if val.replace('.', '', 1).isdigit():
        return float(val) if '.' in val else int(val)
    return val

def retrieve_all_entries(line, layout_dict):
    return {entry_name: retrieve_specific_entry(line, layout_dict, entry_name) for entry_name in layout_dict}

def parse_bk_file(filepath):
    products = []
    current_product = {}
    
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.rstrip('\n')
            if not line:
                continue
            
            record_type = line[:3]
            if record_type == "B01":
                offset = file_layout_dict_b01["product_description"]["offset"]
                length = file_layout_dict_b01["product_description"]["length"]
                product_name = line[offset:offset + length].strip()

                current_product = {
                    product_name: {
                        "B01": line, "B02": None, "P01": None
                    }
                }
            elif record_type == "B02" and current_product:
                name = list(current_product.keys())[0]
                current_product[name]["B02"] = line
            elif record_type == "P01" and current_product:
                name = list(current_product.keys())[0]
                current_product[name]["P01"] = line
                products.append(current_product)
                current_product = {}

    for product in products:
        name = list(product.keys())[0]
        # Use default empty strings if strings are unexpectedly missing from stream block
        product[name]["B01_entries"] = retrieve_all_entries(product[name]["B01"] or " "*210, file_layout_dict_b01)
        product[name]["B02_entries"] = retrieve_all_entries(product[name]["B02"] or " "*210, file_layout_dict_b02)
        product[name]["P01_entries"] = retrieve_all_entries(product[name]["P01"] or " "*240, file_layout_dict_p01)
        
    return products

# ==========================================
# EXCEL EXPORT ENGINE
# ==========================================
def export_to_excel(products, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Product Catalog Master"
    ws.views.sheetView[0].showGridLines = True
    
    navy_header_fill = PatternFill(start_color="2F3E46", end_color="2F3E46", fill_type="solid")
    font_main_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    
    b01_sub_fill = PatternFill(start_color="DCEEFF", end_color="DCEEFF", fill_type="solid")
    b01_row_fill_a = PatternFill(start_color="F2F8FF", end_color="F2F8FF", fill_type="solid")
    b01_row_fill_b = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
    
    b02_sub_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid") 
    b02_row_fill_a = PatternFill(start_color="F4F9F1", end_color="F4F9F1", fill_type="solid") 
    b02_row_fill_b = PatternFill(start_color="EBF5E6", end_color="EBF5E6", fill_type="solid") 
    
    p01_sub_fill = PatternFill(start_color="E8E1F5", end_color="E8E1F5", fill_type="solid") 
    p01_row_fill_a = PatternFill(start_color="F6F3FA", end_color="F6F3FA", fill_type="solid") 
    p01_row_fill_b = PatternFill(start_color="EFEAF6", end_color="EFEAF6", fill_type="solid") 

    font_sub_header = Font(name="Segoe UI", size=10, bold=True, color="333333")
    font_data = Font(name="Segoe UI", size=10, color="222222")
    thin_border_side = Side(style='thin', color='D9D9D9')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    segments = [
        ("Core Item Demographics (B01)", len(file_layout_dict_b01)),
        ("Logistics & Warehouse Paths (B02)", len(file_layout_dict_b02)),
        ("Pricing & Financial Metrics (P01)", len(file_layout_dict_p01))
    ]
    
    current_col = 1
    for name, length in segments:
        ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + length - 1)
        header_cell = ws.cell(row=1, column=current_col, value=name)
        header_cell.font = font_main_header
        header_cell.fill = navy_header_fill
        header_cell.alignment = Alignment(horizontal="center", vertical="center")
        current_col += length
    ws.row_dimensions[1].height = 26

    all_fields = list(file_layout_dict_b01.keys()) + list(file_layout_dict_b02.keys()) + list(file_layout_dict_p01.keys())
    len_b01 = len(file_layout_dict_b01)
    len_b02 = len(file_layout_dict_b02)
    
    for col_idx, field_key in enumerate(all_fields, start=1):
        clean_title = field_key.replace("_", " ").title()
        sub_cell = ws.cell(row=2, column=col_idx, value=clean_title)
        sub_cell.font = font_sub_header
        sub_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        sub_cell.border = cell_border
        
        if col_idx <= len_b01:
            sub_cell.fill = b01_sub_fill
        elif col_idx <= (len_b01 + len_b02):
            sub_cell.fill = b02_sub_fill
        else:
            sub_cell.fill = p01_sub_fill
            
    ws.row_dimensions[2].height = 28
    
    current_row = 3
    for product in products:
        prod_key = list(product.keys())[0]
        data_map = product[prod_key]
        
        row_values = []
        for f in file_layout_dict_b01.keys(): row_values.append(data_map["B01_entries"].get(f, ""))
        for f in file_layout_dict_b02.keys(): row_values.append(data_map["B02_entries"].get(f, ""))
        for f in file_layout_dict_p01.keys(): row_values.append(data_map["P01_entries"].get(f, ""))
        
        is_zebra_row = (current_row % 2 == 0)
        
        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = font_data
            cell.border = cell_border
            
            if col_idx <= len_b01:
                cell.fill = b01_row_fill_b if is_zebra_row else b01_row_fill_a
            elif col_idx <= (len_b01 + len_b02):
                cell.fill = b02_row_fill_b if is_zebra_row else b02_row_fill_a
            else:
                cell.fill = p01_row_fill_b if is_zebra_row else p01_row_fill_a
                
            if isinstance(val, float):
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif isinstance(val, int):
                cell.alignment = Alignment(horizontal="right")
                
        current_row += 1
        
    ws.freeze_panes = "D3"
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
        
    wb.save(filename)

# ==========================================
# CORE UI WINDOW CLASS
# ==========================================
class BK1ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BK1 Flat-File Parser")
        self.root.geometry("500x280")
        self.root.resizable(False, False)
        
        self.parsed_data = None
        self.loaded_filename = ""

        # Modern UI styling updates
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Main wrapper frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title Description
        title_lbl = ttk.Label(main_frame, text="Flat-File Catalog Transpiler", font=("Segoe UI", 14, "bold"))
        title_lbl.pack(pady=(0, 15))

        # File Select Area
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        self.file_label = ttk.Label(file_frame, text="No .bk# file loaded.", font=("Segoe UI", 10), wraplength=320)
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.upload_btn = ttk.Button(file_frame, text="Browse...", command=self.load_file)
        self.upload_btn.pack(side=tk.RIGHT, padx=5)

        # Progress Status Ring Indicator (Muted Label text)
        self.status_lbl = ttk.Label(main_frame, text="Awaiting flat file upload...", font=("Segoe UI", 9, "italic"), foreground="gray")
        self.status_lbl.pack(pady=15)

        # Save Button (Native Tkinter control used here to easily manipulate custom backgrounds)
        self.save_btn = tk.Button(
            main_frame, 
            text="Export to Excel", 
            font=("Segoe UI", 11, "bold"),
            bg="#D6D6D6", 
            fg="#757575", 
            state=tk.DISABLED, 
            relief="flat",
            command=self.save_file,
            cursor="arrow"
        )
        self.save_btn.pack(fill=tk.X, ipady=8, pady=(10, 0))

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Select product catalog file",
            filetypes=[("BK files", "*.bk1;*.bk2;*.bk3"), ("BK1 files", "*.bk1"), ("BK2 files", "*.bk2"), ("BK3 files", "*.bk3"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            self.status_lbl.config(text="Processing and validating streams...", foreground="blue")
            self.root.update_idletasks()
            
            # Fire parsing sequence engine
            self.parsed_data = parse_bk_file(file_path)
            self.loaded_filename = os.path.basename(file_path)
            
            # File loaded successfully state adjustments
            self.file_label.config(text=f"Loaded: {self.loaded_filename}")
            self.status_lbl.config(text=f"Success! Found {len(self.parsed_data)} composite items.", foreground="green")
            
            # Shift Save Button to Active Green State
            self.save_btn.config(
                state=tk.NORMAL, 
                bg="#2ECC71", 
                fg="white", 
                activebackground="#27AE60", 
                activeforeground="white",
                cursor="hand2"
            )
        except Exception as e:
            messagebox.showerror("Parsing Error", f"Failed to extract records from file:\n{str(e)}")
            self.status_lbl.config(text="Parsing failure.", foreground="red")

    def save_file(self):
        if not self.parsed_data:
            return
        
        # Open directory selection with filters for target options
        out_path = filedialog.asksaveasfilename(
            title="Export Transpiled Dataset",
            initialfile=os.path.splitext(self.loaded_filename)[0],
            filetypes=[("Excel Spreadsheet", "*.xlsx"), ("JSON Matrix Document", "*.json")]
        )
        if not out_path:
            return

        try:
            if out_path.endswith(".json"):
                with open(out_path, "w", encoding="utf-8") as f:
                    json.dump(self.parsed_data, f, indent=4)
            elif out_path.endswith(".xlsx"):
                export_to_excel(self.parsed_data, out_path)
            else:
                # Default safety extension catch
                out_path += ".xlsx"
                export_to_excel(self.parsed_data, out_path)
                
            messagebox.showinfo("Export Successful", f"File saved cleanly to:\n{os.path.basename(out_path)}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not write target document to file location:\n{str(e)}")

if __name__ == "__main__":
    window = tk.Tk()
    app = BK1ConverterApp(window)
    window.mainloop()