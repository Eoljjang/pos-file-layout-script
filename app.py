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
        product[name]["B01_entries"] = retrieve_all_entries(product[name]["B01"] or " "*210, file_layout_dict_b01)
        product[name]["B02_entries"] = retrieve_all_entries(product[name]["B02"] or " "*210, file_layout_dict_b02)
        product[name]["P01_entries"] = retrieve_all_entries(product[name]["P01"] or " "*240, file_layout_dict_p01)
        
    return products

# ==========================================
# EXCEL EXPORT ENGINE (VERTICAL APPEND)
# ==========================================
def export_to_excel(products, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Product Catalog Master"
    ws.views.sheetView[0].showGridLines = True
    
    font_section_header = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
    font_sub_header = Font(name="Segoe UI", size=10, bold=True, color="333333")
    font_data = Font(name="Segoe UI", size=10, color="222222")
    
    navy_header_fill = PatternFill(start_color="2F3E46", end_color="2F3E46", fill_type="solid")
    
    b01_sub_fill = PatternFill(start_color="DCEEFF", end_color="DCEEFF", fill_type="solid")
    b01_row_fill_a = PatternFill(start_color="F2F8FF", end_color="F2F8FF", fill_type="solid")
    b01_row_fill_b = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
    
    b02_sub_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid") 
    b02_row_fill_a = PatternFill(start_color="F4F9F1", end_color="F4F9F1", fill_type="solid") 
    b02_row_fill_b = PatternFill(start_color="EBF5E6", end_color="EBF5E6", fill_type="solid") 
    
    p01_sub_fill = PatternFill(start_color="E8E1F5", end_color="E8E1F5", fill_type="solid") 
    p01_row_fill_a = PatternFill(start_color="F6F3FA", end_color="F6F3FA", fill_type="solid") 
    p01_row_fill_b = PatternFill(start_color="EFEAF6", end_color="EFEAF6", fill_type="solid") 

    thin_border_side = Side(style='thin', color='D9D9D9')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    segments_config = [
        {"name": "Core Item Demographics (B01)", "layout": file_layout_dict_b01, "sub_fill": b01_sub_fill, "fill_a": b01_row_fill_a, "fill_b": b01_row_fill_b, "entry_key": "B01_entries"},
        {"name": "Logistics & Warehouse Paths (B02)", "layout": file_layout_dict_b02, "sub_fill": b02_sub_fill, "fill_a": b02_row_fill_a, "fill_b": b02_row_fill_b, "entry_key": "B02_entries"},
        {"name": "Pricing & Financial Metrics (P01)", "layout": file_layout_dict_p01, "sub_fill": p01_sub_fill, "fill_a": p01_row_fill_a, "fill_b": p01_row_fill_b, "entry_key": "P01_entries"}
    ]
    
    current_row = 1
    
    for segment in segments_config:
        layout_dict = segment["layout"]
        fields = list(layout_dict.keys())
        num_fields = len(fields)
        
        if num_fields == 0:
            continue
            
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_fields)
        header_cell = ws.cell(row=current_row, column=1, value=segment["name"])
        header_cell.font = font_section_header
        header_cell.fill = navy_header_fill
        header_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[current_row].height = 30
        current_row += 1
        
        for col_idx, field_key in enumerate(fields, start=1):
            clean_title = field_key.replace("_", " ").title()
            sub_cell = ws.cell(row=current_row, column=col_idx, value=clean_title)
            sub_cell.font = font_sub_header
            sub_cell.fill = segment["sub_fill"]
            sub_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            sub_cell.border = cell_border
            
        ws.row_dimensions[current_row].height = 26
        current_row += 1
        
        for idx, product in enumerate(products):
            prod_key = list(product.keys())[0]
            data_map = product[prod_key]
            
            is_zebra_row = (idx % 2 == 0)
            current_fill = segment["fill_b"] if is_zebra_row else segment["fill_a"]
            
            for col_idx, field_key in enumerate(fields, start=1):
                val = data_map[segment["entry_key"]].get(field_key, "")
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = font_data
                cell.fill = current_fill
                cell.border = cell_border
                
                if isinstance(val, float):
                    cell.number_format = "$#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif isinstance(val, int):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
            ws.row_dimensions[current_row].height = 20
            current_row += 1
            
        current_row += 2 

    ws.freeze_panes = "A3" 
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value and "Section" not in str(cell.value) and len(str(cell.value)) > max_len:
                max_len = len(str(cell.value))
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
        
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
        
        # Color definitions
        self.bg_base = "#1F232A"
        self.bg_surface = "#2D3139"
        self.text_primary = "#E2E8F0"
        self.text_secondary = "#8E8E93" 
        self.ios_blue = "#0A84FF"       
        self.ios_green = "#30D158"      
        self.ios_red = "#FF453A"        
        
        self.root.configure(bg=self.bg_base)
        
        try:
            from ctypes import windll, byref, c_int, sizeof
            HWND = windll.user32.GetParent(self.root.winfo_id())
            windll.dwmapi.DwmSetWindowAttribute(HWND, 20, byref(c_int(1)), sizeof(c_int))
        except Exception:
            pass

        self.parsed_data = None
        self.loaded_filename = ""
        self.font_family = ("-apple-system", "SF Pro Text", "Helvetica Neue", "Segoe UI", "Arial")

        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.style.configure("TFrame", background=self.bg_base)
        self.style.configure("Surface.TFrame", background=self.bg_surface)
        self.style.configure("TLabel", background=self.bg_base, foreground=self.text_primary)
        self.style.configure("Surface.TLabel", background=self.bg_surface, foreground=self.text_primary)
        
        self.style.configure(
            "IOS.TButton", 
            background=self.bg_surface, 
            foreground=self.ios_blue, 
            font=(self.font_family, 10, "bold"),
            borderwidth=1,
            bordercolor=self.ios_blue,
            lightcolor=self.ios_blue,
            darkcolor=self.ios_blue,
            focuscolor=self.bg_surface,
            padding=(10, 4)
        )
        self.style.map(
            "IOS.TButton", 
            background=[("active", "#3A3F4B")],
            foreground=[("active", "#64B5FF")],
            bordercolor=[("active", "#64B5FF")],
            lightcolor=[("active", "#64B5FF")],
            darkcolor=[("active", "#64B5FF")]
        )
        
        self.style.layout("IOS.TButton", [
            ('Button.border', {'sticky': 'nswe', 'children': [
                ('Button.focus', {'sticky': 'nswe', 'children': [
                    ('Button.padding', {'sticky': 'nswe', 'children': [
                        ('Button.label', {'sticky': 'nswe'})
                    ]})
                ]})
            ]})
        ])
        
        main_frame = ttk.Frame(root, padding="24")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_lbl = ttk.Label(main_frame, text="TGS POS File Layout Converter", font=(self.font_family, 16, "bold"))
        title_lbl.pack(pady=(0, 20), anchor="w")

        file_frame = ttk.Frame(main_frame, style="Surface.TFrame", padding="12")
        file_frame.pack(fill=tk.X, pady=5)
        
        self.file_label = ttk.Label(
            file_frame, 
            text="No .bk# file loaded.", 
            font=(self.font_family, 10), 
            wraplength=300, 
            style="Surface.TLabel"
        )
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.upload_btn = ttk.Button(file_frame, text="Browse...", style="IOS.TButton", command=self.load_file)
        self.upload_btn.pack(side=tk.RIGHT, padx=5)

        # Multi-use Dynamic Status Label Configuration
        self.status_lbl = ttk.Label(
            main_frame, 
            text="Awaiting flat file upload...", 
            font=(self.font_family, 10), 
            foreground=self.text_secondary,
            wraplength=440,
            justify="center"
        )
        self.status_lbl.pack(pady=15, anchor="center")

        self.save_btn = tk.Button(
            main_frame, 
            text="Export to Excel", 
            font=(self.font_family, 11, "bold"),
            bg="#2C2C2E", 
            fg="#48484A", 
            state=tk.DISABLED, 
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            bd=0,
            command=self.save_file,
            cursor="arrow"
        )
        self.save_btn.pack(fill=tk.X, ipady=10, pady=(10, 0))

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Select product catalog file",
            filetypes=[("BK files", "*.bk1;*.bk2;*.bk3"), ("BK1 files", "*.bk1"), ("BK2 files", "*.bk2"), ("BK3 files", "*.bk3"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            self.status_lbl.config(text="Processing and validating streams...", foreground=self.ios_blue)
            self.root.update_idletasks()
            
            self.parsed_data = parse_bk_file(file_path)
            self.loaded_filename = os.path.basename(file_path)
            
            self.file_label.config(text=f"Loaded: {self.loaded_filename}")
            self.status_lbl.config(text=f"Success! Found {len(self.parsed_data)} composite items.", foreground=self.ios_green) 
            
            self.save_btn.config(
                state=tk.NORMAL, 
                bg=self.ios_blue, 
                fg="white", 
                activebackground="#0070E6", 
                activeforeground="white",
                cursor="hand2"
            )
        except Exception as e:
            self.status_lbl.config(text=f"Parsing failure: {str(e)}", foreground=self.ios_red) 

    def save_file(self):
        if not self.parsed_data:
            return
        
        out_path = filedialog.asksaveasfilename(
            title="Export Transpiled Dataset",
            initialfile=os.path.splitext(self.loaded_filename)[0],
            filetypes=[("Excel Spreadsheet", "*.xlsx"), ("JSON Matrix Document", "*.json")]
        )
        if not out_path:
            return

        try:
            self.status_lbl.config(text="Writing data layers to file destination...", foreground=self.ios_blue)
            self.root.update_idletasks()

            if out_path.endswith(".json"):
                with open(out_path, "w", encoding="utf-8") as f:
                    json.dump(self.parsed_data, f, indent=4)
            else:
                if not out_path.endswith(".xlsx"):
                    out_path += ".xlsx"
                export_to_excel(self.parsed_data, out_path)
                
            # Direct In-App Success Status Placement
            self.status_lbl.config(
                text=f"Excel file saved successfully to: {os.path.basename(out_path)}", 
                foreground=self.ios_green
            )
        except Exception as e:
            # Inline Error Delivery Optimization Context
            error_msg = str(e)
            if "Permission denied" in error_msg:
                error_msg = "Permission Denied. Close the file if it's open in Excel and retry."
            
            self.status_lbl.config(
                text=f"Export failed: {error_msg}", 
                foreground=self.ios_red
            )

if __name__ == "__main__":
    window = tk.Tk()
    app = BK1ConverterApp(window)
    window.mainloop()