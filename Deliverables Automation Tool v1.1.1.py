import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import openpyxl
import csv
import os
import xlwings as xw
from datetime import datetime

# Deliverables Automation Tool with Wafermap
# Author: Rose Anne Lafuente
# Licensed Electronics Engineer | Product Engineer II | Python Automation
# Description:
#   Automates CSV-to-Excel workflows with pivot tables, custom formatting,
#   End Test validation, and wafermap visualization for yield and defect tracking.
#   Features include:
#     - Scrollable status box for enhanced log navigation
#     - Deterministic wafermap coloring via defined C1_MARK color_map
#     - Accurate C1_MARK lookup for ET mapping
#     - GUI title and developer label for professional branding
#   Built with Python, Tkinter, OpenPyXL, and xlwings.

class AutomatingDeliverables:
    def __init__(self, root):
        self.root = root
        self.root.title("Automating Deliverables")
        self.root.geometry("800x550")

        # Professional Neutral Theme
        self.bg_color = "#f5f5f5"
        self.fg_color = "#222222"
        self.entry_bg = "#ffffff"
        self.btn_bg = "#e0e0e0"
        self.btn_active = "#BEE395"

        self.root.configure(bg=self.bg_color)

        # --- Title Frame ---
        title_frame = tk.Frame(self.root, bg=self.bg_color)
        title_frame.pack(pady=(10,0))

        # Product name (bold)
        title_label = tk.Label(
            title_frame,
            text="Automating Deliverables",
            font=("Meiryo", 12, "bold"),
            fg="darkblue",
            bg=self.bg_color
        )
        title_label.pack(side="left")

        # Version (italic)
        version_label = tk.Label(
            title_frame,
            text=" v1.1.0",
            font=("Meiryo", 12, "italic"),
            fg="darkblue",
            bg=self.bg_color
        )
        version_label.pack(side="left")

        # --- Subtle 'Developed by' line just below title ---
        dev_label = tk.Label(
            self.root,
            text="Developed by Rose Anne Lafuente | 2026",
            font=("Arial", 7, "italic"),   # very small font
            fg="gray",
            bg=self.bg_color
        )
        dev_label.pack(pady=(0,10))

        self.path_var = tk.StringVar()

        # Build the rest of the interface
        self.create_file_selection_frame()
        self.create_filter_selector([])   # Show filter selector immediately (empty at first)
        self.create_status_box()
        self.create_exit_button()

    def create_file_selection_frame(self):
        # File Selection frame with subtle border and spacing
        input_frame = tk.LabelFrame(
            self.root,
            text="File Selection",
            padx=10, pady=10,
            bd=2,
            relief="groove",
            font=("Segoe UI", 10, "bold")
        )
        input_frame.pack(fill="x", padx=15, pady=10)

        # Label inside the frame
        label = tk.Label(input_frame, text="Select CSV File:")
        label.pack(side="left", padx=(0, 10), pady=5)

        # Entry box inside the frame
        path_entry = tk.Entry(input_frame, textvariable=self.path_var,
                              bg="white", fg="black", insertbackground="black")
        path_entry.pack(side="left", padx=10, pady=5, fill="x", expand=True)

        # ✅ Convert button inside the same frame
        convert_btn = tk.Button(
            input_frame,
            text="Convert to Excel",
            width=18,
            command=self.convert_to_excel,
            bg=self.btn_bg,
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        convert_btn.pack(side="right", padx=10, pady=5)
        # Browse button inside the frame
        browse_btn = tk.Button(input_frame, text="Browse", width=12, command=self.browse_file)
        browse_btn.pack(side="right", pady=5)

    def get_unique_c1_mark_values(raw_items):
        flat = []
        for item in raw_items:
            if isinstance(item, list):   # flatten nested lists
                flat.extend(item)
            elif item is not None:
                flat.append(item)

        # Strip whitespace but keep case
        cleaned = [str(i).strip() for i in flat if i]

        # Deduplicate while preserving order (case-sensitive)
        unique = list(dict.fromkeys(cleaned))
        return unique
    
    def create_filter_selector(self, items):
        # Pivot Filter Selection frame with subtle border and spacing
        filter_frame = tk.LabelFrame(
            self.root,
            text="Pivot Filter Selection",
            padx=10, pady=10,
            bd=2,
            relief="groove",
            font=("Segoe UI", 10, "bold")
        )

        filter_frame.pack(fill="x", padx=15, pady=10)

        tk.Label(
            filter_frame,
            text="Select C1_MARK:"
        ).pack(side="left", padx=5, expand=True, fill="x")

        # ✅ Clean and deduplicate items
        clean_items = [str(i) for i in items if i is not None]
        unique_items = list(dict.fromkeys(clean_items))

        self.filter_var = tk.StringVar()
        self.filter_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.filter_var,
            values=unique_items,
            state="readonly",
            width=25
        )
        self.filter_dropdown.pack(side="left", padx=10)

        gen_pivot_btn = tk.Button(
            filter_frame,
            text="Generate Pivot Table",
            width=18,
            command=self.generate_pivot,
            bg="#8BD3E6",
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        gen_pivot_btn.pack(side="left", padx=10, expand=True, fill="x")


        check_test_btn = tk.Button(
            filter_frame,
            text="Check End Test No",
            width=18,
            command=self.check_end_test,
            bg="#E6E6FA",
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        check_test_btn.pack(side="left", padx=10, expand=True, fill="x")


        gen_wafermap_btn = tk.Button(
            filter_frame,
            text="Generate Wafermap",
            width=18,
            command=self.generate_wafermap,
            bg="#92D050",
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        gen_wafermap_btn.pack(side="left", padx=10, expand=True, fill="x")

    def create_status_box(self):
        # Frame to hold text + scrollbars
        status_frame = tk.LabelFrame(self.root, text="", padx=10, pady=10)
        status_frame.pack(fill="both", expand=True, padx=15, pady=10)

        # Create a container frame for grid layout
        container = tk.Frame(status_frame)
        container.pack(fill="both", expand=True)

        # Text box
        self.status_box = tk.Text(
            container,
            height=10,
            wrap="word",
            bg="white",
            fg="black",
            state="disabled"
        )

        # Scrollbars
        self.status_vsb = tk.Scrollbar(container, orient="vertical", command=self.status_box.yview)
        self.status_hsb = tk.Scrollbar(container, orient="horizontal", command=self.status_box.xview)

        # Link scrollbars to text box
        self.status_box.configure(
            yscrollcommand=self.status_vsb.set,
            xscrollcommand=self.status_hsb.set
        )

        # Layout with grid
        self.status_box.grid(row=0, column=0, sticky="nsew")
        self.status_vsb.grid(row=0, column=1, sticky="ns")
        self.status_hsb.grid(row=1, column=0, sticky="ew")

        # Make the text box expand with the frame
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def create_exit_button(self):
        # Create an "invisible" frame with same background as root
        exit_frame = tk.Frame(self.root, bg=self.bg_color)
        exit_frame.pack(fill="x", side="bottom", padx=15, pady=5)

        # Place Exit button aligned right
        exit_btn = tk.Button(exit_frame, text="EXIT", width=12,
                             bg="#d32f2f", fg="white", command=self.root.destroy)
        exit_btn.pack(side="right", pady=10)

        clear_btn = tk.Button(exit_frame, text="Clear All", width=12,
                      command=self.clear_all,
                      bg="#ffcccc", fg=self.fg_color, activebackground=self.btn_active)
        clear_btn.pack(side="right", padx=10)

    def show_status(self, message, color=None, clear=False):
        # Default to black unless explicitly set to red
        if color is None:
            color = "#000000"  # black

        self.status_box.config(state="normal")

        if clear:
            self.status_box.delete("1.0", "end")

        if message:
            self.status_box.insert("end", message + "\n")

            # Unique tag per line so colors don't overwrite
            line_tag = f"status_{self.status_box.index('end-2l')}"
            start_index = self.status_box.index("end-2l")
            end_index = self.status_box.index("end-1c")
            self.status_box.tag_add(line_tag, start_index, end_index)
            self.status_box.tag_config(line_tag, foreground=color)

        self.status_box.config(state="disabled")
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            self.path_var.set(file_path)
            # Just show status that file is selected
            self.show_status(f"📂 Selected file:{file_path}", color="black")
            
    def convert_to_excel(self):
        file_path = self.path_var.get()
        if not file_path:
            self.show_status("⚠️ No file selected. Please browse for a CSV first.", color="#d32f2f")
            return

        try:
            # --- Convert CSV to Excel (vectorized) ---
            wb = openpyxl.Workbook()
            ws = wb.active

            sheet_name = os.path.splitext(os.path.basename(file_path))[0]
            ws.title = sheet_name[:31].replace(":", "_").replace("/", "_").replace("\\", "_")

            # Read CSV into list of lists
            with open(file_path, newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = []
                for row in reader:
                    parsed = []
                    for value in row:
                        try:
                            if value.isdigit():
                                parsed.append(int(value))
                            else:
                                parsed.append(float(value))
                        except ValueError:
                            parsed.append(value)
                    rows.append(parsed)

            # Vectorized write: append all rows
            for r in rows:
                ws.append(r)

            out_file = os.path.splitext(file_path)[0] + ".xlsx"
            wb.save(out_file)
            wb.close()

            # --- Open with xlwings to read filter items ---
            app = xw.App(visible=False)
            wb_xlw = app.books.open(out_file)
            sht = wb_xlw.sheets[0]

            # Scan all of column G until we find "C1_MARK"
            col_g = sht.range("G1:G" + str(sht.cells.last_cell.row)).value
            header_row = None
            for i, val in enumerate(col_g, start=1):
                if str(val).strip() == "C1_MARK":
                    header_row = i
                    break

            if not header_row:
                self.show_status("❌ 'C1_MARK' not found in Column G.", color="#d32f2f")
                wb_xlw.close()
                app.quit()
                return

            # Collect filter items from C1_MARK column
            last_row = sht.range((header_row, 7)).end("down").row
            raw_items = sht.range((header_row+1, 7), (last_row, 7)).value

            # Deduplicate case-sensitive
            flat = [str(i).strip() for i in raw_items if i]
            unique_items = list(dict.fromkeys(flat))

            self.filter_dropdown['values'] = unique_items
            self.out_file = out_file
            self.base_name = sheet_name

            wb_xlw.close()
            app.quit()

            self.show_status(f"\n✅ Conversion complete: CSV → .xlsx\nFile saved at: {out_file}\n\nFilter options loaded.")

        except Exception as e:
            self.show_status(f"❌ Error: {e}", color="#d32f2f")

    def generate_pivot(self):
        selected = self.filter_var.get()
        if not selected:
            self.show_status("⚠️ Please select a C1_MARK value first.", color="#d32f2f")
            return
        
        self.show_status(f"\nℹ️ Generating pivot table...")

        app = None
        wb_xlw = None
        try:
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)
            sht = wb_xlw.sheets[self.base_name]

            # Scan all of column G until we find "C1_MARK"
            col_g = sht.range("G1:G" + str(sht.cells.last_cell.row)).value
            header_row = None
            for i, val in enumerate(col_g, start=1):
                if str(val).strip() == "C1_MARK":
                    header_row = i
                    break

            if not header_row:
                self.show_status("❌ 'C1_MARK' not found in Column G.", color="#d32f2f")
                wb_xlw.close()
                app.quit()
                return

            # --- Find ET header ---
            row_values = sht.range((header_row, 7), (header_row, sht.range((header_row, 7)).end("right").column)).value
            et_col = None
            for idx, val in enumerate(row_values, start=7):
                if str(val).strip().upper() == "ET":
                    et_col = idx
                    break
            if not et_col:
                raise ValueError("'ET' column not found to the right of C1_MARK")

            # --- Define pivot source range ---
            last_row = sht.range((header_row, 7)).end("down").row
            pivot_range = sht.range((header_row, 7), (last_row, et_col))

            # --- Create Pivot sheet ---
            pivot_sheet = wb_xlw.sheets.add("Pivot", after=sht)

            # --- Create pivot cache and table ---
            pivot_cache = wb_xlw.api.PivotCaches().Create(SourceType=1, SourceData=pivot_range.api)
            table_name = f"PivotTable_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            pivot_table = pivot_cache.CreatePivotTable(TableDestination=pivot_sheet.range("A3").api, TableName=table_name)

            # --- Filter: C1_MARK ---
            pf = pivot_table.PivotFields("C1_MARK")
            pf.Orientation = 3
            valid_items = [item.Name for item in pf.PivotItems()]
            if selected in valid_items:
                pf.CurrentPage = selected
                self.show_status(f"\nApplied filter: {selected}")
            else:
                self.show_status(f"⚠️ Selected '{selected}' not found in C1_MARK items {valid_items}", color="#d32f2f")
                return

            # --- Rows: ET ---
            pivot_table.PivotFields("ET").Orientation = 1

            # --- Values: Count of FT ---
            pivot_table.AddDataField(pivot_table.PivotFields("FT"), "Count of FT", -4112)

            # --- Fallout Table Logic ---
            data = pivot_sheet.range("A4").expand().value
            sheet = wb_xlw.sheets[self.base_name]
            theoretical_num = None
            for i, val in enumerate(sheet.range("A:A").value, start=1):
                if str(val).strip().upper() == "THEORETICAL_NUM":
                    theoretical_num = sheet.range((i, 1)).offset(0, 2).value
                    break

            fallout_table = []
            for row in data:
                if not row or not row[0] or str(row[0]).strip().lower() == "grand total":
                    continue
                et_val = str(int(row[0])) if isinstance(row[0], (int, float)) and float(row[0]).is_integer() else str(row[0])
                count_val = int(row[1]) if isinstance(row[1], (int, float)) and float(row[1]).is_integer() else row[1]
                fallout = (float(row[1]) / theoretical_num * 100) if theoretical_num else 0
                fallout_table.append([et_val, count_val, f"{fallout:.2f}%"])

            fallout_table.sort(key=lambda x: int(x[1]), reverse=True)
            grand_total_val = str(int(theoretical_num)) if isinstance(theoretical_num, (int, float)) and float(theoretical_num).is_integer() else str(theoretical_num)
            fallout_table.insert(0, ["End Test No.", "Count", "Fallout%"])  # header row
            fallout_table.append(["Grand Total", grand_total_val, ""])

            # --- Vectorized write fallout table ---
            pivot_sheet.range("D3").value = fallout_table

            # --- Apply formatting ---
            last_row_ft = 3 + len(fallout_table) - 1
            fallout_range = pivot_sheet.range(f"D3:F{last_row_ft}")
            fallout_range.api.HorizontalAlignment = -4108
            fallout_range.api.VerticalAlignment = -4108
            fallout_range.api.IndentLevel = 0

            # Header row
            pivot_sheet.range("D3:F3").color = (192, 230, 245)
            pivot_sheet.range("D3:F3").api.Font.Bold = True
            # First data row
            pivot_sheet.range("D4:F4").color = (255, 159, 159)
            pivot_sheet.range("D4:F4").api.Font.Bold = True
            # Grand Total row
            pivot_sheet.range(f"D{last_row_ft}:F{last_row_ft}").color = (192, 230, 245)
            pivot_sheet.range(f"D{last_row_ft}:F{last_row_ft}").api.Font.Bold = True

            fallout_range.api.Borders.Weight = 2

            wb_xlw.save()

            # --- Show fallout table in status box ---
            self.status_box.config(state="normal")
            self.status_box.insert(tk.END, "\nPreview Table:\n")
            for et_val, count_val, fallout_val in fallout_table:
                self.status_box.insert(tk.END, f"{str(et_val):<15}{str(count_val):<10}{str(fallout_val)}\n")
            self.status_box.config(state="disabled")

            self.show_status(f"\n✅ Succesfully generated table for C1_MARK:{selected}")

        except Exception as e:
            self.show_status(f"❌ Error generating pivot/fallout: {e}", color="#d32f2f")

        finally:
            if wb_xlw:
                try: wb_xlw.close()
                except: pass
            if app:
                try: app.quit()
                except: pass

    def check_end_test(self):
        app = None
        wb_xlw = None
        try:
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)

            # --- Ensure Pivot sheet exists ---
            try:
                pivot_sheet = wb_xlw.sheets["Pivot"]
            except:
                pivot_sheet = wb_xlw.sheets.add("Pivot")

            data_sheet = wb_xlw.sheets[self.base_name]

            # --- Get highest fails End Test No from D4 ---
            raw_val = pivot_sheet.range("D4").value
            if raw_val is None:
                end_test_no = ""
            elif isinstance(raw_val, float) and raw_val.is_integer():
                end_test_no = str(int(raw_val))
            else:
                end_test_no = str(raw_val).strip()

            self.show_status(f"\n🔍Checking End Test No.: {end_test_no}")

            # --- Find LOLIMIT row in Column F ---
            lolimit_row = data_sheet.range("F1").end("down").row
            lolimit_val = data_sheet.range(f"F{lolimit_row}").value
            if str(lolimit_val).strip().upper() != "LOLIMIT":
                raise ValueError("LOLIMIT not found in Column F")

            # --- Expand reference table ---
            ref_range = data_sheet.range((lolimit_row, 1)).expand("table")

            # --- Locate TESTNO column (Column B) ---
            testno_values = data_sheet.range(
                (lolimit_row + 1, 2),
                (lolimit_row + ref_range.rows.count - 1, 2)
            ).value

            # Normalize TESTNO values to strings
            testno_values = ["" if v is None else str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip() for v in testno_values]

            found_row = None
            if end_test_no in testno_values:
                idx = testno_values.index(end_test_no) + lolimit_row + 1
                found_row = idx

            # --- Vectorized write of header + data ---
            start_cell = pivot_sheet.range("H3")
            header = ["TSNO", "TESTNO", "COMMENT", "MODE", "HILIMIT", "LOLIMIT"]

            if found_row:
                row_values = data_sheet.range((found_row, 1), (found_row, 6)).value
                row_values = ["" if v is None else str(v).strip() for v in row_values]

                # Write header + data in one call
                pivot_sheet.range("H3").value = [header, row_values]

                # --- Apply formatting in bulk ---
                ref_range_excel = pivot_sheet.range("H3:M4")
                header_range = pivot_sheet.range("H3:M3")
                data_range = pivot_sheet.range("H4:M4")

                header_range.color = (192, 230, 245)   # light blue
                header_range.api.Font.Bold = True
                data_range.color = (255, 255, 255)     # white
                data_range.api.Font.Bold = True        # bold data row

                # Borders + alignment
                ref_range_excel.api.Borders.Weight = 2
                ref_range_excel.api.HorizontalAlignment = -4108
                ref_range_excel.api.VerticalAlignment = -4108
                ref_range_excel.api.IndentLevel = 0

                wb_xlw.save()

                # --- Show End Test No. table in status box ---
                self.status_box.config(state="normal")
                self.status_box.insert(tk.END, "\nEnd Test No. Reference:\n")
                self.status_box.insert(
                    tk.END,
                    f"{'TSNO':<10}{'TESTNO':<10}{'COMMENT':<15}{'MODE':<10}{'HILIMIT':<10}{'LOLIMIT'}\n"
                )
                self.status_box.insert(tk.END, "-" * 70 + "\n")
                tsno, testno, comment, mode, hilimit, lolimit = row_values
                self.status_box.insert(
                    tk.END,
                    f"{tsno:<10}{testno:<10}{comment:<15}{mode:<10}{hilimit:<10}{lolimit}\n"
                )
                self.status_box.config(state="disabled")

                # --- Status message depending on limits ---
                if lolimit != "":
                    self.show_status("\n✅ Found with Limits")
                else:
                    self.show_status("\n⚠️ Found with no Limit", color="#FFBF00")
            else:
                self.show_status("\n❌ No End Test No. found in the TESTNO Column", color="#d32f2f")

        except Exception as e:
            self.show_status(f"\n❌ Error checking End Test No: {e}", color="#d32f2f")

        finally:
            if wb_xlw:
                try: wb_xlw.close()
                except: pass
            if app:
                try: app.quit()
                except: pass


    def generate_wafermap(self):
        app = None
        wb_xlw = None
        try:
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)
            data_sheet = wb_xlw.sheets[self.base_name]

            # --- SLOT handling ---
            slot_row = None
            for i, val in enumerate(data_sheet.range("A:A").value, start=1):
                if str(val).strip().upper() == "SLOT":
                    slot_row = i
                    break

            if not slot_row:
                self.show_status("\n⚠️ SLOT header not found in Column A", color="#d32f2f")
                return

            slot_val = data_sheet.range((slot_row+1, 1)).value
            if slot_val is None:
                self.show_status("\n⚠️ SLOT value below header is empty", color="#d32f2f")
                return

            slot_str = str(int(slot_val)).zfill(2)
            self.show_status(f"\n🔍 Generating wafermap for W #{slot_str}...")
            sheet_name = f"W#{slot_str}_wafermap_by_End_Test_No"

            # --- Disable gridlines ---
            #data_sheet.api.Parent.Windows(1).DisplayGridlines = False
            wb_xlw.save()

            # --- Create or reuse Wafermap Pivot Table sheet ---
            try:
                pivot_sheet = wb_xlw.sheets["Wafermap Pivot Table"]
                pivot_sheet.clear()
            except:
                pivot_sheet = wb_xlw.sheets.add("Wafermap Pivot Table", after=data_sheet)

            # --- Create or reuse slot-specific wafermap sheet ---
            try:
                wafermap_sheet = wb_xlw.sheets[sheet_name]
                wafermap_sheet.clear()
            except:
                wafermap_sheet = wb_xlw.sheets.add(sheet_name, after=pivot_sheet)

            # Scan all of column G until we find "C1_MARK"
            col_g = data_sheet.range("G1:G" + str(data_sheet.cells.last_cell.row)).value
            header_row = None
            for i, val in enumerate(col_g, start=1):
                if str(val).strip() == "C1_MARK":
                    header_row = i
                    break

            if not header_row:
                self.show_status("❌ 'C1_MARK' not found in Column G.", color="#d32f2f")
                wb_xlw.close()
                app.quit()
                return

            # --- Read header row ---
            row_values = data_sheet.range(
                (header_row, 1),
                (header_row, data_sheet.range((header_row, 1)).end("right").column)
            ).value

            # --- Locate X, Y, ET columns ---
            x_col = y_col = et_col = None
            for idx, val in enumerate(row_values, start=1):
                if str(val).strip().upper() == "X":
                    x_col = idx
                elif str(val).strip().upper() == "Y":
                    y_col = idx
                elif str(val).strip().upper() in ["ET", "END TEST NO."]:
                    et_col = idx

            if not (x_col and y_col and et_col):
                raise ValueError("Required columns 'X', 'Y', 'ET' not found in header row")

            # --- Define pivot source range ---
            last_row = data_sheet.range((header_row+1, et_col)).end("down").row
            pivot_range = data_sheet.range((header_row, x_col), (last_row, et_col))

            # --- Build ET → C1_MARK mapping right here ---
            et_to_c1 = {}
            for row in range(header_row+1, last_row+1):
                et_val = data_sheet.range((row, et_col)).value
                c1_val = data_sheet.range((row, 7)).value  # Column G = C1_MARK
                if et_val is None or c1_val is None:
                    continue

                # Normalize ET
                if isinstance(et_val, float) and et_val.is_integer():
                    et_str = str(int(et_val))
                else:
                    et_str = str(et_val).strip()

                c1_str = str(c1_val).strip()

                # Log with coordinates
                # self.show_status(f"C1_MARK({row},7) = {c1_str} | ET({row},{et_col}) = {et_str}")

                # Save mapping
                et_to_c1[et_str] = c1_str

            # Debug: show dictionary once built
            # self.show_status(f"ET→C1_MARK dictionary built: {et_to_c1}")

            # --- Create pivot cache and table ---
            pivot_cache = wb_xlw.api.PivotCaches().Create(SourceType=1, SourceData=pivot_range.api)
            table_name = f"PivotTable_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=pivot_sheet.range("A1").api,
                TableName=table_name
            )

            # --- Configure pivot ---
            pivot_table.PivotFields("Y").Orientation = 1
            pivot_table.PivotFields("X").Orientation = 2
            pivot_table.AddDataField(pivot_table.PivotFields("ET"), "Min of ET", -4139)
            pivot_table.ColumnGrand = False
            pivot_table.RowGrand = False
            pivot_sheet.range("A2").value = "No."

            # --- Copy pivot output ---
            pivot_block = pivot_sheet.range("A2").expand()
            data_block = pivot_block.value

            # --- Paste values into wafermap sheet ---
            rows = len(data_block)
            cols = len(data_block[0])
            wafermap_sheet.range((1,1), (rows,cols)).value = data_block

            # --- Alignment ---
            wafermap_sheet.range((1,1), (rows,cols)).api.HorizontalAlignment = -4108
            wafermap_sheet.range((1,1), (rows,cols)).api.VerticalAlignment = -4108

            # --- Find last used row/col ---
            last_col = wafermap_sheet.range("1:1").end("right").column
            last_row = wafermap_sheet.range("A:A").end("down").row

            # --- Header formatting ---
            dark_blue = xw.utils.rgb_to_int((46, 110, 158))
            wafermap_sheet.range((1,1),(1,last_col)).color = (228, 241, 253)
            wafermap_sheet.range((1,1),(1,last_col)).api.Font.Color = dark_blue
            wafermap_sheet.range((1,1),(last_row,1)).color = (228, 241, 253)
            wafermap_sheet.range((1,1),(last_row,1)).api.Font.Color = dark_blue

            # --- Define C1_MARK color mapping dictionary ---
            color_map = {
                "/":"#00FF00",   # updated from 1007
                "$":"#7B68EE",   # updated from 977
                "*":"#87CEEB",   # updated from 977
                "?":"#66FF66",
                "=":"#7FFFD4",   # updated for ET 1001,1002,1005,1006
                "!":"#6495ED",   # updated from 1003
                "#":"#6A5ACD",   # updated from 977
                "%":"#66FF66",
                ".":"#66FF66",
                ":":"#66FF66",
                "^":"#66FF66",
                "+":"#66FF66",
                "-":"#66FF66",
                "{":"#66FF66",
                "}":"#66FF66",
                "(":"#66FF66",
                ")":"#66FF66",
                "_":"#66FF66",
                "|":"#66FF66",
                ";":"#66FF66",
                "@":"#66FF66",
                "\\":"#66FF66",
                "<":"#66FF66",
                ">":"#66FF66",
                "&":"#66FF66",

                "0":"#66FF66",
                "1":"#FFFF99",
                "2":"#FF0000",
                "3":"#FFFFE0",   # updated from ET 977
                "4":"#ADD8E6",   # updated from ET 977
                "5":"#FF8080",
                "6":"#AFEEEE",   # updated from ET 110
                "7":"#99CCFF",
                "8":"#FFCC00",
                "9":"#FFFF00",

                "A":"#2E8B57",   # updated from ET 3
                "B":"#FFCC00",
                "C":"#FFCC00",
                "D":"#99CC00",
                "E":"#99CC00",
                "F":"#7CFC00",   # updated from ET 977
                "G":"#FFFF00",
                "H":"#A6A6A6",
                "I":"#00CCFF",
                "J":"#32CD32",   # updated from ET 977
                "K":"#20B2AA",   # updated from ET 977
                "L":"#FFDEAD",   # updated from ET 977
                "M":"#D9D9D9",
                "N":"#DAA520",   # updated from ET 977
                "O":"#00CCFF",
                "P":"#FFFF99",
                "Q":"#ED7D31",
                "R":"#FFCC00",
                "S":"#FF7C80",
                "T":"#FFCC00",
                "U":"#00CCFF",
                "V":"#008080",
                "W":"#008080",
                "X":"#008080",
                "Y":"#666699",
                "Z":"#666699",

                "a":"#D2691E",   # updated from ET 977
                "b":"#993366",
                "c":"#A52A2A",   # updated from ET 977
                "d":"#E9967A",   # updated from ET 977
                "e":"#660066",
                "f":"#ED7D31",
                "g":"#3366FF",
                "h":"#CCFFFF",
                "i":"#FF7F50",   # updated from ET 977
                "j":"#99CCFF",
                "k":"#CCCCFF",
                "l":"#D9D9D9",
                "m":"#969696",
                "n":"#339966",
                "o":"#333399",
                "p":"#FF6600",
                "q":"#FFFF00",
                "r":"#0066CC",
                "s":"#FF9900",
                "t":"#33CCCC",
                "u":"#008080",
                "v":"#EE82EE",   # updated from ET 977
                "w":"#DDA0DD",   # updated from ET 977
                "x":"#00FFFF",
                "y":"#99CC00",
                "z":"#9932CC"    # updated from ET 977
            }

            # --- Apply colors to wafermap cells using ET → C1_MARK mapping ---
            for r in range(2, last_row+1):
                for c in range(2, last_col+1):
                    cell = wafermap_sheet.range((r,c))
                    et_val = cell.value
                    if et_val is None or str(et_val).strip() == "":
                        continue

                    # Normalize ET
                    if isinstance(et_val, float) and et_val.is_integer():
                        et_str = str(int(et_val))
                    else:
                        et_str = str(et_val).strip()

                    # Lookup C1_MARK from dictionary
                    c1_mark_str = et_to_c1.get(str(et_str))  # lookup still safe as string


                    # Normalize the result if numeric
                    if c1_mark_str is not None:
                        try:
                            f = float(c1_mark_str)
                            if f.is_integer():
                                c1_mark_str = str(int(f))  # convert 1.0 → 1
                        except ValueError:
                            # Not numeric (special chars, mixed letters), leave as-is
                            pass

                        #self.show_status(f"C1_MARK | ET → {c1_mark_str} | {et_str}")

                        # Apply color based on C1_MARK
                        if c1_mark_str in color_map:
                            hex_color = color_map[c1_mark_str].lstrip("#")
                            rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                            cell.color = rgb
                        else:
                            self.show_status(f"⚠️ No color mapping for C1_MARK '{c1_mark_str}'", color="#d32f2f")
                            cell.color = (200,200,200)
                    else:
                        self.show_status(f"⚠️ No C1_MARK found for ET '{et_str}'", color="#d32f2f")
                        cell.color = (200,200,200)
                                
            # --- Copy Row 1 (Ctrl+Shift+Right) and paste it after last used row ---
            row1_vals = wafermap_sheet.range((1,1),(1,last_col)).value
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).value = row1_vals
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).color = (228,241,253)
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).api.Font.Color = dark_blue
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).api.Font.Bold = True  # bold copy of Row 1

            # --- Copy Column A (Ctrl+Shift+Down) and paste it after last used column ---
            colA_vals = wafermap_sheet.range((1,1),(last_row,1)).value

            # Ensure values are shaped as a column (list of lists)
            if isinstance(colA_vals, list) and not isinstance(colA_vals[0], list):
                colA_vals = [[v] for v in colA_vals]

            # Paste Column A into the new rightmost column
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).value = colA_vals
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).color = (228,241,253)
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).api.Font.Color = dark_blue
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).api.Font.Bold = True  # bold copy of Column A

            # --- Add "No." at the very last row of that new column ---
            wafermap_sheet.range((last_row+1, last_col+1)).value = "No."
            wafermap_sheet.range((last_row+1, last_col+1)).color = (228,241,253)
            wafermap_sheet.range((last_row+1, last_col+1)).api.Font.Color = dark_blue
            wafermap_sheet.range((last_row+1, last_col+1)).api.Font.Bold = True  # bold "No." cell

            # --- Also bold the original Row 1 and Column A ---
            wafermap_sheet.range((1,1),(1,last_col)).api.Font.Bold = True
            wafermap_sheet.range((1,1),(last_row,1)).api.Font.Bold = True
            
            # --- Remove gridlines from wafermap sheet ---
            wafermap_sheet.api.Parent.Windows(1).DisplayGridlines = False

            # --- Alignment (center everything including mirrored row/col) ---
            used_range = wafermap_sheet.range((1,1),(last_row+1,last_col+1))
            used_range.api.HorizontalAlignment = -4108  # xlCenter
            used_range.api.VerticalAlignment = -4108    # xlCenter

            
            # --- Borders ---
            used_range = wafermap_sheet.range((1,1),(last_row+1,last_col+1))
            used_range.api.Borders.Weight = 2

            wb_xlw.save()
            wb_xlw.close()
            app.quit()

            self.show_status(f"\n✅ Wafermap created on {sheet_name} sheet.")

            # --- Reopen workbook to safely delete pivot sheet ---
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)

            try:
                pivot_sheet = wb_xlw.sheets["Wafermap Pivot Table"]
                # Activate another sheet first
                wb_xlw.sheets[0].activate()
                pivot_sheet.delete()
                #self.show_status("\n🗑️ Wafermap Pivot Table sheet deleted after reopen.")
            except Exception as e:
                #self.show_status(f"\n⚠️ Could not delete Wafermap Pivot Table: {e}", color="#d32f2f")
                pass
            
            wb_xlw.save()
            wb_xlw.close()
            app.quit()
            
        except Exception as e:
            self.show_status(f"\n❌ Error generating wafermap: {e}", color="#d32f2f")

        finally:
            if wb_xlw:
                try: wb_xlw.close()
                except: pass
            if app:
                try: app.quit()
                except: pass

                                
    def clear_all(self):
        # Reset file path
        self.path_var.set("")

        # Clear status box
        self.show_status("", clear=True)

        # Reset combobox selection and values
        if hasattr(self, "filter_dropdown"):
            self.filter_var.set("")                 # clear current selection
            self.filter_dropdown['values'] = []     # empty the dropdown list

# --- Run the App ---
if __name__ == "__main__":
    root = tk.Tk()
    app = AutomatingDeliverables(root)
    root.mainloop()
