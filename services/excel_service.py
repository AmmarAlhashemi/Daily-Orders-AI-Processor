import os
import re
from datetime import datetime
import pandas as pd
from tkinter import messagebox
from config import ENTITY_A_AC_DIR, ENTITY_B_AC_DIR, MASTER_AC_DIR
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from utils.helpers import open_in_explorer

class ExcelExportService:
    @staticmethod
    def export_orders(processed_data: list, export_option: str) -> str:
        df = pd.DataFrame(
            processed_data,
            columns=["serial", "Recipient Name", "Account Number", "Amount", "Remark", "Designated Person", "Warning"]
        )
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0).round(2)

        now = datetime.now()
        now_date = now.strftime("%d-%m-%Y")

        prefix_map = {
            "ENTITY_A": ("ENTITY_A DO", ENTITY_A_AC_DIR),
            "MASTER": ("MASTER", MASTER_AC_DIR),
            "ENTITY_B": ("ENTITY_B DO", ENTITY_B_AC_DIR)
        }
        file_prefix, output_dir = prefix_map.get(export_option, ("MASTER", MASTER_AC_DIR))

        os.makedirs(output_dir, exist_ok=True)
        
        output_file = os.path.join(output_dir, f"{file_prefix} {now_date}.xlsx")

        if os.path.exists(output_file):
            user_choice = messagebox.askyesnocancel(
                "File Exists",
                f"The file '{os.path.basename(output_file)}' already exists.\n\n"
                "Click **Yes** to replace the file.\n\n"
                "Click **No** to rename the new file.",
                icon="warning"
            )

            if user_choice is None: 
                raise InterruptedError("Export cancelled by user.")

            if user_choice is False:
                base, ext = os.path.splitext(output_file)
                counter = 1
                while os.path.exists(output_file):
                    output_file = f"{base}__{counter}{ext}"
                    counter += 1

        wb = Workbook()
        ws = wb.active
        ws.title = "Daily Orders"

        # Styles
        title_font = Font(name="Bell MT", size=20, bold=True, color="00B050")
        header_font = Font(name="Bell MT", size=12, bold=True)
        order_no_font = Font(name="Calibri", size=14)
        date_font = Font(name="Calibri", size=12)
        alignment_center_bottom = Alignment(horizontal="center", vertical="bottom")
        alignment_center = Alignment(horizontal="center", vertical="center")
        alignment_right = Alignment(horizontal="right", vertical="bottom")
        alignment_left = Alignment(horizontal="left", vertical="bottom")
        currency_format = "#,##0.00"
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        thick_border = Border(left=Side(style="medium"), right=Side(style="medium"), top=Side(style="medium"), bottom=Side(style="medium"))

        # Title
        ws.merge_cells("A2:D2")
        ws["A2"] = "STATEMENT OF DAILY ORDERS"
        ws["A2"].font = title_font
        ws["A2"].alignment = alignment_left

        # Headers
        headers = ["DATE", "ORDER", "DESCRIPTION", "ACCOUNT", "ORDERS", "REMARK"]
        for col_num, header in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col_num, value=header)
            cell.font = header_font
            cell.alignment = alignment_center
            cell.border = Border(left=Side(style="medium"), right=Side(style="medium"), top=Side(style="medium"), bottom=Side(style=None))
        
        h2 = ["", "NO.", "", "NO.", "", ""]
        for col_num, header in enumerate(h2, start=1):
            ws.cell(row=5, column=col_num).border = thin_border
            cell = ws.cell(row=4, column=col_num, value=header)
            cell.font = header_font
            cell.alignment = alignment_center_bottom
            cell.border = Border(left=Side(style="medium"), right=Side(style="medium"), top=Side(style=None), bottom=Side(style="medium"))

        # Date
        ws["A6"] = now.strftime("%d/%m/%Y")
        ws["A6"].font = date_font
        ws["A6"].alignment = alignment_center
        ws["A6"].border = thin_border

        seen_accounts = set()

        for row_idx, row in enumerate(df.itertuples(index=False), start=6):
            ai_warn_val = row[6]
            if ai_warn_val:
                cell_h = ws.cell(row=row_idx, column=8, value=ai_warn_val)
                cell_h.font = Font(name="Calibri", size=11, color="0070C0", bold=True, italic=True)

            for col_idx, value in enumerate(row[0:6], start=2):
                cell = ws.cell(row=row_idx, column=col_idx)

                if col_idx == 5:
                    try:
                        cell.value = float(value)
                    except (ValueError, TypeError):
                        cell.value = 0.00
                    cell.number_format = currency_format
                else:
                    cell.value = value

                if col_idx == 4: # Duplicate account check
                    if value in seen_accounts:
                        existing_warn = ws.cell(row=row_idx, column=8).value
                        combined_warn = f"Duplicated Acc | {existing_warn}" if existing_warn else "Duplicated Acc"
                        ws.cell(row=row_idx, column=8, value=combined_warn)
                        ws.cell(row=row_idx, column=8).font = Font(color="FF0000", bold=True)
                        for c in range(1, 7):
                            ws.cell(row=row_idx, column=c).fill = PatternFill(start_color="FF7979", end_color="FF7979", fill_type="solid")
                    else:
                        if value:
                            seen_accounts.add(value)

                if col_idx in {3, 6}:
                    cell.value = str(value).upper()
                    if col_idx == 6:
                        cell.font = Font(bold=True)
                        if re.search(r"URGENT", cell.value, re.IGNORECASE):
                            cell.font = Font(bold=True, color="FF0000")

                if col_idx in {4, 5}:
                    cell.alignment = alignment_right
                elif col_idx == 3:
                    cell.alignment = alignment_left
                else:
                    cell.alignment = alignment_center_bottom

                if col_idx == 2:
                    cell.font = order_no_font
                    if row_idx < 15:
                        cell.value = str(f"0{row_idx - 5}")

                cell.border = thin_border if col_idx != 7 else Border()

        # Summary Row
        sum_cell = ws.cell(row=len(processed_data) + 7, column=5)
        sum_cell.value = f"=SUM(E6:E{len(processed_data) + 5})"
        sum_cell.font = Font(name="Calibri", size=12, bold=True)
        sum_cell.number_format = currency_format
        sum_cell.alignment = alignment_right
        sum_cell.border = thick_border

        for i in range(1, 7):
            ws.cell(row=len(processed_data) + 6, column=i).border = thin_border
            ws.cell(row=len(processed_data) + 7, column=i).border = thick_border

        for i in range(5, len(processed_data) + 7):
            ws[f"A{i}"].border = Border(left=Side(style="medium"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
            ws[f"F{i}"].border = Border(left=Side(style="thin"), right=Side(style="medium"), top=Side(style="thin"), bottom=Side(style="thin"))

        ws.column_dimensions["A"].width = 12.7
        ws.column_dimensions["B"].width = 9.13
        ws.column_dimensions["C"].width = 38.13
        ws.column_dimensions["D"].width = 20.7
        ws.column_dimensions["E"].width = 14.7
        ws.column_dimensions["F"].width = 14.13
        ws.column_dimensions["H"].width = 35.0

        wb.save(output_file)
        open_in_explorer(output_file)
        return output_file