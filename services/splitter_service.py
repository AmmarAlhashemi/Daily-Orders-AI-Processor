import os
import re
import shutil
from tkinter import messagebox
from datetime import datetime
from config import ENTITY_A_AC_DIR, ENTITY_B_AC_DIR
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

class FileSplitterService:
    @staticmethod
    def divide_file(file_path: str, option: str, smart_split_callback=None) -> tuple[str, str]:
        if not file_path:
            raise ValueError("No file path provided.")

        date_match = re.search(r"\d{2}-\d{2}-\d{4}", file_path)
        if not date_match:
            raise ValueError("Could not find a valid date (DD-MM-YYYY) in the filename.")

        date_out = date_match.group()

        dest_path1 = os.path.join(ENTITY_A_AC_DIR, f"ENTITY_A DO {date_out}.xlsx")
        dest_path2 = os.path.join(ENTITY_B_AC_DIR, f"ENTITY_B DO {date_out}.xlsx")

        for path_var in [dest_path1, dest_path2]:
            if os.path.exists(path_var):
                user_choice = messagebox.askyesnocancel(
                    "File Exists",
                    f"The file '{os.path.basename(path_var)}' already exists.\n\n"
                    "Click **Yes** to replace the file.\n\n"
                    "Click **No** to rename the new file.",
                    icon="warning"
                )

                if user_choice is None:  # Cancel
                    raise InterruptedError("Division cancelled by user.")

                if user_choice is False:  # Auto-rename
                    base, ext = os.path.splitext(path_var)
                    counter = 1
                    while os.path.exists(path_var):
                        path_var = f"{base}__{counter}{ext}"
                        counter += 1
                
                if "A" in os.path.basename(path_var):
                    dest_path1 = path_var
                else:
                    dest_path2 = path_var

        wb_src = load_workbook(file_path)
        ws_src = wb_src.active

        rows_ws1 = []
        rows_ws2 = []

        if option == "obo": # One By One Split
            i = 6
            while ws_src[f"B{i}"].value is not None:
                serialno = float(ws_src[f"B{i}"].value)
                if (serialno - 6) % 2 == 1:
                    rows_ws2.append(i)
                else:
                    rows_ws1.append(i)
                i += 1

        elif option == "bank": # Bank Difference Split
            i = 6
            while ws_src[f"D{i}"].value is not None:
                cell_value = ws_src[f"D{i}"].value
                try:
                    if ws_src[f"F{i}"].value is None:
                        rows_ws2.append(i)
                    elif ws_src[f"F{i}"].value is not None:
                        rows_ws1.append(i)
                    else:
                        if len(str(cell_value)) in [12, 13, 14] and str(cell_value).startswith("1"):
                            rows_ws2.append(i)
                        else:
                            rows_ws1.append(i)
                except Exception as e:
                    print(f"Error processing row {i}: {e}")
                i += 1

        elif option == "ai_smart" and smart_split_callback:
            i = 6
            excel_data_for_ai = []
            serial_to_row_map = {}

            while ws_src[f"B{i}"].value is not None:
                serial_str = str(ws_src[f"B{i}"].value).strip()
                amt_val = ws_src[f"E{i}"].value
                rem_val = ws_src[f"F{i}"].value

                try:
                    amt = float(amt_val) if amt_val is not None else 0.0
                except:
                    amt = 0.0

                remark_str = str(rem_val or "").upper().strip()

                excel_data_for_ai.append({
                    "serial": serial_str,
                    "amount": amt,
                    "remark": remark_str
                })
                serial_to_row_map[serial_str] = i
                i += 1

            ai_result = smart_split_callback(excel_data_for_ai)

            for s_id in ai_result.get("delete_from_ENTITY_A", []):
                if str(s_id).strip() in serial_to_row_map:
                    rows_ws1.append(serial_to_row_map[str(s_id).strip()])

            for s_id in ai_result.get("delete_from_ENTITY_B", []):
                if str(s_id).strip() in serial_to_row_map:
                    rows_ws2.append(serial_to_row_map[str(s_id).strip()])

        # Execute File Generation
        shutil.copy(file_path, dest_path1)
        shutil.copy(file_path, dest_path2)

        wb1 = load_workbook(dest_path1)
        ws1 = wb1.active
        wb2 = load_workbook(dest_path2)
        ws2 = wb2.active

        ws1["A2"].font = Font(name="Bell MT", size=20, bold=True, color="00B0F0")
        ws2["A2"].font = Font(name="Bell MT", size=20, bold=True, color="7030A0")

        rows_ws1.sort()
        rows_ws2.sort()

        for row_idx in reversed(rows_ws1):
            ws1.delete_rows(row_idx)
        for row_idx in reversed(rows_ws2):
            ws2.delete_rows(row_idx)

        for ws in [ws1, ws2]:
            idx = 6
            while ws[f"B{idx}"].value is not None:
                try:
                    ws[f"B{idx}"].value = f"0{idx-5}" if idx-5 < 10 else idx-5
                except Exception as e:
                    print(e)
                idx += 1
            ws[f"E{idx+1}"].value = f"=SUM(E6:E{idx-1})"
            ws["A6"].font = Font(name="Calibri", size=12)
            ws["A6"].alignment = Alignment(horizontal="center", vertical="center")
            ws["A6"].value = date_out.replace("-", "/")

        wb1.save(dest_path1)
        wb2.save(dest_path2)

        return dest_path1, dest_path2