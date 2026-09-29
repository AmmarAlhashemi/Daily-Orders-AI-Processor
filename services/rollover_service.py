import os
import shutil
from datetime import timedelta
import win32com.client as win32
from utils.helpers import hex_to_excel_color
from config import ENTITY_A_AC_DIR, ENTITY_B_AC_DIR, MASTER_AC_DIR

def perform_monthly_rollover_and_archive(today):
    """Executes file archiving and ledger rollover within local data directory."""
    os.system("")

    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    RESET = "\033[0m"

    first_day_this_month = today.replace(day=1)
    last_day_prev_month = first_day_this_month - timedelta(days=1)

    prev_month_upper = last_day_prev_month.strftime("%b").upper()
    prev_month_title = last_day_prev_month.strftime("%b").title().upper()
    prev_year = last_day_prev_month.strftime("%Y")
    prev_month_date_str = last_day_prev_month.strftime("%m-%Y")

    curr_month_upper = today.strftime("%b").upper()
    curr_year = today.strftime("%Y")

    print("\n", f"{CYAN}━" * 65 + RESET)
    print(f"{BOLD}{CYAN}  WE ARE IN {curr_month_upper}! STARTING AUTO ARCHIVE FILES FOR {prev_month_upper}{RESET}")
    print(f"{CYAN}━" * 65 + RESET)

    # PART 1: File Archiving using local paths
    print(f"\n{BOLD}{CYAN}[1/2] Archiving Previous Month Files...{RESET}")

    archive_targets = [
        {"prefix": "ENTITY_A DO", "base_dir": ENTITY_A_AC_DIR, "format": "%b DO-%Y"},
        {"prefix": "MASTER", "base_dir": MASTER_AC_DIR, "format": "%b-%Y"},
        {"prefix": "ENTITY_B DO", "base_dir": ENTITY_B_AC_DIR, "format": "%b DO-%Y"},
    ]

    archive_summary = {}

    for target in archive_targets:
        file_prefix = target["prefix"]
        archive_base_dir = target["base_dir"]
        folder_type = os.path.basename(archive_base_dir)

        prev_month_folder_name = last_day_prev_month.strftime(target["format"]).upper()
        archived_count = 0

        if os.path.exists(archive_base_dir):
            target_month_dir = os.path.join(archive_base_dir, prev_month_folder_name)
            os.makedirs(target_month_dir, exist_ok=True)

            items = os.listdir(archive_base_dir)
            total_items = len(items)

            for idx, item in enumerate(items, 1):
                item_path = os.path.join(archive_base_dir, item)
                print(f"\r  {CYAN}⠋{RESET} Processing {folder_type} ({idx}/{total_items})... {DIM}{item[:30]}...{RESET}", end="", flush=True)

                if (
                    os.path.isfile(item_path)
                    and item.startswith(file_prefix)
                    and item.endswith(".xlsx")
                    and f"-{prev_month_date_str}" in item
                ):
                    dest_path = os.path.join(target_month_dir, item)
                    try:
                        shutil.copy2(item_path, dest_path)
                        if os.path.exists(dest_path) and os.path.getsize(dest_path) == os.path.getsize(item_path):
                            os.remove(item_path)
                            archived_count += 1
                    except Exception:
                        pass

            print(f"\r\033[K  {GREEN}•{RESET} {folder_type:<15} : Moved {BOLD}{archived_count}{RESET} files to {DIM}{prev_month_folder_name}{RESET}")
        else:
            print(f"  {YELLOW}•{RESET} {folder_type:<15} : Folder not found")

        archive_summary[folder_type] = archived_count

    # PART 2: Excel Ledger Rollover
    print(f"\n{BOLD}{CYAN}[2/2] Rolling Over Active Excel Templates...{RESET}")

    excel_accounts = [
        {"base_path": ENTITY_A_AC_DIR, "prefix": "ENTITY_A AC"},
        {"base_path": ENTITY_B_AC_DIR, "prefix": "ENTITY_B AC"}
    ]

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    rollover_summary = {}

    try:
        for acc in excel_accounts:
            base_path = str(acc["base_path"])
            prefix = acc["prefix"]

            print(f"\r  {CYAN}⠋{RESET} Opening Excel Engine for {prefix}...", end="", flush=True)

            prev_filename = f"{prefix} {prev_month_upper}-{prev_year}.xlsx"
            prev_filepath = os.path.join(base_path, prev_filename)

            curr_filename = f"{prefix} {curr_month_upper}-{curr_year}.xlsx"
            curr_filepath = os.path.join(base_path, curr_filename)

            prev_folder = f"{prev_month_title} DO-{prev_year}"
            prev_folder_path = os.path.join(base_path, prev_folder)

            if not os.path.exists(prev_filepath):
                print(f"\r\033[K  {YELLOW}•{RESET} {prefix:<15} : Missing source ({prev_month_upper}-{prev_year}.xlsx)")
                rollover_summary[prefix] = "Skipped (Not Found)"
                continue

            wb = None
            try:
                wb = excel.Workbooks.Open(prev_filepath)
                ws = wb.Sheets(1)

                last_row_a = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                date_rows = []
                for r in range(1, last_row_a + 1):
                    val = ws.Cells(r, 1).Value
                    if val and str(val).strip().upper() == "DATE":
                        date_rows.append(r)
                    if len(date_rows) == 2:
                        break

                if len(date_rows) >= 2:
                    start_del = date_rows[0]
                    end_del = date_rows[1] - 1
                    ws.Range(f"{start_del}:{end_del}").EntireRow.Delete()

                max_r = max(ws.Cells(ws.Rows.Count, col).End(-4162).Row for col in range(1, 7))
                ws.Cells(max_r, 1).Select()
                excel.ActiveWindow.ScrollRow = max_r

                wb.Save()

                if not os.path.exists(prev_folder_path):
                    os.makedirs(prev_folder_path)

                copy_dest = os.path.join(prev_folder_path, prev_filename)
                wb.SaveCopyAs(copy_dest)
                wb.Close(SaveChanges=True)
                wb = None

                os.rename(prev_filepath, curr_filepath)

                wb = excel.Workbooks.Open(curr_filepath)
                ws = wb.Sheets(1)

                ws.Cells(4, 1).Value = f"MONTH:{curr_month_upper}-{curr_year}"

                maxrow_f = ws.Cells(ws.Rows.Count, 6).End(-4162).Row
                temp_variable = ws.Cells(maxrow_f, 6).Value

                last_row_a = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                first_date_row = None
                for r in range(1, last_row_a + 1):
                    val = ws.Cells(r, 1).Value
                    if val and str(val).strip().upper() == "DATE":
                        first_date_row = r
                        break

                end_copy_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
                old_maxrow = max(
                    ws.Cells(ws.Rows.Count, 1).End(-4162).Row,
                    ws.Cells(ws.Rows.Count, 6).End(-4162).Row
                )

                insert_start_row = old_maxrow + 4
                ws.Range(f"{first_date_row}:{end_copy_row}").Copy()
                ws.Range(f"{insert_start_row}:{insert_start_row}").Insert(Shift=-4121)
                excel.Application.CutCopyMode = False

                block_size = end_copy_row - first_date_row + 1
                inserted_row_start = insert_start_row
                new_maxrow = insert_start_row + block_size - 1

                del_start = inserted_row_start + 3
                del_end = new_maxrow - 4
                if del_end >= del_start:
                    ws.Range(f"{del_start}:{del_end}").EntireRow.Delete()

                rows_deleted = (del_end - del_start + 1) if del_end >= del_start else 0
                current_maxrow = new_maxrow - rows_deleted

                ob_count = 0
                target_ob_row = None
                for r in range(1, current_maxrow + 1):
                    val_c = ws.Cells(r, 3).Value
                    if val_c and "OPENING BALANCE B/F" in str(val_c).upper():
                        ob_count += 1
                        if ob_count == 2:
                            target_ob_row = r
                            break

                if target_ob_row:
                    ws.Cells(target_ob_row, 6).Value = temp_variable

                ws.Range(f"A{first_date_row}:F{end_copy_row}").Interior.Color = hex_to_excel_color("#F4B084" if acc["prefix"] == "ENTITY_A AC" else "#FFD966")

                max_r = max(ws.Cells(ws.Rows.Count, col).End(-4162).Row for col in range(1, 7))
                ws.Range(f"A{max_r - 2}:E{max_r - 3}").ClearContents()

                ws.Cells(max_r, 1).Select()
                excel.ActiveWindow.ScrollRow = max_r

                wb.Save()
                print(f"\r\033[K  {GREEN}•{RESET} {prefix:<15} : Rolled over to {BOLD}{curr_month_upper}-{curr_year}{RESET}")
                rollover_summary[prefix] = "Completed"

            except Exception as e:
                print(f"\r\033[K  {RED}•{RESET} {prefix:<15} : Failed ({e})")
                rollover_summary[prefix] = "Failed"
            finally:
                if wb:
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        pass
    finally:
        excel.Quit()

    print("\n", f"{CYAN}━" * 65 + RESET)
    print(f" {BOLD}{GREEN}* MONTHLY SEQUENCE COMPLETE *{RESET}")
    print(f"{CYAN}━" * 65 + RESET)
    print(f"  • {BOLD}Total Files Archived{RESET} : {sum(archive_summary.values())} files")
    print(f"  • {BOLD}Active Accounts Rolled{RESET}: {len([v for v in rollover_summary.values() if v == 'Completed'])} of {len(excel_accounts)}")
    print(f"{CYAN}━" * 65 + RESET + "\n")