import os
import tkinter as tk
from tkinter import messagebox
import threading
import customtkinter as ctk
from config import BG_MAIN
from ui.left_panel import LeftPanel
from ui.right_panel import RightPanel
from services.api_client import APIClient
from services.excel_service import ExcelExportService
from services.splitter_service import FileSplitterService

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Daily Orders AI Processor")
        self.geometry("950x670+65+0")
        self.configure(fg_color=BG_MAIN)

        self.api_client = APIClient()

        # Layout
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(padx=25, pady=25, fill=ctk.BOTH, expand=True)

        self.left_panel = LeftPanel(self.main_frame)
        self.left_panel.pack(side=ctk.LEFT, padx=(0, 20), fill=ctk.BOTH, expand=True)

        self.right_panel = RightPanel(
            self.main_frame,
            on_export_callback=self.handle_export,
            on_divide_callback=self.handle_divide
        )
        self.right_panel.pack(side=ctk.RIGHT, fill=ctk.Y, padx=(10, 0))

    def handle_export(self, export_option: str, update_status_cb, reset_ui_cb):
        raw_data = self.left_panel.get_text()

        if not raw_data:
            self.after(0, lambda: reset_ui_cb("export"))
            messagebox.showerror("Validation Error", "No raw data provided!")
            return

        try:
            update_status_cb("export", "Processing with Gemini AI...")

            try:
                response = self.api_client.parse_orders(
                    raw_data,
                    use_fallback=False
                )
            except Exception as ex:
                if str(ex) != "PRIMARY_MODEL_FAILED":
                    raise

                update_status_cb("export", "Waiting for user decision...⏳")

                user_choice = messagebox.askyesno(
                    "Model Unavailable",
                    "Gemini 3.6 Flash is currently unavailable.\n\n"
                    "Would you like to try with Gemini 3.5 Flash?"
                )

                if not user_choice:
                    return

                update_status_cb("export", "Processing with Gemini 3.5... 🔄")

                response = self.api_client.parse_orders(
                    raw_data,
                    use_fallback=True
                )

            transactions = response.get("transactions", [])
            processed_data = []

            for transaction in transactions:
                processed_data.append([
                    transaction.get("serial", ""),
                    transaction.get("recipient_name", ""),
                    transaction.get("account_number", ""),
                    transaction.get("amount", 0.0),
                    transaction.get("remark", ""),
                    transaction.get("designated_person", ""),
                    transaction.get("ai_warning", "")
                ])

            if not processed_data:
                messagebox.showerror(
                    "Processing Error",
                    "No valid transaction data was extracted."
                )
                return

            update_status_cb("export", "Generating Excel File...")

            output_file = ExcelExportService.export_orders(
                processed_data,
                export_option
            )
            self.after(0, lambda: reset_ui_cb("export"))

            messagebox.showinfo(
                "Success",
                f"Data exported successfully:\n{output_file}"
            )

        except Exception as ex:
            messagebox.showerror(
                "Export Error",
                f"An error occurred during export:\n{ex}"
            )

        finally:
            self.after(0, lambda: reset_ui_cb("export"))

    def handle_divide(self, file_path: str, divide_option: str, update_status_cb, reset_ui_cb):
        if not file_path:
            self.after(0, lambda: reset_ui_cb("div"))
            messagebox.showerror(
                "Validation Error", "Please select an Excel file first."
            )
            return

        def smart_split_callback(excel_data):
            update_status_cb("div", "Processing with Gemini AI...")
            try:
                return self.api_client.smart_split(
                    excel_data, use_fallback=False
                )
            except Exception as ex:
                if str(ex) != "PRIMARY_MODEL_FAILED":
                    raise

                update_status_cb("div", "Waiting for user decision...⏳")

                event = threading.Event()
                user_choice = [False]

                def ask_on_main_thread():
                    user_choice[0] = messagebox.askyesno(
                        "Model Unavailable",
                        "Gemini 3.6 Flash is currently unavailable.\n\n"
                        "Would you like to try splitting with Gemini 3.5 Flash?",
                    )
                    event.set()

                self.after(0, ask_on_main_thread)
                event.wait()

                if not user_choice[0]:
                    raise InterruptedError("Division cancelled by user.")

                update_status_cb("div", "Processing with Gemini 3.5... 🔄")

                return self.api_client.smart_split(
                    excel_data, use_fallback=True
                )

        try:
            update_status_cb("div", "Preparing file...")

            p1, p2 = FileSplitterService.divide_file(
                file_path=file_path,
                option=divide_option,
                smart_split_callback=(
                    smart_split_callback if divide_option == "ai_smart" else None
                ),
            )

            self.after(0, lambda: reset_ui_cb("div"))

            messagebox.showinfo(
                "Success",
                f"Files created successfully inside local 'data' folder:\n\n"
                f"1. {os.path.basename(p1)}\n"
                f"2. {os.path.basename(p2)}",
            )

        except InterruptedError:
            pass

        except Exception as ex:
            messagebox.showerror(
                "Division Error", f"Failed to split file:\n{ex}"
            )

        finally:
            self.after(0, lambda: reset_ui_cb("div"))