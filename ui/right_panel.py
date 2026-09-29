import os
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox
import customtkinter as ctk
from config import (
    MASTER_AC_DIR, ACCENT_CYAN, ACCENT_PURPLE, ACCENT_PURPLE_H,
    COLOR_SUCCESS, COLOR_SUCCESS_H, COLOR_DANGER,
    BG_FRAME, BORDER_COLOR, TEXT_WHITE, TEXT_MUTED
)

class RightPanel(ctk.CTkFrame):
    def __init__(self, parent, on_export_callback, on_divide_callback):
        super().__init__(
            parent,
            width=310,
            corner_radius=18,
            fg_color=BG_FRAME,
            border_width=1.5,
            border_color=BORDER_COLOR
        )
        self.pack_propagate(False)

        self.on_export_callback = on_export_callback
        self.on_divide_callback = on_divide_callback

        # Export Section
        self.export_label = ctk.CTkLabel(self, text="⚙️ EXPORT DATA", font=("Segoe UI", 13, "bold"), text_color=ACCENT_CYAN)
        self.export_label.pack(anchor=ctk.W, pady=(20, 15), padx=20)

        self.export_option = ctk.StringVar(value="MASTER")

        self.radio_master = ctk.CTkRadioButton(
            self, text="MASTER", variable=self.export_option, value="MASTER",
            font=("Segoe UI", 12, "bold"), fg_color=ACCENT_CYAN, hover_color=ACCENT_CYAN, text_color=TEXT_WHITE
        )
        self.radio_master.pack(anchor=ctk.W, pady=8, padx=20)

        self.radio_A = ctk.CTkRadioButton(
            self, text="ENTITY_A", variable=self.export_option, value="A",
            font=("Segoe UI", 12), fg_color=ACCENT_CYAN, hover_color=ACCENT_CYAN, text_color=TEXT_MUTED
        )
        self.radio_A.pack(anchor=ctk.W, pady=8, padx=20)

        self.radio_B = ctk.CTkRadioButton(
            self, text="ENTITY_B", variable=self.export_option, value="B",
            font=("Segoe UI", 12), fg_color=ACCENT_CYAN, hover_color=ACCENT_CYAN, text_color=TEXT_MUTED
        )
        self.radio_B.pack(anchor=ctk.W, pady=8, padx=20)

        self.btn_container = tk.Frame(self, bg=BG_FRAME, width=300, height=80)
        self.btn_container.pack_propagate(False)
        self.btn_container.pack()

        self.custom_export_button = ctk.CTkButton(
            self.btn_container,
            text="EXPORT TO EXCEL",
            command=self.start_export_process,
            font=("Segoe UI", 13, "bold"),
            corner_radius=10,
            height=42,
            fg_color=COLOR_SUCCESS,
            hover_color=COLOR_SUCCESS_H,
            cursor="hand2",
            text_color=TEXT_WHITE
        )
        self.custom_export_button.pack(fill=ctk.X, pady=(20, 20), padx=20)

        self.processing_label = tk.Label(
            self.btn_container,
            text="",
            font=("Segoe UI", 11, "bold", "italic"),
            fg=ACCENT_CYAN,
            bg=BG_FRAME
        )

        self.ai_progress = ctk.CTkProgressBar(
            self,
            orientation="horizontal",
            mode="indeterminate",
            progress_color=ACCENT_CYAN,
            fg_color="#1E2942",
            height=6
        )

        self.separator = ctk.CTkFrame(self, height=1.5, fg_color=BORDER_COLOR)
        self.separator.pack(fill=ctk.X, pady=10, padx=20)

        # File Divider Section
        self.div_label = ctk.CTkLabel(self, text="🧩 FILE DIVIDER", font=("Segoe UI", 13, "bold"), text_color=ACCENT_CYAN)
        self.div_label.pack(anchor=ctk.W, pady=(15, 12), padx=20)

        self.div_file_path = ctk.StringVar()

        self.choose_button = ctk.CTkButton(
            self,
            text="📂 Select Excel File",
            command=self.select_file,
            font=("Segoe UI", 12, "bold"),
            corner_radius=10,
            height=36,
            fg_color="#1F2937",
            hover_color="#374151",
            text_color=TEXT_WHITE,
            border_width=1,
            border_color="#4B5563"
        )
        self.choose_button.pack(anchor=ctk.W, fill=ctk.X, pady=(0, 10), padx=20)

        self.file_label = ctk.CTkLabel(
            self, text="No file selected", font=("Segoe UI", 12, "bold"), text_color=TEXT_MUTED, wraplength=260, justify="left"
        )
        self.file_label.pack(anchor=ctk.W, pady=(0, 10), padx=20)

        self.divide_option = ctk.StringVar(value="obo")

        self.odd_button = ctk.CTkRadioButton(
            self, text="One By One Split", variable=self.divide_option, value="obo",
            font=("Segoe UI", 12), fg_color=ACCENT_CYAN, text_color=TEXT_WHITE
        )

        self.even_button = ctk.CTkRadioButton(
            self, text="Bank Difference Split", variable=self.divide_option, value="bank",
            font=("Segoe UI", 12), fg_color=ACCENT_CYAN, text_color=TEXT_WHITE
        )

        self.ai_smart_button = ctk.CTkRadioButton(
            self, text="Smart Split with AI (Gemini)", variable=self.divide_option, value="ai_smart",
            font=("Segoe UI", 12), fg_color=ACCENT_CYAN, text_color=TEXT_WHITE
        )

        self.btn_container_2 = tk.Frame(self, bg=BG_FRAME, width=300, height=80)
        self.btn_container_2.pack_propagate(False)

        self.divide_button = ctk.CTkButton(
            self.btn_container_2,
            text="DIVIDE FILE",
            command=self.start_divide_process,
            font=("Segoe UI", 13, "bold"),
            corner_radius=10,
            height=42,
            fg_color=ACCENT_PURPLE,
            hover_color=ACCENT_PURPLE_H,
            text_color=TEXT_WHITE
        )

        self.processing_label_2 = tk.Label(
            self.btn_container_2,
            text="",
            font=("Segoe UI", 11, "bold", "italic"),
            fg=ACCENT_CYAN,
            bg=BG_FRAME
        )

        self.ai_progress_2 = ctk.CTkProgressBar(
            self,
            orientation="horizontal",
            mode="indeterminate",
            progress_color=ACCENT_CYAN,
            fg_color="#1E2942",
            height=6
        )

    def select_file(self):
        selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")], 
                        initialdir=str(MASTER_AC_DIR),
                        initialfile=f"MASTER {datetime.now().strftime("%d-%m-%Y")}.xlsx"
                    )
        if selected:
            self.div_file_path.set(selected)
            self.file_label.configure(text=f"📂 {os.path.basename(selected)}", text_color=ACCENT_CYAN)
            self.file_label.pack(anchor="w", pady=(5, 10), padx=15)
            self.odd_button.pack(anchor="w", pady=6, padx=15)
            self.even_button.pack(anchor="w", pady=6, padx=15)
            self.ai_smart_button.pack(anchor="w", pady=6, padx=15)
            self.btn_container_2.pack()
            self.divide_button.pack(fill=ctk.X, pady=(20, 0), padx=20)
        else:
            self.file_label.configure(text="❌ No file chosen", text_color=COLOR_DANGER)

    def start_export_process(self):
        self.show_progress("export", "Preparing connection...")
        threading.Thread(
            target=lambda: self.on_export_callback(
                self.export_option.get(),
                self.update_status,
                self.reset_ui
            ),
            daemon=True
        ).start()

    def start_divide_process(self):
        self.show_progress("div", "Dividing data...")
        threading.Thread(
            target=lambda: self.on_divide_callback(
                self.div_file_path.get(),
                self.divide_option.get(),
                self.update_status,
                self.reset_ui
            ),
            daemon=True
        ).start()

    def show_progress(self, section: str, msg: str):
        if section == "export":
            self.custom_export_button.pack_forget()
            self.processing_label.configure(text=msg)
            self.processing_label.pack(expand=True)
            self.ai_progress.pack(fill=ctk.X, pady=(0, 15), padx=20, after=self.processing_label)
            self.ai_progress.start()
        else:
            self.divide_button.pack_forget()
            self.processing_label_2.configure(text=msg)
            self.processing_label_2.pack(expand=True)
            self.ai_progress_2.pack(fill=ctk.X, pady=(0, 15), padx=20, after=self.processing_label_2)
            self.ai_progress_2.start()

    def update_status(self, section: str, msg: str):
        if section == "export":
            self.processing_label.configure(text=msg)
        else:
            self.processing_label_2.configure(text=msg)

    def reset_ui(self, section: str):
        if section == "export":
            self.processing_label.pack_forget()
            self.custom_export_button.pack(fill=ctk.X, pady=(20, 20), padx=20)
            self.ai_progress.stop()
            self.ai_progress.pack_forget()
        else:
            self.processing_label_2.pack_forget()
            self.divide_button.pack(fill=ctk.X, pady=(20, 0), padx=20)
            self.ai_progress_2.stop()
            self.ai_progress_2.pack_forget()