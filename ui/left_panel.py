import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from config import (
    ACCENT_CYAN, ACCENT_PURPLE, ACCENT_PURPLE_H,
    COLOR_DANGER, COLOR_DANGER_H, BG_FRAME, BORDER_COLOR, TEXT_WHITE
)

class LeftPanel(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")

        self.text_label = ctk.CTkLabel(
            self,
            text="✨ Raw Data Input :",
            font=("Segoe UI", 15, "bold"),
            text_color=ACCENT_CYAN
        )
        self.text_label.pack(anchor=ctk.W, pady=(0, 12))

        self.text_field = ctk.CTkTextbox(
            self,
            font=("Calibri", 15),
            fg_color=BG_FRAME,
            text_color=TEXT_WHITE,
            corner_radius=14,
            border_width=1.5,
            border_color=BORDER_COLOR
        )
        self.text_field.pack(fill=ctk.BOTH, expand=True, pady=(0, 20))
        self.text_field._textbox.configure(insertbackground=ACCENT_CYAN)

        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(fill=ctk.X)

        self.paste_button = ctk.CTkButton(
            self.button_frame,
            text="Paste Data",
            command=self.paste_data,
            font=("Segoe UI", 13, "bold"),
            corner_radius=10,
            height=40,
            fg_color=ACCENT_PURPLE,
            hover_color=ACCENT_PURPLE_H,
            text_color=TEXT_WHITE
        )
        self.paste_button.pack(side=ctk.LEFT, padx=(0, 12), expand=True, fill=ctk.X)

        self.clear_button = ctk.CTkButton(
            self.button_frame,
            text="Clear Data",
            command=self.clear_data,
            font=("Segoe UI", 13, "bold"),
            corner_radius=10,
            height=40,
            fg_color=COLOR_DANGER,
            hover_color=COLOR_DANGER_H,
            text_color=TEXT_WHITE
        )
        self.clear_button.pack(side=ctk.LEFT, expand=True, fill=ctk.X)

    def paste_data(self):
        try:
            clipboard_text = self.clipboard_get()
            self.text_field.insert("1.0", clipboard_text)
        except Exception as e:
            messagebox.showerror("Clipboard Error", f"Failed to paste data: {e}")

    def clear_data(self):
        self.text_field.delete("1.0", tk.END)

    def get_text(self) -> str:
        return self.text_field.get("1.0", tk.END).strip()