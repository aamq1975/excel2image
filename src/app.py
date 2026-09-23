"""
Excel Image Extractor
----------------------
A small Windows desktop app that lets you pick an Excel file (.xlsx / .xlsm)
and extracts every embedded picture directly into a folder you choose -
no zip file is created.

Built with Tkinter + the standard library only (no Pillow, no openpyxl)
so the packaged .exe stays as small as possible. An .xlsx file is
already a zip archive internally, with embedded pictures stored under
"xl/media/" - so extraction is just reading those entries straight out
of the archive and writing them to disk.

Developer: Muhammad Aamir
Email: aamq1975@gmail.com
WhatsApp: +92-300-2376216
"""

import os
import sys
import zipfile
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox

APP_TITLE = "Excel Image Extractor"
EXCEL_GREEN = "#217346"
EXCEL_GREEN_DARK = "#1b5e3a"
EXTRACT_GREEN = "#2f9e4f"
EXTRACT_GREEN_DARK = "#237c3d"
FOLDER_BLUE = "#2b6cb0"
FOLDER_BLUE_DARK = "#1f4e80"
DISABLED_GRAY = "#c3cbc7"
BG_COLOR = "#f2f4f3"
TEXT_DARK = "#1f2a24"
TEXT_MUTED = "#5b6b63"
LINK_COLOR = "#1a5fb4"
LINK_HOVER = "#0d3d80"

DEV_NAME = "Muhammad Aamir"
DEV_EMAIL = "aamq1975@gmail.com"
DEV_WHATSAPP_DISPLAY = "+92-300-2376216"
DEV_WHATSAPP_NUMBER = "923002376216"  # digits only, for the wa.me link


def resource_path(relative_path: str) -> str:
    """Resolve a bundled resource's path, whether running as a plain
    script or as a PyInstaller --onefile exe (which unpacks bundled
    data into a temporary _MEIPASS folder at runtime)."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def extract_images_from_excel(excel_path: str, output_folder: str) -> int:
    """Extract every embedded image from an .xlsx/.xlsm file straight
    into output_folder. Returns the number of images extracted.
    Raises ValueError if the file has no embedded images, and
    zipfile.BadZipFile if it isn't a valid Excel (OOXML) file.
    """
    with zipfile.ZipFile(excel_path) as archive:
        media_entries = sorted(
            name for name in archive.namelist()
            if name.startswith("xl/media/") and not name.endswith("/")
        )

        if not media_entries:
            raise ValueError("No embedded images were found in this Excel file.")

        os.makedirs(output_folder, exist_ok=True)

        extracted = 0
        for entry in media_entries:
            data = archive.read(entry)
            filename = os.path.basename(entry)
            dest_path = os.path.join(output_folder, filename)

            # Avoid silently overwriting same-named images from a
            # previous extraction into the same folder.
            if os.path.exists(dest_path):
                name, ext = os.path.splitext(filename)
                counter = 1
                while os.path.exists(dest_path):
                    dest_path = os.path.join(output_folder, f"{name}_{counter}{ext}")
                    counter += 1

            with open(dest_path, "wb") as out_file:
                out_file.write(data)
            extracted += 1

        return extracted


class ExcelImageExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.selected_file = ""
        self.last_output_folder = ""

        root.title(APP_TITLE)
        root.configure(bg=BG_COLOR)
        root.resizable(False, False)

        # --- Window / taskbar icon ---
        try:
            icon_img = tk.PhotoImage(file=resource_path("icon.png"))
            root.iconphoto(True, icon_img)
            self._icon_ref = icon_img  # keep a reference so it isn't garbage-collected
        except Exception:
            pass  # icon is cosmetic - never block the app over it

        self._build_ui()
        self._center_window(520, 480)

    def _center_window(self, width: int, height: int) -> None:
        self.root.update_idletasks()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = (screen_w // 2) - (width // 2)
        y = (screen_h // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(width, height)

    def _make_3d_button(self, parent, text, command, bg, active_bg,
                         fg="white", font_size=12, padx=18, pady=12):
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground="white",
            disabledforeground="#7a847e",
            font=("Segoe UI", font_size, "bold"),
            relief=tk.RAISED,
            bd=5,
            padx=padx,
            pady=pady,
            cursor="hand2",
        )
        # Simple 3D "press" feedback
        def on_press(e):
            if btn["state"] != tk.DISABLED:
                btn.configure(relief=tk.SUNKEN, bd=2)

        def on_release(e):
            if btn["state"] != tk.DISABLED:
                btn.configure(relief=tk.RAISED, bd=5)

        btn.bind("<ButtonPress-1>", on_press)
        btn.bind("<ButtonRelease-1>", on_release)
        return btn

    def _make_link_label(self, parent, text, url_command, font_size=10):
        label = tk.Label(
            parent,
            text=text,
            bg=BG_COLOR,
            fg=LINK_COLOR,
            font=("Segoe UI", font_size, "underline"),
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda e: url_command())
        label.bind("<Enter>", lambda e: label.configure(fg=LINK_HOVER))
        label.bind("<Leave>", lambda e: label.configure(fg=LINK_COLOR))
        return label

    def _build_ui(self):
        # --- Header ---
        header = tk.Frame(self.root, bg=BG_COLOR)
        header.pack(pady=(22, 8))

        tk.Label(
            header,
            text=APP_TITLE,
            bg=BG_COLOR,
            fg=TEXT_DARK,
            font=("Segoe UI", 20, "bold"),
        ).pack()

        tk.Label(
            header,
            text="Extract embedded pictures from an Excel file into a folder",
            bg=BG_COLOR,
            fg=TEXT_MUTED,
            font=("Segoe UI", 11),
            wraplength=440,
            justify="center",
        ).pack(pady=(4, 0))

        # --- Selected file display ---
        self.file_label = tk.Label(
            self.root,
            text="No Excel file selected",
            bg="white",
            fg=TEXT_MUTED,
            font=("Segoe UI", 11),
            relief=tk.SUNKEN,
            bd=2,
            anchor="w",
            padx=10,
            pady=10,
            width=42,
        )
        self.file_label.pack(pady=(18, 14), padx=24)

        # --- Buttons ---
        btn_frame = tk.Frame(self.root, bg=BG_COLOR)
        btn_frame.pack(pady=4)

        self.select_btn = self._make_3d_button(
            btn_frame, "📂  Select Excel File", self.on_select_file,
            bg=EXCEL_GREEN, active_bg=EXCEL_GREEN_DARK,
        )
        self.select_btn.grid(row=0, column=0, padx=8, pady=4)

        self.extract_btn = self._make_3d_button(
            btn_frame, "🖼  Extract Images", self.on_extract_images,
            bg=EXTRACT_GREEN, active_bg=EXTRACT_GREEN_DARK,
        )
        self.extract_btn.grid(row=0, column=1, padx=8, pady=4)
        self.extract_btn.configure(state=tk.DISABLED, bg=DISABLED_GRAY)

        self.open_folder_btn = self._make_3d_button(
            self.root, "📁  Open Extracted Folder", self.on_open_folder,
            bg=FOLDER_BLUE, active_bg=FOLDER_BLUE_DARK, font_size=11,
        )
        self.open_folder_btn.pack(pady=(6, 2))
        self.open_folder_btn.configure(state=tk.DISABLED, bg=DISABLED_GRAY)

        # --- Status ---
        self.status_label = tk.Label(
            self.root,
            text="",
            bg=BG_COLOR,
            fg=EXCEL_GREEN_DARK,
            font=("Segoe UI", 11, "bold"),
            wraplength=460,
            justify="center",
        )
        self.status_label.pack(pady=(14, 0))

        # --- Footer: developer details ---
        tk.Frame(self.root, bg="#d7ddda", height=1).pack(side=tk.BOTTOM, fill=tk.X, padx=20)

        footer = tk.Frame(self.root, bg=BG_COLOR)
        footer.pack(side=tk.BOTTOM, pady=14)

        tk.Label(
            footer,
            text=f"Developed by {DEV_NAME}",
            bg=BG_COLOR,
            fg=TEXT_DARK,
            font=("Segoe UI", 10, "bold"),
        ).pack()

        contact_row = tk.Frame(footer, bg=BG_COLOR)
        contact_row.pack(pady=(4, 0))

        email_link = self._make_link_label(
            contact_row, f"✉ {DEV_EMAIL}", self.open_email
        )
        email_link.pack(side=tk.LEFT, padx=(0, 14))

        tk.Label(
            contact_row, text="|", bg=BG_COLOR, fg=TEXT_MUTED, font=("Segoe UI", 10)
        ).pack(side=tk.LEFT)

        whatsapp_link = self._make_link_label(
            contact_row, f"💬 WhatsApp {DEV_WHATSAPP_DISPLAY}", self.open_whatsapp
        )
        whatsapp_link.pack(side=tk.LEFT, padx=(14, 0))

    # --- Contact link handlers ---

    def open_email(self):
        # mailto: links are opened by the OS's default mail handler,
        # which on Windows is typically launched via the default browser.
        webbrowser.open(f"mailto:{DEV_EMAIL}")

    def open_whatsapp(self):
        # wa.me links open WhatsApp Web (or the desktop app if installed)
        # through the default browser.
        webbrowser.open(f"https://wa.me/{DEV_WHATSAPP_NUMBER}")

    # --- Core actions ---

    def on_select_file(self):
        path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not path:
            return

        self.selected_file = path
        display_name = os.path.basename(path)
        if len(display_name) > 40:
            display_name = display_name[:37] + "..."
        self.file_label.configure(text=display_name, fg=TEXT_DARK)

        self.extract_btn.configure(state=tk.NORMAL, bg=EXTRACT_GREEN)
        self.status_label.configure(text="")

    def on_extract_images(self):
        if not self.selected_file:
            return

        if not self.selected_file.lower().endswith((".xlsx", ".xlsm")):
            messagebox.showerror(
                APP_TITLE,
                "This file type isn't supported.\n\n"
                "Please select a modern Excel file (.xlsx or .xlsm).\n"
                "Older .xls files use a different internal format and "
                "aren't supported.",
            )
            return

        output_folder = filedialog.askdirectory(
            title="Choose a folder to save the extracted images"
        )
        if not output_folder:
            return

        try:
            count = extract_images_from_excel(self.selected_file, output_folder)
        except zipfile.BadZipFile:
            messagebox.showerror(
                APP_TITLE, "This doesn't look like a valid Excel file."
            )
            return
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return
        except PermissionError:
            messagebox.showerror(
                APP_TITLE,
                "Couldn't write to that folder. Please choose a different "
                "location or check folder permissions.",
            )
            return
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Something went wrong:\n{e}")
            return

        self.last_output_folder = output_folder
        self.open_folder_btn.configure(state=tk.NORMAL, bg=FOLDER_BLUE)

        self.status_label.configure(
            text=f"✓ {count} image(s) extracted to:\n{output_folder}"
        )
        messagebox.showinfo(
            APP_TITLE, f"Done! {count} image(s) extracted to:\n{output_folder}"
        )

    def on_open_folder(self):
        if self.last_output_folder and os.path.isdir(self.last_output_folder):
            try:
                os.startfile(self.last_output_folder)  # Windows-only, by design
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Couldn't open that folder:\n{e}")
        else:
            messagebox.showwarning(
                APP_TITLE, "That folder is no longer available."
            )


def main():
    root = tk.Tk()
    ExcelImageExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
