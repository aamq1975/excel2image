# Excel Image Extractor

A small Windows app: pick an `.xlsx`/`.xlsm` file, choose a folder, and
every picture embedded in that Excel file is extracted straight into
the folder — no zip is created.

Built with **Python's standard library only** (Tkinter for the UI,
`zipfile` for extraction) — no Pillow, no openpyxl, nothing extra is
bundled into the final .exe, which keeps it as small as possible.

## Files in this folder

| File | Purpose |
|---|---|
| `app.py` | The application itself |
| `icon.png` | App icon, used for the window/taskbar icon at runtime |
| `icon.ico` | Same icon converted to `.ico`, used for the `.exe` file icon |
| `installer.iss` | Inno Setup script to build a Windows installer |

## What's in the app

- Select an Excel file, extract its images to a folder you choose.
- **Open Extracted Folder** button appears once extraction succeeds,
  and opens that folder directly in File Explorer.
- The developer's email and WhatsApp number in the footer are clickable
  links: the email opens your default mail app (`mailto:` link), and
  WhatsApp opens `wa.me` in your default browser (which hands off to
  WhatsApp Web or the desktop app if installed).

## Developer

Muhammad Aamir — aamq1975@gmail.com — WhatsApp +92-300-2376216
(this already appears in the app's footer)

---

## Step 1 — Install Python on Windows

1. Download Python 3.11 or 3.12 from https://python.org/downloads (the
   Windows installer).
2. Run it, and **tick "Add python.exe to PATH"** before clicking Install.
3. Confirm it worked — open Command Prompt and run:
   ```
   python --version
   ```

Tkinter ships with the standard Windows Python installer, so no extra
install is needed for the UI.

## Step 2 — Run the app directly (optional, to test it first)

From this folder in Command Prompt:
```
python app.py
```
The window should open with the Excel logo/icon, a **Select Excel
File** button and an **Extract Images** button. Try it on a real
`.xlsx` file before packaging it.

## Step 3 — Install PyInstaller

```
pip install pyinstaller
```

## Step 4 — Build the smallest practical single-file .exe

From this folder, run:

```
pyinstaller --onefile --windowed --noupx ^
  --name "ExcelImageExtractor" ^
  --icon "icon.ico" ^
  --add-data "icon.png;." ^
  app.py
```

(`^` is the line-continuation character in Command Prompt — you can
also paste it as one single line without the `^` characters.)

What each flag does:
- `--onefile` — bundles everything into one `.exe` (simplest to hand
  to someone; a `--onedir` build starts marginally faster but is a
  folder full of files instead of one exe).
- `--windowed` — no console window pops up behind the GUI.
- `--icon "icon.ico"` — sets the `.exe` file's own icon.
- `--add-data "icon.png;."` — bundles `icon.png` so the app can show
  it as the window/taskbar icon at runtime.
- `--noupx` — skip UPX compression by default (see the size note
  below for when to turn it back on).

The finished exe will be at `dist\ExcelImageExtractor.exe`.

### Getting it smaller (optional, for the size-conscious)

A Tkinter + Python app bundled with PyInstaller will realistically
land somewhere around **8–14 MB**, because the entire Python
interpreter and Tk/Tcl runtime have to travel inside the exe — there's
no way around that with this toolchain, but a few things shave off
real size:

1. **Use a virtual environment with nothing extra installed.**
   PyInstaller bundles whatever's importable in your environment, so
   building from a clean venv (rather than your everyday Python
   install with lots of packages) keeps stray libraries out:
   ```
   python -m venv build-env
   build-env\Scripts\activate
   pip install pyinstaller
   ```
   Then run the `pyinstaller` command from Step 4 inside this venv.

2. **Try UPX compression.** Download UPX from
   https://github.com/upx/upx/releases (grab the Windows zip), unzip
   it somewhere, then rebuild pointing PyInstaller at it:
   ```
   pyinstaller --onefile --windowed ^
     --name "ExcelImageExtractor" ^
     --icon "icon.ico" ^
     --add-data "icon.png;." ^
     --upx-dir "C:\path\to\upx" ^
     app.py
   ```
   UPX can shrink the exe noticeably, though antivirus/SmartScreen is
   occasionally more suspicious of UPX-packed exes — test the result
   before distributing it.

3. **Skip `--onefile` if you want the fastest possible startup** —
   use `--onedir` instead, and zip the resulting `dist\ExcelImageExtractor\`
   folder for distribution. The total bytes are similar either way;
   `--onefile` just unpacks itself into a temp folder every time it
   runs, so `--onedir` launches a bit quicker.

There isn't a way to get a Tkinter+PyInstaller exe down to a couple of
MB — that would mean dropping Tkinter for a much lighter (and far more
work to set up) GUI toolkit, which isn't worth it for an app this
size.

## Step 5 — (Optional) Build a proper Windows installer

If you'd rather hand people a normal installer (Start Menu shortcut,
uninstaller, etc.) instead of a bare .exe:

1. Install **Inno Setup** (free): https://jrsoftware.org/isdl.php
2. Make sure `dist\ExcelImageExtractor.exe` already exists (Step 4).
3. Open `installer.iss` (in this folder) in the Inno Setup Compiler.
4. Click **Build > Compile**.
5. The installer is written to an `Output\` folder as
   `ExcelImageExtractor_Setup.exe` — that's the single file you share.

Running that installer gives the user Start Menu + optional desktop
shortcuts and a normal Windows uninstall entry.

---

## Notes

- Only `.xlsx` and `.xlsm` files are supported (modern, zip-based
  Excel formats — which is also why extraction doesn't need any
  extra library: the pictures already sit inside the file as-is
  under `xl/media/`). Older `.xls` files use a completely different
  binary format and aren't supported.
- If an Excel file has no embedded pictures, the app tells you that
  instead of creating an empty folder silently.
