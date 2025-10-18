import os
import time
import win32print
import win32api
import shutil
import json
import re
from datetime import datetime
import tkinter as tk
from tkinter import simpledialog

# ---------------- CONFIGURATION ----------------

WATCH_FOLDER = r"C:\PrintDrop"
PRINTED_FOLDER = r"C:\PrintDrop\Printed"

INTERACTIVE_MODE = True  # Ask for copies via dialog

SUPPORTED_EXTENSIONS = {
    '.pdf', '.doc', '.docx', '.txt', '.xlsx', '.rtf',
    '.csv', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff'
}

os.makedirs(PRINTED_FOLDER, exist_ok=True)

# ---------------- LOGGING ----------------

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

# ---------------- UTILITIES ----------------

def ask_print_copies(filename, filepath):
    """Ask user for number of copies in a dialog"""
    try:
        root = tk.Tk()
        root.withdraw()
        copies = simpledialog.askinteger(
            title="Print Copies",
            prompt=f"How many copies to print for:\n{filename}?",
            minvalue=1,
            maxvalue=100,
        )
        root.destroy()

        if copies is None:
            # User canceled — delete the file
            log_message(f"User canceled printing for: {filename} → file deleted.")
            try:
                os.remove(filepath)
            except Exception as e:
                log_message(f"Error deleting canceled file {filename}: {e}")
            return None
        return copies
    except Exception as e:
        log_message(f"Error showing dialog: {e}")
        return 1

def cleanup_printed_folder(folder_path, max_age_seconds=3600):
    now = time.time()
    deleted_files = 0
    try:
        for filename in os.listdir(folder_path):
            path = os.path.join(folder_path, filename)
            if os.path.isfile(path) and now - os.path.getmtime(path) > max_age_seconds:
                os.remove(path)
                deleted_files += 1
                log_message(f"Deleted old printed file: {filename}")
    except Exception as e:
        log_message(f"Cleanup error: {e}")
    if deleted_files:
        log_message(f"Cleanup complete. Deleted {deleted_files} old files.")

def is_file_ready(filepath):
    try:
        if os.path.getsize(filepath) == 0:
            return False
        with open(filepath, 'rb') as f:
            f.read(1024)
        return True
    except Exception:
        return False

# ---------------- PRINTING ----------------

def get_default_printer():
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return "No default printer"

def print_pdf_with_sumatra(filepath, copies=1, printer_name=None):
    try:
        import subprocess
        printer_name = printer_name or get_default_printer()
        sumatra_path = r"C:\PrintDrop\tools\SumatraPDF.exe"
        if not os.path.exists(sumatra_path):
            log_message("SumatraPDF not found.")
            return False
        success = 0
        for i in range(copies):
            cmd = [sumatra_path, '-print-to', printer_name, '-silent', filepath]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                log_message(f"Printed copy {i+1} of {os.path.basename(filepath)}")
                success += 1
        return success > 0
    except Exception as e:
        log_message(f"Sumatra print error: {e}")
        return False

def print_word_with_com(filepath, copies=1):
    try:
        import subprocess
        ps_script = f'''
        try {{
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $doc = $word.Documents.Open("{filepath}")
            for ($i=1; $i -le {copies}; $i++) {{
                $doc.PrintOut()
                Start-Sleep -Seconds 3
            }}
            $doc.Close()
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }} catch {{
            Write-Error $_.Exception.Message
        }}
        '''
        result = subprocess.run(['powershell', '-Command', ps_script], capture_output=True, text=True, timeout=90)
        return result.returncode == 0
    except Exception as e:
        log_message(f"Word COM error: {e}")
        return False

def print_with_specific_app(filepath, copies=1, printer_name=None):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.pdf':
        return print_pdf_with_sumatra(filepath, copies, printer_name)
    elif ext in {'.doc', '.docx'}:
        return print_word_with_com(filepath, copies)
    else:
        return False

def print_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        log_message(f"Unsupported file type: {filepath}")
        return False

    filename = os.path.basename(filepath)
    settings = get_print_settings(filepath)
    copies = settings.get('copies', 1)

    if INTERACTIVE_MODE:
        user_copies = ask_print_copies(filename, filepath)
        if user_copies is None:
            return False  # user canceled — file already deleted
        copies = user_copies

    log_message(f"Printing {filename} ({copies} copies)")

    if print_with_specific_app(filepath, copies):
        return True

    try:
        for i in range(copies):
            win32api.ShellExecute(0, "print", filepath, None, ".", 0)
        log_message(f"Sent {copies} copies to printer: {filename}")
        return True
    except Exception as e:
        log_message(f"Shell print failed for {filename}: {e}")
        return False

# ---------------- SETTINGS ----------------

def parse_filename_settings(filename):
    s = {'copies': 1}
    m = re.search(r'_copies(\d+)_', filename)
    if m:
        s['copies'] = int(m.group(1))
    return s

def load_config_file(filepath):
    path = filepath + ".config"
    if os.path.exists(path):
        try:
            with open(path, 'r') as f:
                return json.load(f)
        except Exception as e:
            log_message(f"Config read error: {e}")
    return None

def get_print_settings(filepath):
    settings = {'copies': 1}
    conf = load_config_file(filepath)
    if conf:
        settings.update(conf)
    filename = os.path.basename(filepath)
    settings.update(parse_filename_settings(filename))
    return settings

# ---------------- MAIN LOOP ----------------

def main():
    log_message("=== Auto-Print Service Started ===")
    log_message(f"Watching: {WATCH_FOLDER}")
    log_message(f"Default printer: {get_default_printer()}")
    processed = {}
    last_cleanup = 0
    CLEANUP_INTERVAL = 600

    while True:
        try:
            if not os.path.exists(WATCH_FOLDER):
                time.sleep(10)
                continue

            for name in os.listdir(WATCH_FOLDER):
                path = os.path.join(WATCH_FOLDER, name)
                if not os.path.isfile(path) or name.startswith('.'):
                    continue
                if not is_file_ready(path):
                    continue

                mtime = os.path.getmtime(path)
                if processed.get(path) == mtime:
                    continue

                if name.endswith(".config"):
                    processed[path] = mtime
                    continue

                log_message(f"New file: {name}")

                if print_file(path):
                    try:
                        dest = os.path.join(PRINTED_FOLDER, name)
                        if os.path.exists(dest):
                            base, ext = os.path.splitext(name)
                            dest = os.path.join(PRINTED_FOLDER, f"{base}_{datetime.now():%Y%m%d_%H%M%S}{ext}")
                        shutil.move(path, dest)
                        log_message(f"Moved to printed folder: {os.path.basename(dest)}")

                        cfg = path + ".config"
                        if os.path.exists(cfg):
                            shutil.move(cfg, dest + ".config")
                            log_message("Moved config file.")
                    except Exception as e:
                        log_message(f"Move error: {e}")
                else:
                    # If print_file returns False because of cancel, the file is already deleted
                    if os.path.exists(path):
                        log_message(f"Failed to print {name}")

                processed[path] = mtime

            # Forget deleted files
            for old in list(processed.keys()):
                if not os.path.exists(old):
                    del processed[old]

            # Cleanup schedule
            if time.time() - last_cleanup > CLEANUP_INTERVAL:
                cleanup_printed_folder(PRINTED_FOLDER, max_age_seconds=3600)
                last_cleanup = time.time()

        except KeyboardInterrupt:
            log_message("Service stopped by user.")
            break
        except Exception as e:
            log_message(f"Main loop error: {e}")
        time.sleep(5)

# ---------------- RUN ----------------

if __name__ == "__main__":
    main()