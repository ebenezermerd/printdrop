import os
import time
import win32print
import win32api
import shutil
import json
import re
from datetime import datetime

# Folder to watch
WATCH_FOLDER = r"C:\PrintDrop"
PRINTED_FOLDER = r"C:\PrintDrop\Printed"

# Supported file types for printing
SUPPORTED_EXTENSIONS = {
    '.pdf', '.doc', '.docx', '.txt', '.xlsx', '.rtf',
    '.csv', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff'
}

# Ensure Printed folder exists
os.makedirs(PRINTED_FOLDER, exist_ok=True)


def log_message(message):
    """Log messages with timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


def cleanup_printed_folder(folder_path, max_age_seconds=3600):
    """Delete files older than a given age (in seconds) from the Printed folder."""
    now = time.time()
    deleted_files = 0
    try:
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                file_age = now - os.path.getmtime(file_path)
                if file_age > max_age_seconds:
                    try:
                        os.remove(file_path)
                        deleted_files += 1
                        log_message(f"Deleted old printed file: {filename}")
                    except Exception as e:
                        log_message(f"Error deleting {filename}: {e}")
    except Exception as e:
        log_message(f"Error during cleanup of Printed folder: {e}")

    if deleted_files:
        log_message(f"Cleanup complete. Deleted {deleted_files} old files.")


def is_file_ready(filepath):
    """Check if file is ready for processing (not being written to)"""
    try:
        if os.path.getsize(filepath) == 0:
            return False
        with open(filepath, 'rb') as f:
            f.read(1024)
        return True
    except (IOError, PermissionError, OSError):
        return False


def print_pdf_with_sumatra(filepath, copies=1, printer_name=None):
    """Print PDF using SumatraPDF command line"""
    try:
        if printer_name is None:
            printer_name = get_default_printer()
        filename = os.path.basename(filepath)
        sumatra_path = r"C:\PrintDrop\tools\SumatraPDF.exe"
        if not os.path.exists(sumatra_path):
            log_message(f"SumatraPDF not found at {sumatra_path}")
            return False
        import subprocess
        success_count = 0
        for copy_num in range(copies):
            try:
                cmd = [sumatra_path, '-print-to', printer_name, '-silent', filepath]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                if result.returncode == 0:
                    log_message(f"SumatraPDF print successful for copy {copy_num + 1}: {filename}")
                    success_count += 1
                else:
                    log_message(f"SumatraPDF print failed for copy {copy_num + 1}: {filename} - {result.stderr}")
            except Exception as e:
                log_message(f"SumatraPDF error for copy {copy_num + 1}: {e}")
        return success_count > 0
    except Exception as e:
        log_message(f"SumatraPDF printing error: {e}")
        return False


def print_word_with_com(filepath, copies=1):
    """Print Word document using COM automation"""
    try:
        filename = os.path.basename(filepath)
        import subprocess
        success_count = 0
        for copy_num in range(copies):
            try:
                ps_script = f'''
                try {{
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $doc = $word.Documents.Open("{filepath}")
                    $doc.PrintOut()
                    Start-Sleep -Seconds 3
                    $doc.Close()
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                    Write-Host "Success"
                }} catch {{
                    Write-Error $_.Exception.Message
                }}
                '''
                result = subprocess.run(['powershell', '-Command', ps_script],
                                        capture_output=True, text=True, timeout=60)
                if result.returncode == 0 and "Success" in result.stdout:
                    log_message(f"Word COM print successful for copy {copy_num + 1}: {filename}")
                    success_count += 1
                else:
                    log_message(f"Word COM print failed for copy {copy_num + 1}: {filename} - {result.stderr}")
            except Exception as e:
                log_message(f"Word COM error for copy {copy_num + 1}: {e}")
        return success_count > 0
    except Exception as e:
        log_message(f"Word COM printing error: {e}")
        return False


def print_with_specific_app(filepath, copies=1, printer_name=None):
    """Print using specific applications for different file types"""
    try:
        if printer_name is None:
            printer_name = get_default_printer()
        filename = os.path.basename(filepath)
        file_ext = os.path.splitext(filepath)[1].lower()

        if file_ext == '.pdf':
            log_message(f"Using SumatraPDF for {filename}")
            return print_pdf_with_sumatra(filepath, copies, printer_name)
        elif file_ext in {'.doc', '.docx'}:
            log_message(f"Using Word COM automation for {filename}")
            return print_word_with_com(filepath, copies)
        elif file_ext in {'.txt', '.csv'}:
            log_message(f"Using PowerShell Out-Printer for {filename}")
            import subprocess
            success_count = 0
            for copy_num in range(copies):
                try:
                    ps_cmd = f'Get-Content "{filepath}" | Out-Printer -Name "{printer_name}"'
                    result = subprocess.run(['powershell', '-Command', ps_cmd],
                                            capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        log_message(f"PowerShell print successful for copy {copy_num + 1}: {filename}")
                        success_count += 1
                    else:
                        log_message(f"PowerShell print failed for copy {copy_num + 1}: {filename}")
                except Exception as e:
                    log_message(f"PowerShell error for copy {copy_num + 1}: {e}")
            return success_count > 0
        else:
            log_message(f"No specific app handler for {file_ext} files")
            return False
    except Exception as e:
        log_message(f"Specific app printing error: {e}")
        return False


def print_with_powershell(filepath, copies=1, printer_name=None):
    """Simple PowerShell printing - works great for text files"""
    try:
        if printer_name is None:
            printer_name = get_default_printer()
        filename = os.path.basename(filepath)
        file_ext = os.path.splitext(filepath)[1].lower()
        if file_ext not in {'.txt', '.csv'}:
            return False
        import subprocess
        success_count = 0
        for copy_num in range(copies):
            try:
                ps_cmd = f'Get-Content "{filepath}" | Out-Printer -Name "{printer_name}"'
                result = subprocess.run(['powershell', '-Command', ps_cmd],
                                        capture_output=True, text=True, timeout=60)
                if result.returncode == 0:
                    log_message(f"PowerShell print successful for copy {copy_num + 1}: {filename}")
                    success_count += 1
                else:
                    log_message(f"PowerShell print failed for copy {copy_num + 1}: {filename} - {result.stderr}")
            except Exception as e:
                log_message(f"PowerShell print error for copy {copy_num + 1}: {e}")
        return success_count > 0
    except Exception as e:
        log_message(f"PowerShell printing error: {e}")
        return False


def print_file_fallback(filepath, copies=1):
    """Enhanced fallback printing methods"""
    try:
        filename = os.path.basename(filepath)
        file_ext = os.path.splitext(filepath)[1].lower()
        if print_with_powershell(filepath, copies):
            return True
        if file_ext in {'.txt', '.csv'}:
            import subprocess
            success_count = 0
            for copy_num in range(copies):
                try:
                    result = subprocess.run(['print', '/D:' + get_default_printer(), filepath],
                                            capture_output=True, timeout=30)
                    if result.returncode == 0:
                        log_message(f"Direct print successful for copy {copy_num + 1}: {filename}")
                        success_count += 1
                    else:
                        result2 = subprocess.run(['copy', filepath, 'PRN'],
                                                 capture_output=True, shell=True, timeout=30)
                        if result2.returncode == 0:
                            log_message(f"Raw print successful for copy {copy_num + 1}: {filename}")
                            success_count += 1
                except Exception as e:
                    log_message(f"Direct print error for copy {copy_num + 1}: {e}")
            return success_count > 0
        return False
    except Exception as e:
        log_message(f"Fallback printing error: {e}")
        return False


def print_file(filepath):
    """Print the specified file with custom settings"""
    try:
        file_ext = os.path.splitext(filepath)[1].lower()
        if file_ext not in SUPPORTED_EXTENSIONS:
            log_message(f"Skipping unsupported file type: {filepath}")
            return False
        settings = get_print_settings(filepath)
        filename = os.path.basename(filepath)
        log_message(f"Print settings for {filename}: {settings}")
        copies = settings.get('copies', 1)

        log_message(f"Attempting specific application printing for {filename}")
        if print_with_specific_app(filepath, copies):
            return True
        log_message(f"Specific app method failed, trying PowerShell for {filename}")
        if print_with_powershell(filepath, copies):
            return True

        log_message(f"PowerShell method failed, trying ShellExecute for {filename}")
        success_count = 0
        for copy_num in range(copies):
            try:
                win32api.ShellExecute(0, "print", filepath, None, ".", 0)
                success_count += 1
                if copies > 1:
                    log_message(f"Sent copy {copy_num + 1}/{copies} to printer: {filename}")
            except Exception as e:
                log_message(f"Error printing copy {copy_num + 1} of {filename}: {e}")
                if copy_num == 0:
                    log_message(f"Attempting enhanced fallback printing method for {filename}")
                    if print_file_fallback(filepath, copies):
                        return True
        if success_count > 0:
            if copies == 1:
                log_message(f"Sent to printer: {filename}")
            else:
                log_message(f"Sent {success_count}/{copies} copies to printer: {filename}")
            return True
        else:
            log_message(f"Failed to print any copies of: {filename}")
            return False
    except Exception as e:
        log_message(f"Error printing {filepath}: {e}")
        return False


def get_default_printer():
    """Get the default printer name"""
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return "No default printer"


def parse_filename_settings(filename):
    """Parse print settings from filename format: filename_copies3_pages1-5.pdf"""
    settings = {'copies': 1, 'pages': None}
    copies_match = re.search(r'_copies(\d+)_', filename.lower())
    if copies_match:
        settings['copies'] = int(copies_match.group(1))
    pages_match = re.search(r'_pages([0-9,-]+)_', filename.lower())
    if pages_match:
        settings['pages'] = pages_match.group(1)
    return settings


def load_config_file(filepath):
    """Load print configuration from .config file"""
    config_path = filepath + '.config'
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            log_message(f"Error reading config file {config_path}: {e}")
    return None


def get_print_settings(filepath):
    """Get print settings for a file (config file > filename > defaults)"""
    settings = {'copies': 1, 'pages': None, 'duplex': False, 'color': True, 'quality': 'normal'}
    config_settings = load_config_file(filepath)
    if config_settings:
        settings.update(config_settings)
        return settings
    filename = os.path.basename(filepath)
    filename_settings = parse_filename_settings(filename)
    settings.update(filename_settings)
    return settings


def main():
    """Main monitoring loop"""
    log_message("=== Auto-Print Service Started ===")
    log_message(f"Watching folder: {WATCH_FOLDER}")
    log_message(f"Default printer: {get_default_printer()}")
    log_message(f"Supported file types: {', '.join(sorted(SUPPORTED_EXTENSIONS))}")

    processed_files = {}
    last_cleanup = 0
    CLEANUP_INTERVAL = 600  # 10 minutes

    while True:
        try:
            if not os.path.exists(WATCH_FOLDER):
                log_message(f"Watch folder does not exist: {WATCH_FOLDER}")
                time.sleep(10)
                continue

            files = [os.path.join(WATCH_FOLDER, f)
                     for f in os.listdir(WATCH_FOLDER)
                     if os.path.isfile(os.path.join(WATCH_FOLDER, f)) and not f.startswith('.')]

            for filepath in files:
                filename = os.path.basename(filepath)
                if not is_file_ready(filepath):
                    log_message(f"File not ready yet: {filename}")
                    continue

                mtime = os.path.getmtime(filepath)
                last_mtime = processed_files.get(filepath)

                # Skip unchanged files only
                if last_mtime == mtime:
                    continue

                log_message(f"Processing new/updated file: {filename}")

                if filename.endswith('.config'):
                    processed_files[filepath] = mtime
                    continue

                if print_file(filepath):
                    try:
                        destination = os.path.join(PRINTED_FOLDER, filename)
                        if os.path.exists(destination):
                            name, ext = os.path.splitext(filename)
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            destination = os.path.join(PRINTED_FOLDER, f"{name}_{timestamp}{ext}")
                        shutil.move(filepath, destination)
                        log_message(f"Moved to printed folder: {os.path.basename(destination)}")

                        config_path = filepath + '.config'
                        if os.path.exists(config_path):
                            config_dest = destination + '.config'
                            shutil.move(config_path, config_dest)
                            log_message(f"Moved config file: {os.path.basename(config_dest)}")

                    except Exception as e:
                        log_message(f"Error moving file {filename}: {e}")
                else:
                    log_message(f"Failed to print: {filename}")

                processed_files[filepath] = mtime

            # Forget files that no longer exist
            for old_path in list(processed_files.keys()):
                if not os.path.exists(old_path):
                    del processed_files[old_path]

        except KeyboardInterrupt:
            log_message("Auto-print service stopped by user")
            break
        except Exception as e:
            log_message(f"Unexpected error: {e}")

        if time.time() - last_cleanup > CLEANUP_INTERVAL:
            cleanup_printed_folder(PRINTED_FOLDER, max_age_seconds=3600)
            last_cleanup = time.time()

        time.sleep(5)


if __name__ == "__main__":
    main()