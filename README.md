# PrintDrop: Shared Drop-to-Print Host (Windows 10/11)

A small Windows-based "drop to print" service. One Windows PC (physically cabled to a printer) hosts a shared folder. Any device on the network can drop a file into that shared folder and the host will auto-print it, then move it to `Printed/`.

Works best for PDFs, Word docs, and text/CSV files. Multiple copies are supported via filename or a companion `.config` file.

- Watch folder: `C:\\PrintDrop`
- Printed folder: `C:\\PrintDrop\\Printed`
- Service script: `auto_print.py`
- Tools: `tools/SumatraPDF.exe` (for PDF printing)

> Important: The host PC must be Windows 10 or Windows 11. Other devices (Windows/macOS/Linux/iOS/Android) simply access the shared folder.

---

## 1) Windows 10/11: Share the PrintDrop folder (network share)

Follow these steps on the Windows PC that is physically connected to the printer.

1. Create the folder
	 - Open File Explorer and create `C:\\PrintDrop`.
	 - Inside it, create a subfolder `C:\\PrintDrop\\Printed`.

2. Copy the app files into `C:\\PrintDrop`
	 - Place the following into `C:\\PrintDrop`:
		 - `auto_print.py`
		 - `START_SERVICE.bat`
		 - `manage_service.ps1`
		 - `tools/` folder (must include `SumatraPDF.exe`)
		 - Optional: keep `TEMPLATES/` for reference

3. Enable network discovery and file sharing
	 - Open: Control Panel > Network and Sharing Center > Advanced sharing settings
	 - For your active profile, turn on:
		 - "Turn on network discovery"
		 - "Turn on file and printer sharing"
	 - If your environment is simple/trusted, you can temporarily disable "Password protected sharing". Otherwise, leave it on and use a Windows user account that clients can authenticate with.

4. Share the folder
	 - Right-click `C:\\PrintDrop` > Properties > Sharing tab > Advanced Sharing…
	 - Check "Share this folder".
	 - Share name: `PrintDrop`
	 - Permissions…
		 - In a trusted LAN: grant `Everyone` = Change + Read
		 - Or, add specific users/groups and grant Change + Read
	 - Apply/OK.

5. Windows Firewall (if needed)
	 - Allow "File and Printer Sharing" through Windows Defender Firewall.

6. Note your host name or IP address
	 - Press Win+R, type `cmd`, run `hostname` (copy the value), or find your IP with `ipconfig`.
	 - Your share path will be `\\\\HOSTNAME\\PrintDrop` or `\\\\IP\\PrintDrop`.

---

## 2) Make sure the printer and software are ready on the host

- Connect the printer by USB (or reliable cable) to the host PC. Install its Windows driver.
- Set the target printer as Windows default printer:
	- Settings > Bluetooth & devices > Printers & scanners > select your printer > Set as default.
- Install Python 3.9+ on the host (from python.org). During install, check "Add Python to PATH".
- Install Python dependency (pywin32):
	- Open Windows Terminal or PowerShell (Run as administrator recommended) and run:
		- `pip install pywin32`
- Ensure `tools/SumatraPDF.exe` is present at `C:\\PrintDrop\\tools\\SumatraPDF.exe`.
- Optional (for .doc/.docx): Microsoft Word must be installed for COM-based printing of Word files.

Supported file types (as of now):
- `.pdf`, `.doc`, `.docx`, `.txt`, `.xlsx`, `.rtf`, `.csv`, `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.tif`, `.tiff`

Notes and limitations:
- Copy count is honored. Page ranges/duplex/color/quality are read from config but may not apply to all file types in this version.
- PDFs use SumatraPDF with silent printing; Word uses COM automation; text/CSV uses PowerShell.

---

## 3) Start the auto-print background service on the host

Option A — quick start (Batch file):
- Double-click `START_SERVICE.bat`. It will start the Python script minimized in the background.

Option B — PowerShell manager:
- Right-click `manage_service.ps1` > Run with PowerShell
- Or open PowerShell in `C:\\PrintDrop` and run:
	- `./manage_service.ps1 start`
	- Other commands: `stop`, `status`, `restart`

Keep the PC on and connected to the network. The service watches `C:\\PrintDrop` every 5 seconds, prints ready files, then moves them to `C:\\PrintDrop\\Printed`. Files in `Printed/` older than ~1 hour are auto-cleaned.

Auto-start with Windows (recommended):
- Open Task Scheduler > Create Task…
	- General: Name = `Auto-Print Service`, Run whether user is logged on or not, Run with highest privileges.
	- Triggers: At log on (or At startup).
	- Actions: Start a program = `C:\\PrintDrop\\START_SERVICE.bat`.
	- Conditions/Settings: adjust as needed.

---

## 4) How other devices connect to the host and drop files

All devices just need SMB access to the shared folder. They don’t need printer drivers or apps—dropping the file is enough.

- Windows 10/11 clients
	- Press Win+R, type `\\\\HOSTNAME\\PrintDrop` (or `\\\\IP\\PrintDrop`) and press Enter.
	- Or map as a network drive: File Explorer > This PC > Map network drive.
	- If prompted, enter credentials for a user on the host machine that has share permission.
	- Drag files into the folder; they print automatically within a few seconds.

- macOS
	- Finder > Go > Connect to Server…
	- Enter `smb://HOSTNAME/PrintDrop` (or `smb://IP/PrintDrop`) and Connect.
	- Authenticate if prompted. Drag files into the share.

- Linux
	- File manager: Location `smb://HOSTNAME/PrintDrop` (or `smb://IP/PrintDrop`).
	- Or mount via CIFS. Example:
		- `sudo apt install cifs-utils`
		- `sudo mount -t cifs //HOSTNAME/PrintDrop /mnt/printdrop -o username=YOURUSER`

- iOS (iPhone/iPad)
	- Files app > … > Connect to Server > `smb://HOSTNAME/PrintDrop`
	- Sign in if required. Copy files into the share.

- Android
	- Many file managers support SMB (e.g., Files by Google, Samsung My Files, CX File Explorer).
	- Add a network location for `smb://HOSTNAME/PrintDrop` and copy files in.

Optional: Share the printer directly (not required for drop-to-print)
- If you also want clients to add the shared printer:
	- Host: Settings > Printers & scanners > select printer > Printer properties > Sharing tab > Share this printer.
	- Clients (Windows): `\\\\HOSTNAME\\<SharedPrinterName>` > Connect.

---

## 5) Controlling how a file prints

You have two simple ways to request multiple copies or page ranges:

1) Filename suffixes (quick)
- Add `_copiesN_` to the filename, e.g. `invoice_copies3_.pdf` prints 3 copies.
- Add `_pages1-5_` or `_pages1,3,5_` for page ranges (support varies; PDFs may require additional configuration).

2) Companion `.config` file (more explicit)
- Create a JSON file next to the document with the same name plus `.config`.
- Example: for `invoice.pdf`, create `invoice.pdf.config` containing:
	```json
	{
		"copies": 2,
		"pages": null,
		"duplex": false,
		"color": true,
		"quality": "normal"
	}
	```
- Templates available in `TEMPLATES/`:
	- `single_copy.config`
	- `multiple_copies.config`
	- `specific_pages.config`

Current behavior:
- Copy count is honored broadly.
- Page ranges/duplex/color/quality are parsed but may not apply to all file types in this version. PDF page ranges depend on SumatraPDF command-line options and are not enabled by default here.

---

## 6) Troubleshooting

- The host shows: "Watch folder does not exist"
	- Ensure `C:\\PrintDrop` exists and is spelled exactly.

- Files aren’t printing
	- Make sure the printer is set as the Windows default.
	- Confirm the file type is supported.
	- PDFs: verify `C:\\PrintDrop\\tools\\SumatraPDF.exe` exists.
	- Word docs: Microsoft Word must be installed for COM printing.
	- Large files might take a moment; the app waits until the file is fully copied.

- Network device can’t access the share
	- Use `\\\\IP\\PrintDrop` instead of hostname.
	- Confirm Advanced sharing settings > Network discovery and File/Printer sharing are ON.
	- Check share permissions and NTFS permissions allow write (Change) for your users.
	- If password-protected sharing is ON, use valid host credentials.

- Files disappear but nothing prints
	- Check `C:\\PrintDrop\\Printed` for the moved copy.
	- Look at the service console (or re-run `START_SERVICE.bat`) to see log messages.

- Want the service to start automatically
	- Use Task Scheduler as described above.

---

## 7) What’s inside this project

- `auto_print.py` — watches `C:\\PrintDrop`, prints files, moves them to `Printed/`, and cleans up older printed files.
- `START_SERVICE.bat` — starts the service minimized.
- `manage_service.ps1` — start/stop/status/restart helper.
- `tools/SumatraPDF.exe` — headless PDF printing.
- `TEMPLATES/*.config` — sample config files you can copy/rename next to your own documents.

---

## 8) Safety and privacy notes

- This is intended for a trusted local network. If you must expose it across subnets/VPNs, enforce authentication and restrict who can write to `C:\\PrintDrop`.
- The `Printed/` folder is auto-cleaned (files older than ~1 hour are deleted). If retention matters, back up `Printed/` or adjust the script.

---

## Quick checklist

- [ ] Host PC (Windows 10/11) has the printer cabled and set as default.
- [ ] `C:\\PrintDrop` exists, is shared as `PrintDrop` with write permission for intended users.
- [ ] Python 3 and `pywin32` installed on host.
- [ ] `tools/SumatraPDF.exe` present.
- [ ] Service started: `START_SERVICE.bat` or `./manage_service.ps1 start`.
- [ ] Clients can reach `\\\\HOSTNAME\\PrintDrop` and drop files.

Happy printing!
