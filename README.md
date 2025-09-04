# Windows Patch Audit Script

A PowerShell script to **quickly audit Windows Update status** on a local (or future remote) machine.  
Generates clean reports with pending updates, last scan/install times, and reboot status.

---

## ✨ Features
- Detects **pending Windows updates** (KB ID, title, severity if available).
- Reports **last scan** and **last successful install** times.
- Flags if a **reboot is required**.
- Captures **Windows Update service status** and WSUS configuration.
- Exports results to:
  - **CSV** (easy Excel import),
  - **JSON** (automation friendly),
  - **HTML** (human-readable).
- Uses **PSWindowsUpdate** if available (for richer results), otherwise falls back to the built-in Windows Update COM API.
- No external dependencies required.

---

## 📋 Requirements
- Windows 10/11 or Windows Server 2016+.
- PowerShell 5.1 or higher (included by default).
- Optional: [`PSWindowsUpdate`](https://www.powershellgallery.com/packages/PSWindowsUpdate) module for extended reporting.

---

## 🚀 Usage

Save the script as `Audit-WindowsPatches.ps1` and run from PowerShell:

```powershell
# Default run – CSV output to .\PatchAudit
.\Audit-WindowsPatches.ps1

# JSON output to a custom folder
.\Audit-WindowsPatches.ps1 -Format JSON -OutputPath 'C:\Reports\PatchAudit'

# HTML report, try PSWindowsUpdate if available
.\Audit-WindowsPatches.ps1 -Format HTML -UsePSWindowsUpdate

# Force fallback to COM API (ignore PSWindowsUpdate)
.\Audit-WindowsPatches.ps1 -UsePSWindowsUpdate:$false

---

## 📂 Output

The script creates timestamped reports in the output folder, e.g.:

PatchAudit-MyPC-20250904-1430.csv
PatchAudit-MyPC-20250904-1430.pending.json


CSV/JSON/HTML → summary of machine patch status.

.pending.json → detailed list of pending updates (KB, title, severity).

.error.json → only generated if an error occurs.

---

## ⚠️ Notes

Severity may be blank if using COM API (Microsoft doesnt always expose it.)

If PSWindowsUpdate is installed, the script automatically provides richer fields.

Remote auditing of multiple machines is planned for v1.1.

---

## 🛠️ Roadmap

v1.1 → Add remote host auditing (-ComputerName list).

v1.2 → Add Windows Defender signature freshness check.

v1.3 → Enhanced HTML styling + per-update table.

---

## 📜 License

This script is provided for **personal and organizational internal use only**.  
- You may modify and adapt it for your own environment.  
- You may share it within your team or organization.  
- You may not resell, redistribute, or repackage it as your own product.  

If you would like to include this script in a commercial product, training package, or distribution, please contact the author for permission.


---

## 👨‍💻 Author

Created by AdamBerndt.
If you find this useful, feedback and improvements are welcome!

---
