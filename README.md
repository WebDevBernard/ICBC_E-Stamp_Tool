<h1 align="center">ICBC E-Stamp Tool</h1>

## Downloads

Download the required files from the latest release:

- 📄 **Config File**  
  [config.xlsx](https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/releases/download/v1.0.0/config.xlsx)

- 🏷 **ICBC E-Stamp Tool**  
  [icbc_e-stamp_and_copy_tool.exe](https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/releases/download/v1.0.0/icbc_e-stamp_tool.exe)

## Summary

A lightweight tool that automatically detects ICBC policy document PDFs, applies the digital validation stamp, and backs up the original file to a shared folder.

<table align="center">
<tr>
<td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/redacted_before.png" alt="Unstamped Policy Document"/></td>
<td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/redacted_after.png" alt="Stamped Policy Document"/></td>
</tr>
<tr>
<td colspan="2" align="center">(Left) Unstamped policy document • (Right) Stamped policy document</td>
</tr>
</table>

## Features

- 🖋️ Automatically stamps ICBC customer and batch copy policy documents
- ✏️ Backs up the original PDF to a shared folder using client name and license plate
- 🔍 Duplicate protection using client name and transaction timestamp
- 📊 Sorts files into producer folders using producer two code
- 📁 Matches similar files when the producer code is missing by checking existing client names
- ⏳ Automatically archives files older than one year when the tool runs
- 🆓 Free to use and share

## Setup

### First-Time Setup — Build the ICBC Copies Folder

1. Copy your existing **ICBC Copies** folder from the shared drive to your Desktop.

2. If files exist on multiple computers, collect all PDFs into folders and place those folders inside one main folder. Copy that main folder to your Desktop.
   _(This step is only if you do not have an existing shared folders or you have many unidentified pdfs)_

3. Create a new folder (on your Desktop or any location) and place the following files inside it:
   - `icbc_e-stamp_and_copy_tool.exe`
   - `config.xlsx`

4. Open **config.xlsx** and go to the **Config** worksheet.

5. Set **Event B3** to: `Create ICBC Copies Folder Tool`

6. Fill in the following cells:
   - **B7** — path to the copied ICBC Copies folder (input)
   - **B9** — path where the new ICBC Copies folder should be created (output)
   - **A18 / B18 onwards** — producer codes and folder names (including ex-CSRs and ex-producers)

7. Run `icbc_e-stamp_tool.exe`.

8. The tool creates a new **ICBC Copies** folder and a `log.txt` file listing:
   - files that could not be copied
   - files without a producer code

9. Move the new ICBC Copies folder back to the shared drive.

---

### Daily Use — ICBC E-Stamp and Copy Tool

10. In **config.xlsx**, set **Event B3** back to: `ICBC E-Stamp and Copy Tool`  
    _(This is the default — if B3 is blank or config.xlsx is missing, the tool runs this mode automatically.)_

11. Fill in the following cells:
    - **B13** — Path to the shared ICBC copies folder created in Step 9
    - **B15** — Agency Number used for Certificate Replacement and Same Day Reprint

12. Place `icbc_e-stamp_and_copy_tool.exe` and `config.xlsx` on each computer that processes ICBC transactions.

13. Create a Desktop shortcut to `icbc_e-stamp_and_copy_tool.exe`.

14. Each time the tool runs:
    - the policy PDF in Downloads is stamped
    - an unmodified copy is backed up to the shared drive

---

## Frequently Asked Questions

---

### ❓ It's not doing anything

**Check your Microsoft Edge download settings at:** `edge://settings/downloads`

Required settings:

- **Download location:** `C:\Users\<your_username>\Downloads`
- **Ask me what to do with each download:** Off

Also verify:

- The paths in `config.xlsx` are correct
- `config.xlsx` is in the same folder as the executable
- Producer codes and folder names are filled in (rows 18 and below)

> If you only need stamping, `config.xlsx` can be removed entirely — the tool will default to the ICBC E-Stamp and Copy Tool automatically.

---

### ❓ Where is the stamped copy folder?

The stamped copy folder named **ICBC E-Stamp Copies** appears on your Desktop, or inside the script folder if your Desktop is synced with OneDrive.

> The tool checks the **10 most recently modified PDFs** in the input folder.

---

### ❓ Can I restamp a backup copy?

**Yes.** Open the backup PDF and use **Save As** to place it back in Downloads, then run the tool.

The file will not be duplicated in the backup folder if it already exists.

---

### ❓ Why are some files copied to the wrong folder?

If a producer code is missing, the tool searches for an existing file with the same client name and moves the new file to match.

**Example:**

Existing file:

```
root/_Archive/2023/CSR1/Steve Smith - ABC123.pdf
```

New file without a producer code:

```
Steve Smith - EFG456.pdf
```

The tool will place the new file in:

```
root/CSR1/Steve Smith - EFG456.pdf
```

> Do not manually create folders inside the shared backup folder. If a producer code was entered incorrectly, move the file to the correct folder manually.

---

### ❓ My archive folder contains another archive folder

This usually happens if the `_Archive` folder was accidentally moved inside itself.

**Fix:** Set B3 to `Create ICBC Copies Folder Tool` and run `icbc_e-stamp_and_copy_tool.exe` to rebuild the ICBC Copies folder cleanly.

---

## Building the Executable

Clone the repository:

```bash
git clone https://github.com/WebDevBernard/ICBC_E-Stamp_Tool.git
cd ICBC_E-Stamp_Tool
```

Install dependencies and run auto-py-to-exe:

```bash
pip install -r requirements.txt
python -m auto_py_to_exe
```

In the GUI:

1. Select the script: `/py/main.py`
2. Settings:
   - One File
   - Console Based
3. Optional: pick an `icon.ico`
4. Click **Convert .PY to .EXE**

---

## How the Tool Detects ICBC Policy PDFs

<table align="center">
<tr>
<td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/transaction_timestamp.png" alt="Transaction Timestamp Area"/></td>
</tr>
<tr>
<td align="center">The highlighted area is where the script checks for the ICBC transaction timestamp.</td>
</tr>
</table>

---

## License

Licensed under the **AGPL-3.0 License**

https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/LICENSE.txt

This project uses:

- PyMuPDF
- MuPDF

Both licensed under **GNU Affero General Public License v3.0**.

Any modified or redistributed versions must also publish their source code under the same license.
