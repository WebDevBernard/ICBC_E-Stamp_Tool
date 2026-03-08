<h1 align="center">ICBC E-Stamp Tool</h1>

A lightweight tool that automatically detects ICBC policy PDFs, applies the ICBC digital validation stamp, and backs up the original file to a structured shared folder.

The tool also organizes policy files by producer code, checks for duplicates, and maintains an archive of older files.

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

## Downloads

Download the required files from the latest release:

- 📄 **Config File**  
  [config.xlsx](https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/releases/download/v1.0.0/config.xlsx)

- 🗂 **Create ICBC Folder Tool**  
  [create_icbc_folder_tool.exe](https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/releases/download/v1.0.0/create_icbc_folder_tool.exe)

- 🏷 **ICBC E-Stamp and Copy Tool**  
  [icbc_e-stamp_and_copy_tool.exe](https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/releases/download/v1.0.0/icbc_e-stamp_and_copy_tool.exe)

## Setup

1. Copy your existing **ICBC Copies** folder from the shared drive to your Desktop.

2. If files exist on multiple computers, collect all PDFs into folders and place those folders inside one main folder. Copy that main folder to your Desktop.

3. Create a new folder on your Desktop (or any location) and place the following files inside it:
   - `create_icbc_folder_tool.exe`
   - `icbc_e-stamp_and_copy_tool.exe`
   - `config.xlsx`

4. Open **config.xlsx**.

5. In the **Create ICBC Folder Tool** worksheet:
   - Cell B1: path to the copied ICBC Copies folder
   - Cell B2: path where the new ICBC Copies folder should be created
   - Enter all producer codes and folder names

6. Run `create_icbc_folder_tool.exe`.

7. The tool creates a new **ICBC Copies** folder and a `log.txt` file listing:
   - files that could not be copied
   - files without a producer code

8. Move the new ICBC Copies folder back to the shared drive.

9. In **config.xlsx**, open the **ICBC E-Stamp and Copy Tool** worksheet.

10. Enter:
    - Cell B1: path to the shared ICBC Copies folder
    - Cell B2: Agency Number used for Certificate Replacement and Same Day Reprint
    - producer codes and folder names

11. Place `icbc_e-stamp_and_copy_tool.exe` and `config.xlsx` on each computer that processes ICBC transactions.

12. Create a Desktop shortcut to `icbc_e-stamp_and_copy_tool.exe`.

13. Each time the tool runs:
    - the policy PDF in Downloads is stamped
    - an unmodified copy is backed up to the shared drive

## Frequently Asked Questions

### It's not doing anything

Check Microsoft Edge download settings at:

`edge://settings/downloads`

Settings should be:

- Download location
  `C:\Users\<your_username>\Downloads`

- Turn **off**
  `Ask me what to do with each download`

Also verify:

- The paths in `config.xlsx` are correct
- The Excel file is in the same folder as the executable
- Producer codes and folders are filled in

If you only need stamping, the Excel file can be removed.

### Where is the stamped copy folder

It appears:

- On your Desktop
- Or inside the script folder if Desktop is synced with OneDrive

The tool checks the **10 most recently modified PDFs** in Downloads.

### Can I restamp a backup copy

Yes.

Open the backup PDF and use **Save As** to place it in Downloads.

The file will not be duplicated in the backup folder if it already exists.

### Why are some files copied to the wrong folder

If a producer code is missing, the tool searches for an existing file with the same client name.

Example:

Existing file:

`root/archive/2023/Bernard/Steve Smith - ABC123.pdf`

New file:

`Steve Smith - EFG456.pdf`

The tool will place the new file in:

`root/Bernard/Steve Smith - EFG456.pdf`

Do not manually create folders inside the structure.

If a producer code was entered incorrectly, move the file to the correct folder manually.

### My archive folder contains another archive folder

This usually happens if the archive folder was accidentally moved.

Run `create_icbc_folder_tool.exe` to rebuild the ICBC Copies folder.

## Building the Executables

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

1. Select the script:
   - `/py/icbc_e-stamp_and_copy_tool.py`
   - `/py/create_icbc_folder_tool.py`

2. Settings:
   - One File
   - Console Based

3. Optional icon:
   - `/py/icon.ico`
   - `/py/grayscale.ico`

4. Click **Convert .PY to .EXE**

Move `config.xlsx` from the `/assets` folder into the same directory as the executable.

## How the Tool Detects ICBC Policy PDFs

<table align="center">
<tr>
<td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/transaction_timestamp.png" alt="Transaction Timestamp Area"/></td>
</tr>
<tr>
<td align="center">The highlighted area is where the script checks for the ICBC transaction timestamp.</td>
</tr>
</table>

## License

Licensed under the **AGPL-3.0 License**

https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/LICENSE.txt

This project uses:

- PyMuPDF
- MuPDF

Both licensed under **GNU Affero General Public License v3.0**.

Any modified or redistributed versions must also publish their source code under the same license.
