<h1 align="center">ICBC E-Stamp Tool</h1>

<p align="center">A simple solution to digitally stamp most ICBC policy documents. Save time and paper.
The code is written in a single Python file, allows you to easily verify this script.</p>

<table align="center">
  <tr>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/redacted_before.png" alt="Unstamped Policy Document" /></td>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/redacted_after.png" alt="Stamped Policy Document" /></td>
  </tr>
  <tr>
    <td colspan="2" align="center">(Left) Unstamped policy document, (Right) Stamped policy document</td>
  </tr>
</table>

<hr/>

## How to Use

### Method 1 - Using the Packaged Zip File

- Extract the zip file and place it anywhere on your computer.
- Create a desktop shortcut for "icbc_e-stamp_tool.exe".
- In Microsoft Edge, open settings at `edge://settings/downloads`:
  - Set downloads to "C:\Users\<your_username>\Downloads"
  - Toggle off "Ask me what to do with each download".
- An optional Excel file "BM3KXR.xlsx" is available for additional customization. You can leave this Excel sheet blank if you don't need any customizations.
- After processing a Policy Centre transaction, double-click the "icbc_e-stamp_tool.exe - Shortcut". This will create a folder on your desktop called "ICBC E-Stamp Copies (this folder can be deleted)" with the stamped policy documents.
- The "Unsorted E-Stamp Copies" sub-folder allows you to print the stamped agent copy.

### Method 2 - Create the EXE File

- Clone or download the zip of this repository.
- Navigate to the file folder and install the dependencies using `pip install -r requirements.txt`.
- Run `python -m auto_py_to_exe` and create the .exe using the Python file "/py/icbc_e-stamp_tool.py".
- Follow the same steps as Method 1.

<hr/>

## Script Not Working?

- Ensure the Downloads directory in Microsoft Edge is set to "C:\Users\<your_username>\Downloads" and "Ask me what to do with each download" is toggled off.
- The script will not copy the same policy document. If it cannot copy, try deleting the "ICBC E-Stamp Copies (this folder can be deleted)" folder from your desktop.
- Note that not all documents will work. This script scans a set of coordinates for the phrase "Transaction Timestamp".

<table align="center">
  <tr>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/transaction_timestamp.png" alt="Transaction Timestamp Area" /></td>
  </tr>
  <tr>
    <td align="center">The highlighted area shows where the script checks if it is an ICBC policy.</td>
  </tr>
</table>
