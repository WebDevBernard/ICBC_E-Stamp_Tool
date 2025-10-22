<h1 align="center">ICBC E-Stamp Tool</h1>

This script provides a one-click solution to apply the digital validation stamp to most ICBC policy documents. It scans your Downloads folder, applies the stamp, and generates two PDF files: one customer copy and one batch copy.

<table align="center">
  <tr>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/redacted_before.png" alt="Unstamped Policy Document" /></td>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/redacted_after.png" alt="Stamped Policy Document" /></td>
  </tr>
  <tr>
    <td colspan="2" align="center">(Left) Unstamped policy document, (Right) Stamped policy document</td>
  </tr>
</table>

## Quick Setup

1. You can run this script in your terminal or create an exe by cloning this repository on to your local machine. I
   recommend creating the exe for users who do not have Python installed on their system.

```bash
git clone https://github.com/WebDevBernard/ICBC_E-Stamp_Tool.git
cd ICBC_E-Stamp_Tool
```

2. Now install the dependencies and run auto_py_to_exe

```bash
pip install -r requirements.txt
python -m auto_py_to_exe
```

3. In the GUI, select the script location under `/py/icbc_e-stamp_tool.py`. Change settings to `One File` and leave settings to `Console Based`. Browse Icon in `/py/icon.ico`. Now select
   `Convert .PY To .EXE`

4. In Microsoft Edge, open settings at `edge://settings/downloads`:

   - Set downloads to "C:\Users\\<your_username>\Downloads"
   - Toggle off "Ask me what to do with each download".

## Usage

After processing a Policy Centre transaction, double-click “icbc_e-stamp_tool.exe”. This will create a folder named “ICBC E-Stamp Copies” in the same location as the executable file, containing the stamped policy documents.

Inside this folder, the “ICBC Batch Copies” subfolder contains the stamped agent copy for batching.

## Bonus

<table align="center">
  <tr>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/transaction_timestamp.png" alt="Transaction Timestamp Area" /></td>
  </tr>
  <tr>
    <td align="center">The highlighted area shows where the script checks if it is an ICBC policy document.</td>
  </tr>
</table>

### ✅ Changes October 15, 2025:

1. Full rewrite; Made leaner and faster.
2. Removed Excel customizations, now scan up to 10 pdfs at a time.
3. Fixed issue with stamping standalone Payment Plan Agreements.
4. Is now a single exe from 200 MB to 30 MB.

## License

This project is licensed under the [AGPL-3.0 License](LICENSE).

It uses [PyMuPDF](https://pymupdf.readthedocs.io/) (based on [MuPDF](https://mupdf.com/)),
both licensed under the GNU Affero General Public License v3.0.

If you modify or redistribute this project, you must also make your source
code available under the same license.
