<h1 align="center">ICBC E-Stamp Tool</h1>

This script offers a one-click solution to apply a digital validation stamp to most ICBC policy documents. For your convenience, it will automatically find the pdf that looks like a policy document and apply the ICBC digital validation stamp.

In addition to stamping, the script includes a fillable Excel sheet that can copy and rename an unmodified or blank policy document to a shared drive or other backup location. It will preserve the metadata such as the modified date and can sort into seperate folders based on the name code in "Producer 2".

There are two scripts included: `icbc_e-stamp_and_copy_tool` and the `bulk_copy_icbc_tool`. The `bulk_copy_icbc_tool` is a tool that can do the copy step without limiting how many pdfs to scan. This allows you to create a central store .

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

3. In the GUI, select the script location under `/py/icbc_e-stamp_and_copy_tool.py` or `/py/bulk_copy_icbc_tool`. Change settings to `One File` and leave settings to `Console Based`. Browse Icon in `/py/icon.ico` or `grayscale.ico`. Now select
   `Convert .PY To .EXE`

4. In Microsoft Edge, open settings at `edge://settings/downloads`:

   - Set downloads to "C:\Users\\<your_username>\Downloads"
   - Toggle off "Ask me what to do with each download".

## Usage

After processing a Policy Centre transaction, double-click **icbc_e-stamp_tool.exe**. This will create a folder named **ICBC E-Stamp Copies** on Desktop, containing the stamped policy documents.

There is another folder that gets generated inside this folder called "ICBC Batch Copies". This contains the stamped agent copy for batching.

**This script will check for duplicates, it is not necessary to delete the output folder. Scanning is limited to 10 pdfs.**

The Excel sheet is only needed for copying blank policy documents to a backup location; the script works without it. icbc_e-stamp_and_copy_tool requires an existing folder path; subfolders must exist or files will copy to the root. This prevents accidental folder creation from typos.

The `bulk_copy_icbc_tool` does not require an output folder or any subfolder to exist already. It will automatically create the output folder and any producer folder. After the script completes, it will generate a log.txt in the script folder with all the pdfs that could not be copied.

## Bonus

<table align="center">
  <tr>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/transaction_timestamp.png" alt="Transaction Timestamp Area" /></td>
  </tr>
  <tr>
    <td align="center">The highlighted area shows where the script checks if it is an ICBC policy document.</td>
  </tr>
</table>

## License

This project is licensed under the [AGPL-3.0 License](https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/LICENSE.txt).

It uses [PyMuPDF](https://pymupdf.readthedocs.io/) (based on [MuPDF](https://mupdf.com/)),
both licensed under the GNU Affero General Public License v3.0.

If you modify or redistribute this project, you must also make your source
code available under the same license.
