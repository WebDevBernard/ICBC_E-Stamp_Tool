<h1 align="center">ICBC E-Stamp Tool</h1>

This script offers a one-click solution to apply a digital validation stamp to most ICBC policy documents. For your convenience, it will automatically find the pdf that looks like a policy document and apply the ICBC digital validation stamp.

<table align="center">
  <tr>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/redacted_before.png" alt="Unstamped Policy Document" /></td>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/redacted_after.png" alt="Stamped Policy Document" /></td>
  </tr>
  <tr>
    <td colspan="2" align="center">(Left) Unstamped policy document, (Right) Stamped policy document</td>
  </tr>
</table>

## Features

- 🖋️ Stamps a customer copy and batch copy ICBC policy document.
- 🔍 Checks for duplicates using the client's name and transaction timestamp. It will not overide or delete anything other than empty folders in your shared backup.
- ✏️ Copies the original policy document to a shared backup folder, renaming it using the client’s first name, last name, and license plate.
- 📊 Sort files in to producer folder using the producer two code.
- 📁 Smart sorting - Will place files into producer folder even if it is missing the producer two code. Searches if client name is in any subfolder (including the archive) and matches that name to any producer subfolder in the root directory.
- ⏳ Will continuously archive files older than two years, as long any user runs the `icbc_e-stamp_and_copy_tool`.
- 🔢 All files archived will get reincremented as well (e.g. ABC123 (3).pdf → ABC123.pdf, ABC123 (4).pdf → ABC123 (1).pdf).
- 🆓 Free to use and share.

## Usage

- Complete the `config.xlsx` Excel sheet.

- Run the `bulk_copy_icbc_tool` first to create the main shared folder.
- Always run the `bulk_copy_icbc_tool` on a new empty folder or on a path name that does not yet exist. Doing so ensures the script uses the cached “Read” data instead of reopening each PDF, which greatly improves the speed during the "Copy" process.
- If you have multiple computers containing unidentified PDFs or PDFs that the tool has already renamed, place all their folders into a single parent directory. The script will then scan that directory and all its subfolders. If duplicate files are detected, only the first matching file will be copied.
- A `log.txt` file is also generated, containing a list of any files that could not be copied, as well as files without a producer two code that were moved.

- The `icbc_e-stamp_and_copy_tool` can be placed on each computer that does ICBC Policy Centre.
- No need to put the `bulk_copy_icbc_tool` on every computer, but keep one as backup in case you ever need to reset the folder.

## Frequently Asked Questions

### It's not doing anything...

Make sure you have downloads set up properly in Microsoft Edge, open settings at `edge://settings/downloads`:

- Set downloads to "C:\Users\\<your_username>\Downloads"
- Toggle off "Ask me what to do with each download".

Make sure the path names are correct in the Excel Sheet, `config.xlsx`, and you have all the corresponding producer two code + subfolder name filled out. The Excel sheet also has to be in the same directory as the script. If you just need stamping, you can delete the Excel sheet.

### Where are my ICBC E-Stamp copies?

- Either on your Desktop or inside the script folder if you are using OneDrive Desktop. Stamping is limited to the last 10 modified pdfs in Downloads.

### Can I restamp using the backup copy? Will it make a copy into the backup folder?

- You can restamp using the backup, just copy the file back into your Downloads folder. The file won't get duplicated in the share folder if it is already there.

### Why are some files copying to the wrong folder when there is no producer two code?

- If there is no producer two code, the script will try to find a file name with the same client name. If it finds a match it will return that parent subfolder name and append that to the root directory. So if the file is called `root/archive/2023/sub1/abc123.pdf`, and the file being copied is also called `abc123.pdf`, it will copy that file to `root/sub1/abc123.pdf`. This is also why you should not manually create folders inside the shared folder.

- In order to keep files without the producer two code in ending in the wrong place, manually move those files out of the producer folder into the root directory. This includes all the archived producer folders that contains that client name. Next time, the file with the same name will get copy into the root folder.

### I accidentally put my archive folder into another folder, and now my archive folder has an archive folder?

- Welcome to what I call archive hell 🔱🔥. To fix this, simply run the `bulk_copy_icbc_tool` to create a new shared folder.

### How do I create the exe?

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

3. In the GUI, select the script location under `/py/icbc_e-stamp_and_copy_tool.py` or `/py/bulk_copy_icbc_tool.py`. Change settings to `One File` and leave settings to `Console Based`. Browse Icon in `/py/icon.ico` or `grayscale.ico`. Now select
   `Convert .PY To .EXE`

4. Fill out the `config.xlsx` Excel sheet found in the `/assets` folder and move it into the same folder as the exe.

### How can it tell what pdfs are ICBC policy documents?

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
