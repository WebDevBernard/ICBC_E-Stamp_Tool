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
- ✏️ Copies the original policy document to a shared backup folder, renaming it using the client’s first name, last name and licence plate.
- 🔍 Checks for duplicates using the client's name and transaction timestamp. It will not overide or delete anything other than empty folders in your shared backup.
- 📊 Sort files into producer folder using the producer two code.
- 📁 Match similar files - Will place files into producer folder even if it is missing the producer two code. Searches if client name is in any subfolder (including the archive) and matches that name to any producer subfolder in the root directory.
- ⏳ Will continuously archive files older than one year, as long any user runs the `icbc_e-stamp_and_copy_tool`.
- 🔢 All files archived will get reincremented as well (e.g. ABC123 (2).pdf → ABC123 (1).pdf, ABC123 (1).pdf → ABC123.pdf).
- 🆓 Free to use and share.

## How to Setup

- Copy your existing ICBC Copies folder from the shared drive to your Desktop. If you do not have a shared folder or have many unidentified policy documents across multiple computers, place all the folders into a single parent directory, then copy that entire folder to your Desktop. The script will scan that directory and all its subfolders. If duplicate files are detected, only the first matching file will be copied.

- In the `config.xlsx` Excel file, select the Bulk Copy ICBC Tool worksheet and specify the path where you copied the ICBC Copies folder. On the line below, specify the path and folder name where you want the new ICBC Copies folder to be created, preferably on your Desktop so you can easily find it later. Then enter the name codes for all producers, including former producers, along with their corresponding folder names so the script knows how to sort the files.

- Run the `bulk_copy_icbc_tool` to create the new ICBC Copies folder. Always run the tool on an empty directory. This ensures the script uses the cached read data instead of reopening each PDF, which significantly improves speed during the searching and copying process.

- A `log.txt` file will also be generated. It contains a list of any files that could not be copied, as well as files without a producer two code that were moved.

- Move the newly created ICBC Copies folder back to your shared drive. Then, in `config.xlsx`, copy that shared drive path into the ICBC E-Stamp and Copy Tool worksheet. Enter the producer name codes and corresponding folder names there as well.

- The `icbc_e-stamp_and_copy_tool` can be installed on each computer that uses ICBC Policy Centre. There is no need to install the `bulk_copy_icbc_tool` on every computer, but keep one copy as a backup in case you need to reset the folder in the future.

- Each time someone runs the `icbc_e-stamp_and_copy_tool`, it will stamp the ICBC policy document and back up an unmodified copy to the shared drive.

## Frequently Asked Questions

### It's not doing anything...

Make sure you have downloads set up properly in Microsoft Edge, open settings at `edge://settings/downloads`:

- Set downloads to "C:\Users\\<your_username>\Downloads"
- Toggle off "Ask me what to do with each download".

Make sure the path names are correct in the Excel Sheet, `config.xlsx`, and you have all the corresponding producer two code + subfolder name filled out. The Excel sheet also has to be in the same directory as the script. If you just need stamping, you can delete the Excel sheet.

### Where did the ICBC E-Stamp Copies folder go?

- Either on your Desktop or inside the script folder if you are using OneDrive Desktop. Stamping is limited to the last 10 modified pdfs in Downloads.

### Can I restamp using the backup copy? Will it make a copy into the backup folder?

- You can restamp using the backup, just open the pdf and Save As a new file in your Downloads folder. The file won't get duplicated in the shared folder if it is already there.

### Why are some files copying to the wrong folder when there is no producer two code?

- If there is no producer two code, the script will try to find a file name with the same client name. If it finds a match it will return that parent subfolder name and append that to the root directory. So if the file is called `root/archive/2023/sub1/Bernard Yang - abc123.pdf`, and the file being copied also starts with the same name `Bernard Yang - efg456.pdf`, it will copy that file to `root/sub1/Bernard Yang - efg456.pdf`. For this reason, it is important you do not manually create folders.

- If a CSR mistakenly enters the wrong producer code, just manually move that file to the root folder or correct producer folder. Doing so will prevent the file with the same client name from being copied into that folder.

### I accidentally put my archive folder into another folder, and now my archive folder has an archive folder?

- Welcome to what I call archive hell 🔱🔥. To fix this, simply run the `bulk_copy_icbc_tool` to create a new ICBC copies folder.

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
