<h1 align="center">ICBC E-Stamp Tool</h1>

I work as an insurance broker. The purpose of this script is to provide the broker a one-click solution to
stamp an ICBC policy document and email it directly to the client.

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

3. In the GUI, select the script location under `/py/icbc_e-stamp_tool.py`. Browse Icon in `/py/icon.ico`. Now select
   `Convert .PY To .EXE`

4. Open the folder where the exe was created and add a desktop shortcut for "icbc_e-stamp_tool.exe".

5. In Microsoft Edge, open settings at `edge://settings/downloads`:
    - Set downloads to "C:\Users\\<your_username>\Downloads"
    - Toggle off "Ask me what to do with each download".

6. An optional Excel file "BM3KXR.xlsx" is available for additional customization. You can leave this Excel sheet blank
   if you don't need any customizations.

## Usage

After processing a Policy Centre transaction, double-click the "icbc_e-stamp_tool.exe - Shortcut". This will create a
folder on your desktop called "ICBC E-Stamp Copies (this folder can be deleted)" with the stamped policy documents.
The "Unsorted E-Stamp Copies" sub-folder allows you to print the stamped agent copy.

## Bonus

<table align="center">
  <tr>
    <td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/transaction_timestamp.png" alt="Transaction Timestamp Area" /></td>
  </tr>
  <tr>
    <td align="center">The highlighted area shows where the script checks if it is an ICBC policy.</td>
  </tr>
</table>

If you are trying to build something similar, I've included a script that will write all the coordinates into a txt
file. Just install tabulate and run the script at the end here.

```bash
pip install tabulate
```

```python
import fitz
from pathlib import Path
from tabulate import tabulate

input_dir = Path.home() / "Downloads"
pdf_files = input_dir.glob("*.pdf")
output_dir = Path.home() / "Desktop" / "Coordinates"
output_dir.mkdir(parents=True, exist_ok=True)


def get_text(doc, mode="blocks"):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text(mode)
        field_dict[page_number + 1] = [
            [list(filter(None, item[4].split("\n"))), item[:4]] for item in wlist
        ]
    return field_dict


def write_txt_to_file(dir_path, field_dict):
    with open(dir_path, "w", encoding="utf-8") as file:
        for page, data in field_dict.items():
            file.write(f"Page: {page}\n")
            try:
                file.write(
                    f"{tabulate(data, ['Keywords', 'Coordinates'], tablefmt='grid', maxcolwidths=[50, None])}\n"
                )
            except IndexError:
                continue


def main():
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            write_txt_to_file(output_dir / f"(Block Coordinates) {pdf.stem}.txt", get_text(doc))
            write_txt_to_file(output_dir / f"(Word Coordinates) {pdf.stem}.txt", get_text(doc, "words"))


if __name__ == "__main__":
    main()
```


