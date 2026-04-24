<h1 align="center">ICBC E-Stamp Tool</h1>

## Downloads

Download the required files from the latest release (v1.0.0):

- 📄 **Config File**  
  [config.xlsx](https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/releases/download/v1.0.0/config.xlsx)

- 🏷 **ICBC E-Stamp and Copy Tool**  
  [icbc_e-stamp_and_copy_tool.exe](https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/releases/download/v1.0.0/icbc_e-stamp_and_copy_tool.exe)

## Summary

A Python script that automatically detects ICBC policy document PDFs, applies the digital validation stamp, and backs up the original file to a shared folder.

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

- 🖋️ Automatically detects ICBC policy documents and stamps a customer copy and a copy for the batch
- ✏️ Backs up the original PDF to a shared folder
- 🔍 Duplicate protection using the insured name and transaction timestamp
- 📊 Sorts files into producer folders using the producer two code
- 📁 Matches files without a producer two code to a producer folder by checking for a matching insured name
- ⏳ Continuously archives files older than one year when the script runs
- 🆓 Free to use and share

## Setup

### First-Time Setup — Create ICBC Copies Folder Tool

1. Copy your existing **ICBC Copies** folder from the shared drive to your Desktop.

2. Create a new folder (on your Desktop or any location) and place the following files inside it:
   - `icbc_e-stamp_and_copy_tool.exe`
   - `config.xlsx`

3. Open **config.xlsx** and go to the **Config** worksheet.

4. Set **B3** to: `Create ICBC Copies Folder Tool`

5. Fill in the following cells:
   - **B7** — path to the copied ICBC Copies folder (input)
   - **B9** — path where the new ICBC Copies folder should be created (output)
   - **A18 / B18 onwards** — producer codes and folder names (including ex-CSRs and ex-producers)

6. Run `icbc_e-stamp_and_copy_tool.exe`.

7. The script creates a new **ICBC Copies** folder and a `log.txt` file.

<details>
<summary><b>📄 Click to view what is recorded in log.txt</b></summary>

The `log.txt` file lists files that were skipped or flagged during processing:

- files that are not ICBC policy documents
- ICBC standalone payment plans and payment receipts
- files that could not be opened
- duplicate ICBC policy documents
- files with no producer two code matched to a producer folder
- if the shared folder already exists, log of all files copied

</details>

8. Move the new ICBC Copies folder back to the shared drive.

---

### Daily Use — ICBC E-Stamp and Copy Tool

9. In **config.xlsx**, set **B3** back to: `ICBC E-Stamp and Copy Tool`

10. Fill in the following cells:
    - **B13** — Path to the shared ICBC copies folder created in Step 7 and Step 8
    - **B15** — Agency Number used for Certificate Replacement and Same Day Reprint

11. Place `icbc_e-stamp_and_copy_tool.exe` and `config.xlsx` on each computer that processes ICBC transactions.

12. Create a Desktop shortcut to `icbc_e-stamp_and_copy_tool.exe`.

13. Each time the script runs:
    - the policy document in Downloads is stamped and placed on Desktop
    - an unmodified copy is backed up to the shared drive

### ⚠️ **CRITICAL RULE**

> The script uses the client's name and transaction timestamp in the filename to find duplicates.
>
> 🚫 Do not rename files in the shared backup  
> 🚫 Do not manually create folders inside the shared backup  
> ✅ Only use the Excel sheet to create new folders

---

### Additional Usage - Create ICBC Copies Folder Tool (copying older files to shared folder)

14. The ICBC E-Stamp and Copy Tool checks only the last 10 modified PDFs. Use the Create ICBC Copies Folder Tool to catch any files missed — useful if a computer processes walk-ins only, or if a CSR forgot to run the script. Fill in the following cells:
    - **B7** — path to the Downloads folder (input)
    - **B9** — path to your shared backup folder (output)

15. Set **B3** to: `Create ICBC Copies Folder Tool`. Then run icbc_e-stamp_and_copy_tool.exe.

16. Remember to set **B3** back to: `ICBC E-Stamp and Copy Tool`

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

> If you only need stamping, `config.xlsx` can be removed entirely — the script will default to the ICBC E-Stamp and Copy Tool automatically.

---

### ❓ Where is the stamped copy folder?

The stamped copy folder named **ICBC E-Stamp Copies** appears on your Desktop, or inside the script folder if your Desktop is synced with OneDrive.

> The script checks the **10 most recently modified PDFs** in the input folder.

---

### ❓ Can I restamp a backup copy?

**Yes.** Open the backup PDF and use **Save As** to place it back in Downloads, then run the script.

> The file will not be duplicated in the backup folder if it already exists.

---

### ❓ Why are some files copied to the wrong folder?

If a producer two code is missing, the script searches for an existing file with the same client name and moves the new file to the matching producer folder.

**Example:**

Existing file:

```
root/_Archive/2023/CSR1/Steve Smith - ABC123.pdf
```

New file without a producer two code:

```
Steve Smith - EFG456.pdf
```

The script will place the new file in:

```
root/CSR1/Steve Smith - EFG456.pdf
```

> If a file was copied into the wrong producer folder due to an incorrect producer two code, move all files with that client name to the root folder. Otherwise future files will continue copying into the wrong folder. Running `icbc_e-stamp_and_copy_tool.exe` after moving them will archive the files into the correct year folder.

---

### ❓ My archive folder contains another archive folder

This usually happens if the `_Archive` folder was accidentally moved inside another folder.

> **Fix:** Follow steps 1-8 in setup to rebuild the shared folder.

---

### ❓ Why are some names not **first-middle-surname**?

ICBC prints insured names in **surname-first-middle** order on policy documents (e.g. `SMITH JOHN ROBERT`). The script attempts to reverse these into **first-middle-surname** (e.g. `John Robert Smith`), but reversal only happens when there is enough confidence the name belongs to a person rather than a company.

**The script uses the following to determine if a name belongs to a person or a company:**

```text
IF "Owner's BC Driver's Licence" is present:
    ↓
    Masked licence number (****123 present)?
        ├── YES
        │     → Treat as PERSON
        │     → Apply reversal rules
        │
        │     IF name has 4 or more parts:
        │         → Check compound surname logic:
        │             - Recognized Chinese surname (checks if it is a common
        |               Chinese surname from a list of Chinese surnames)
        │               (WONG JOHN LEE MAN → John Lee Man Wong)
        │                   → First word is surname
        │             - Second word is particle (de / van / von)
        │                   → First word is surname
        |             - Otherwise
        │               (GARCIA LOPEZ JUAN CARLOS → Juan Carlos Garcia Lopez)
        │                   → First two words are surname
        │
        └── NO
              → Treat as COMPANY
              → DO NOT reverse name

ELSE (No "Owner's BC Driver's Licence"):
    → Fallback mode (TOP / Storage / Cancel)

    Apply ONLY fallback rules:

        ├── 27-character truncation AND 4+ parts
        │       → Treat as COMPANY (do not reverse)
        │       → Reason: ICBC has a limit of 27 chars
        |         including spaces, therefore it is likely a
        |         company name unless the person has two first
        |         middle names or two surnames
        │
        ├── Corporate suffix detected (Inc / Ltd / Corp)
        │       → Treat as COMPANY (do not reverse)
        │
        └── Otherwise
                → Treat as PERSON
                → Apply basic reversal (surname-firstname-middlename →
                  firstname-middlename-surname)
                → NO compound surname detection in this mode
```

---

## How the tool detects an ICBC policy document

<table align="center">
<tr>
<td><img src="https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/images/transaction_timestamp.png" alt="Transaction Timestamp Area"/></td>
</tr>
<tr>
<td __align__="center">The highlighted area is where the script checks for the text <em>"Transaction Timestamp"</em> followed by a space and a 14-digit number to determine if it is an ICBC policy document.</td>
</tr>
</table>

---

## License

Licensed under the **AGPL-3.0 License**

https://github.com/WebDevBernard/ICBC_E-Stamp_Tool/blob/main/LICENSE.txt
