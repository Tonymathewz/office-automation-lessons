# Office automation with Python and Excel: lessons learned

Practical gotchas I hit while automating a small transport company's daily Excel reports with Python (openpyxl), LibreOffice and OneDrive. Each one cost real time to find, so I've written them down so you don't have to find them too.

**Who this is for:** anyone writing scripts that fill, update or read Excel files that other people in the office also use.

## Quick reference

| # | Problem | What you see | Fix in one line |
|---|---|---|---|
| 1 | [`insert_rows()` doesn't update formulas](#1-openpyxl-insert_rows-does-not-update-formulas) | Totals below the inserted rows are quietly wrong | Rewrite every formula with its own row number |
| 2 | [Formulas saved without values](#2-openpyxl-saves-formulas-without-their-values) | Other tools see blank cells | Recalculate with LibreOffice headless |
| 3 | [No slicer support](#3-openpyxl-cant-make-slicers-so-dont-fake-them) | Corrupted workbook if you hand-build them | Use a single-column AutoFilter |
| 4 | [Fixed print area](#4-check-the-print-area-after-adding-rows) | Last rows missing from the printout | Reset the print area after every change |
| 5 | [Synced folder "saved" but didn't](#5-a-synced-folder-saying-saved-doesnt-mean-it-saved) | Old file still there | Read the file back and compare |
| 6 | [Locked files become `name-1.xlsx`](#6-locked-files-turn-into-name-1xlsx) | Your edit seems lost | Look for `-1`, `-2` copies first |
| 7 | [Overwriting what people typed](#7-never-overwrite-what-people-typed) | Colleagues' entries disappear | Only fill blank cells, one day at a time |
| 8 | [Changes straight from emails](#8-changes-from-emails-need-a-persons-approval) | Wrong bookings nobody can explain | Preview, approve, then apply with a note |

---

## Part 1: Excel and openpyxl

### 1. `openpyxl` `insert_rows()` does not update formulas

**What you see:** you insert rows in the middle of a sheet. The rows below move down, but their formulas keep the old row numbers. For example, `=SUM(N58:O58)` moves to row 61 and still says `N58:O58`. Every total below the insert is now quietly wrong.

**Why:** openpyxl edits the file's cells, not Excel's formula engine. It never adjusts references for you.

**Fix:**
- After every `insert_rows()`, go through every data row and **rewrite** each formula with that row's own current row number.
- Check that the untouched rows are the same before and after, just shifted by the number of rows you inserted.

```python
for r in range(first_data_row, ws.max_row + 1):
    ws[f"F{r}"] = f"=CONCATENATE(C{r},E{r})"
    ws[f"P{r}"] = f"=SUM(N{r}:O{r})"
```

### 2. openpyxl saves formulas without their values

**What you see:** the file looks fine in Excel, but anything that reads it *without* opening Excel (another script, a web viewer, a preview) sees blank cells.

**Why:** openpyxl writes the formula text but no calculated value.

**Fix:** run the file through LibreOffice headless so it recalculates and saves the values:

```bash
soffice --headless --convert-to xlsx --outdir out/ report.xlsx
```

### 3. openpyxl can't make slicers, so don't fake them

**What you see:** you try to add slicers by writing the XML yourself, and the workbook won't open or Excel offers to "repair" it.

**Fix:** if people only need to switch between groups (e.g. AM / PM), a **single-column AutoFilter** on that column works almost as well and is safe:

```python
ws.auto_filter.ref = f"C1:C{ws.max_row}"   # filter on the "Day Time" column only
```

### 4. Check the print area after adding rows

**What you see:** the last row is missing from the printout.

**Why:** a fixed print area (`A1:I75`) doesn't grow when you add rows.

**Fix:** set it again after any change that adds or removes rows:

```python
ws.print_area = f"A1:I{ws.max_row}"
```

---

## Part 2: Files, OneDrive and shared folders

### 5. A synced folder saying "saved" doesn't mean it saved

**What you see:** writing a file into a OneDrive-synced folder reports success, yet a moment later the old file is still there, with the same modified time and the same content.

**Fix:** after every write, **read it back** and confirm that the modified time changed and the content is the new content. Don't trust the "written" message.

### 6. Locked files turn into `name-1.xlsx`

**What you see:** your edit (or a colleague's) seems to have vanished.

**Why:** if someone has the file open in Word or Excel when you save, the save may land as `report-1.xlsx` next to the original.

**Fix:** look for `-1`, `-2` copies before assuming anything was lost, and before you edit the "original".

---

## Part 3: Safe automation with people in the loop

### 7. Never overwrite what people typed

When a script fills a shared tracking sheet every day:
- **Only fill blank cells.** Never overwrite a cell a person typed into.
- Fill **one day at a time** (tomorrow, plus a re-check of today and yesterday), not the whole month in advance.
- **Re-read the file** just before editing, because people change it between runs.

### 8. Changes from emails need a person's approval

When an assistant (a script or an AI) reads emails and changes a booking system:
- Show a **preview** of every change, and apply nothing until a person approves it.
- Put the reason, the date of the email and the sender in the record's notes, so anyone can see later why it changed.
- Amendments often arrive as replies in **old email threads**. Handle those too, not only new emails.

---

## Checklist before you run an Excel script on a shared file

- [ ] I'm working on a **copy**, or I have a backup of the original
- [ ] I **re-read the file** right before editing (someone may have changed it)
- [ ] The script **only fills blank cells** and never overwrites typed entries
- [ ] After inserting rows, every **formula is rewritten** with its own row number
- [ ] The **print area** is reset to the new last row
- [ ] If other tools read the file, it has been **recalculated** (LibreOffice headless)
- [ ] After saving, I **read the file back** and confirmed the change is really there
- [ ] I checked for **`-1` / `-2` copies** next to the file
- [ ] Any change that came from an **email was previewed and approved** by a person, with a note

---

## Related tools

These lessons came from building:
- [excel-web-viewer](https://github.com/Tonymathewz/excel-web-viewer): turn a folder of Excel workbooks into one searchable web page
- [office-claude-skills](https://github.com/Tonymathewz/office-claude-skills): approve-first Claude skills for office work and bookkeeping

MIT License. Use freely.
