# Office automation with Python and Excel: lessons learned

Practical gotchas I hit while automating a small transport company's daily Excel reports with Python (openpyxl), LibreOffice and OneDrive. Each one cost real time to find, so I'm writing them down so you don't have to find them too.

---

## 1. `openpyxl` `insert_rows()` does not update formulas

**What happens:** if you insert rows in the middle of a sheet, the rows below move down, but the formulas in them keep their old row numbers. For example, `=SUM(N58:O58)` moves to row 61 and still says `N58:O58`. Every total below the insert is now quietly wrong.

**Fix:**
- After every `insert_rows()`, go through every data row and **rewrite** each formula using that row's own current row number.
- Don't expect openpyxl to adjust references for you. It never does.
- Check that the untouched rows are the same before and after, just shifted by the number of rows you inserted.

```python
for r in range(first_data_row, ws.max_row + 1):
    ws[f"F{r}"] = f"=CONCATENATE(C{r},E{r})"
    ws[f"P{r}"] = f"=SUM(N{r}:O{r})"
```

## 2. openpyxl saves formulas without their values

**What happens:** openpyxl writes the formula text but no calculated value. Anything that reads the file without opening Excel (another script, a web viewer, a preview) sees blank cells.

**Fix:** run the file through LibreOffice headless so it recalculates and saves the values:

```bash
soffice --headless --convert-to xlsx --outdir out/ report.xlsx
```

## 3. openpyxl can't make slicers, so don't fake them

openpyxl has no slicer support, and hand-built slicer XML can corrupt the workbook. If people just need to switch between groups (e.g. AM / PM), a **single-column AutoFilter** on that column works better and is safe:

```python
ws.auto_filter.ref = f"C1:C{ws.max_row}"   # filter on the "Day Time" column only
```

## 4. Check the print area after adding rows

A fixed print area (`A1:I75`) doesn't grow when you add rows, so the last row silently drops off the printout. After any change that adds or removes rows, set it again:

```python
ws.print_area = f"A1:I{ws.max_row}"
```

## 5. A synced folder saying "saved" doesn't mean it saved

**What happens:** writing a file into a OneDrive-synced folder can report success, yet the old file is still there a moment later, with the same modified time and the same content.

**Fix:** after every write, **read it back** and confirm the modified time changed and the content is the new content. Don't trust the "written" message.

## 6. Locked files turn into `name-1.xlsx`

If someone has the file open in Word or Excel when you save, the save may land as `report-1.xlsx` next to the original. Look for `-1`, `-2` versions before assuming your edit (or someone else's) was lost, and before editing the "original".

## 7. Never overwrite what people typed

When a script fills a shared tracking sheet every day:
- **Only fill blank cells.** Never overwrite a cell a person typed into.
- Fill **one day at a time** (tomorrow, plus a re-check of today and yesterday), not the whole month in advance.
- Re-read the file just before editing, because people edit it in between runs.

## 8. Changes from emails need a person's approval

When an assistant reads emails and changes a booking system:
- Show a **preview** of every change and apply nothing until a person approves it.
- Put the reason, the date of the email and the sender in the record's notes, so anyone can see later why it changed.
- Amendments often arrive as replies in **old email threads**. Handle those, not only new emails.

---

MIT License. Use freely.
