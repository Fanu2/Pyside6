# Jamabandi PySide6 Application

Built directly from the supplied Jamabandi Nakal source.

## Data scope

The source contains a 12-column Jamabandi table. This application retains **all
complete records and all values from columns 1-8** and ignores columns 9-12.

Application columns:

1. Khewat / Jamabandi No.
2. Khatauni No.
3. Name / Patti
4. Owner Name
5. Cultivator
6. Irrigation Source
7. Khasra / Murabba / Killa No.
8. Area / Land Type

Extracted complete records: **384**

The source also contains 5 trailing cells after the final
complete 12-column record; these are not treated as a record.

## Features

- Full supplied first-8-column sample data
- SQLite local database
- Add / edit / delete records
- Search across all 8 columns
- Import XLSX
- Export XLSX
- Export includes all database records, not only filtered records
- Unicode/Hindi support
- Excel formatting, filtering and frozen headers
- Double-click a row to edit
- Keyboard shortcuts

## Run

```bash
python -m pip install -r requirements.txt
python jamabandi_app.py
```

On first run the bundled sample XLSX is imported automatically into `jamabandi.db`.
