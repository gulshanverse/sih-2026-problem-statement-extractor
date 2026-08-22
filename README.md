# SIH 2026 Problem Statements Extractor

This project extracts all SIH 2026 problem statements from the locally saved HTML dump and exports them to an Excel workbook.

## Installation

```bash
pip install -r requirements.txt
```

## Run

```bash
python extract_sih.py
```

Optional custom input/output paths:

```bash
python extract_sih.py --input sih2026ps.html --output output/SIH_2026_Problem_Statements.xlsx
```

## Output

The script creates the following outputs under the `output/` folder:

- `SIH_2026_Problem_Statements.xlsx`
- `SIH_2026_Problem_Statements.csv`

The workbook contains four sheets:

1. `Problem Statements` — cleaned and deduplicated records.
2. `Summary` — extraction metadata and validation results.
3. `Field Statistics` — number of records that contain each field.
4. `Raw Data` — a less-processed view of each record for debugging.

## Troubleshooting

- If fewer than 226 records are found, check whether the HTML structure changed and whether the table is still present under the same `#dataTablePS` identifier.
- If the extraction fails because the HTML layout changed, inspect the table and modal markup in the HTML dump and update the selectors in `extract_sih.py` carefully.
- If dependencies are missing, install them with `pip install -r requirements.txt`.

## Notes

- The original HTML file is left untouched.
- The script uses the actual structure in the archive instead of relying on the live SIH website.
- Duplicate records are removed using a combination of PS ID, title, and description.
