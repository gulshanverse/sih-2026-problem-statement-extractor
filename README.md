# 🚀 SIH 2026 Problem Statements Extractor

> A lightweight Python utility to extract SIH 2026 Problem Statements from a locally saved HTML source and export them into structured Excel and CSV files.

## 🛠️ Built With

- 🐍 Python
- 🍲 BeautifulSoup4 — HTML parsing
- 🐼 Pandas — data processing
- 📊 OpenPyXL — Excel generation
- 🌐 HTML / DataTables — source structure
- 📄 CSV — data export

## ✨ Features

- Extracts all 226 SIH 2026 Problem Statements
- Parses the original HTML structure locally
- Cleans and structures extracted data
- Removes duplicate records
- Preserves complete problem descriptions
- Generates Excel and CSV files
- Validates extracted records
- Works offline after obtaining the HTML source
- Does not depend on the live SIH website

## 📋 Extracted Information

| Field |
|---|
| Problem Statement ID |
| Problem Statement Title |
| Description |
| Organization |
| Department |
| Category |
| Theme |
| Youtube Link |
| Dataset Link |
| Contact info |

## ⚙️ Installation

### 1. Clone the Repository

Clone the repository and move into the project directory:

~~~bash
git clone https://github.com/gulshanverse/sih-2026-problem-statement-extractor.git
cd sih-2026-problem-statement-extractor
~~~

### 2. Install Dependencies

Install the required Python packages:

~~~bash
pip install -r requirements.txt
~~~

---

## ▶️ Usage

The repository already includes the SIH 2026 HTML source file:

~~~text
sih2026ps.html
~~~

So you **do not need to download or provide another input file** for the default setup.

Simply run:

~~~bash
python extract_sih.py
~~~

The extractor will read `sih2026ps.html`, process the 226 SIH 2026 Problem Statements, validate the extracted records, and generate the output files automatically.

---

### 🔧 Custom Input / Output

You can also use a different compatible SIH HTML source file and choose a custom output location:

~~~bash
python extract_sih.py --input path/to/your_file.html --output path/to/output.xlsx
~~~

For example:

~~~bash
python extract_sih.py --input sih2026ps.html --output output/SIH_2026_Problem_Statements.xlsx
~~~

---

## 📦 Output

After a successful run, the generated files are available inside the `output/` directory:

~~~text
output/
├── SIH_2026_Problem_Statements.xlsx
└── SIH_2026_Problem_Statements.csv
~~~

### 📊 Excel Workbook

The generated Excel workbook contains:

1. **Problem Statements** — all extracted, cleaned, and deduplicated problem statements.
2. **Summary** — extraction metadata and validation results.
3. **Field Statistics** — statistics showing how many records contain each available field.
4. **Raw Data** — a less-processed representation of the extracted records for verification and debugging.

### 📄 CSV File

`SIH_2026_Problem_Statements.csv` contains the extracted problem statements in a standard UTF-8 CSV format, making the data easy to use with:

- Microsoft Excel
- Google Sheets
- Pandas
- Data analysis tools
- Other applications supporting CSV files

---

## ✅ Quick Start

For the fastest setup:

~~~bash
git clone https://github.com/gulshanverse/sih-2026-problem-statement-extractor.git
cd sih-2026-problem-statement-extractor
pip install -r requirements.txt
python extract_sih.py
~~~

That's it. The generated Excel and CSV files will be available in:

~~~text
output/
~~~

## ✅ Validation

The extractor validates:

- Expected record count
- Duplicate Problem Statement IDs
- Missing important fields
- Excel generation
- CSV generation

Expected SIH 2026 records:

**226**

## 🧩 How It Works

~~~text
SIH HTML Source
      ↓
HTML Parsing
      ↓
Problem Statement Extraction
      ↓
Data Cleaning
      ↓
Deduplication
      ↓
Validation
      ↓
Excel + CSV
~~~

## 🐛 Troubleshooting

### Fewer than 226 records are extracted

Check whether the HTML structure has changed and whether the main table is still available under:

~~~text
#dataTablePS
~~~

### HTML structure has changed

Inspect the table and modal markup in the HTML source and update the selectors in:

~~~text
extract_sih.py
~~~

### Missing dependencies

Run:

~~~bash
pip install -r requirements.txt
~~~

## 📌 Notes

- The original HTML source is left untouched.
- Extraction is performed locally.
- The tool does not depend on the live SIH website.
- Duplicate records are removed using Problem Statement ID, title, and description.
- The extractor preserves available problem-statement content rather than summarizing it.

## 📄 License

This project is open source and released under the **MIT License**.

The MIT License applies to the original source code of this project. SIH Problem Statement content and other third-party materials remain subject to their respective rights and terms.

## 🤝 Contributing

Contributions, improvements, and bug reports are welcome.

Feel free to open a **GitHub Issue** or submit a **Pull Request**.

---

<div align="center">

⭐ **If you find this project useful, consider giving it a star!**

</div>
