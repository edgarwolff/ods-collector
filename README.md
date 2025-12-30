# ODS Grade Collector

A Python utility to batch-process OpenDocument Spreadsheet (`.ods`) files. It iterates through a directory of grading sheets, extracts a specific grade cell from each file, and outputs a sorted CSV-like list of Student IDs and their grades.

## Features
* **ID Parsing:** Automatically extracts the Student ID from filenames formatted as `StudentNumber-Assignment.ods` (e.g., `1234-q1.ods` → ID: `1234`).
* **CSV Ready:** Output format (`ID, Grade;`) allows for easy import into other grading software or Excel.

## Prerequisites

You need **Python 3** installed. You also need the `pandas` and `odfpy` libraries to read ODS files.

Install the dependencies:

```bash
pip install pandas odfpy
```

## Setup

1.  Place the script (e.g., `grade_collector.py`) in the folder containing your `.ods` files, or keep it in a central tools folder.
2.  Ensure your ODS files follow the naming convention: `ID-text.ods` (e.g., `1234-lab1.ods`).

## Usage

### Basic Usage
Run the script with default settings (looks in current folder, checks `Sheet1`, cell `A0`):

```bash
python grade_collector.py
```

### Customizing the Search
You can override the defaults using command-line arguments:

```bash
python grade_collector.py --folder "./submissions" --sheet "Sheet2" --cell "C159"
```

### Arguments

| Argument | Flag | Default | Description |
| :--- | :--- | :--- | :--- |
| **Folder** | `-f`, `--folder` | `.` | The directory path containing the `.ods` files. |
| **Sheet** | `-s`, `--sheet` | `Sheet1` | The name of the tab/sheet to read. |
| **Cell** | `-c`, `--cell` | `C10` | The cell address containing the grade (e.g., "B5", "Total"). |

### Output Example

The script prints to the console (standard output):

```text
1001, 18.5;
1002, 14.0;
1003, 19.5;
1004, 12.0;
```

### Saving to a File
To save the grades directly to a CSV file, use the redirection operator (`>`):

```bash
python grade_collector.py > grades.csv
```

## Important Note on Formulas
This script reads the **cached value** stored in the ODS file.
* **If the file was saved in LibreOffice/Excel:** The formula result is stored, and the script will read the grade correctly.
* **If the file was generated programmatically and never opened:** The cache might be empty. In this case, the script might return `NaN` or `0`. Open and save the files in LibreOffice to fix this, or use headless LibreOffice to convert them first.