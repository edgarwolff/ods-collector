import argparse
import os
import sys

import pandas as pd

# --- DEFAULT CONFIGURATION ---
DEFAULT_FOLDER = "."
DEFAULT_SHEET = "Sheet1"
DEFAULT_CELL = "A0"
# -----------------------------

def get_cell_value(folder_path, sheet_name, cell_address):
    col_letter = "".join(filter(str.isalpha, cell_address)).upper()
    row_number = int("".join(filter(str.isdigit, cell_address))) - 1

    col_index = 0
    for char in col_letter:
        col_index = col_index * 26 + (ord(char) - ord('A') + 1)
    col_index -= 1

    results = []

    if not os.path.isdir(folder_path):
        sys.stderr.write(f"Error: The folder '{folder_path}' does not exist.\n")
        return

    for filename in os.listdir(folder_path):
        if filename.endswith(".ods"):
            file_path = os.path.join(folder_path, filename)

            try:
                file_id = filename.split('-')[0]
            except IndexError:
                file_id = filename.replace(".ods", "")

            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine="odf", header=None)

                if row_number < len(df) and col_index < len(df.columns):
                    val = df.iloc[row_number, col_index]

                    results.append((file_id, val))

                else:
                    sys.stderr.write(f"[{filename}]: Cell {cell_address} out of bounds\n")

            except Exception as e:
                sys.stderr.write(f"[{filename}]: Failed to read - {e}\n")

    results.sort(key=lambda x: x[0])

    for r_id, r_val in results:
        print(f"{r_id}, {r_val};")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract a specific cell value from all ODS files in a folder.")

    parser.add_argument(
        "-f", "--folder",
        type=str,
        default=DEFAULT_FOLDER,
        help=f"Path to the folder containing ODS files (default: {DEFAULT_FOLDER})"
    )
    parser.add_argument(
        "-s", "--sheet",
        type=str,
        default=DEFAULT_SHEET,
        help=f"Name of the sheet to read (default: {DEFAULT_SHEET})"
    )
    parser.add_argument(
        "-c", "--cell",
        type=str,
        default=DEFAULT_CELL,
        help=f"Cell address to extract, e.g., 'A0' (default: {DEFAULT_CELL})"
    )

    args = parser.parse_args()

    get_cell_value(args.folder, args.sheet, args.cell)
