import pandas as pd
import glob
import os
from openpyxl.utils.cell import coordinate_to_tuple, column_index_from_string  # ADDED: import to convert "G" -> 7

# Update master spreadsheet with group grades from individual group spreadsheets.
def update_master(master_path, group_folder, student_cell="B4", grade_cell="B11", grade_column="G"):  # CHANGED: added grade_column arg
    """
    Update the master Excel sheet by writing each group's total score
    into the specified grade column for every student email listed in the group file.

    Args:
        master_path (str): Path to the master spreadsheet (D2L export).
        group_folder (str): Path to folder containing group spreadsheets (.xlsx).
        student_cell (str): Excel cell (e.g., "B4") containing comma-separated student emails.
        grade_cell (str): Excel cell (e.g., "B11") containing the group's total score (e.g., "85 / 100").
        grade_column (str): Column LETTER in master sheet where grades should be written (e.g., "G").  # ADDED
    """
    master_df = pd.read_excel(master_path)
    email_to_index = {email: idx for idx, email in enumerate(master_df["Email"])}

    group_files = glob.glob(os.path.join(group_folder, "*.xlsx"))

    col_idx = column_index_from_string(grade_column) - 1  # ADDED: convert column letter -> 0-based integer index

    for file in group_files:
        df = pd.read_excel(file, header=None)

        student_row, student_col = coordinate_to_tuple(student_cell)
        grade_row, grade_col = coordinate_to_tuple(grade_cell)

        student_str = df.iloc[student_row - 1, student_col - 1]
        total_str = str(df.iloc[grade_row - 1, grade_col - 1])

        students = [s.strip() for s in student_str.split(",")]
        total_score = float(total_str.split()[0])

        for name in students:
            if name in email_to_index:
                idx = email_to_index[name]
                master_df.iat[idx, col_idx] = total_score  # CHANGED: write by numeric column index instead of "Grade" name
            else:
                print(f"Warning: {name} not found in master sheet.")

    master_df.to_excel(master_path, index=False)
    return master_path


# CLI
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Import group grades into master sheet.")
    parser.add_argument("master", help="Path to master spreadsheet")
    parser.add_argument("group_folder", help="Path to folder of group spreadsheets")
    parser.add_argument("--student_cell", default="B4", help="Cell with student names (e.g. B4)")
    parser.add_argument("--grade_cell", default="B11", help="Cell with total grade (e.g. B11)")
    parser.add_argument("--grade_column", default="G", help="Column LETTER in master sheet where grades go (e.g. G)")  # ADDED

    args = parser.parse_args()

    try:
        out = update_master(
            args.master,
            args.group_folder,
            args.student_cell,
            args.grade_cell,
            args.grade_column,  # ADDED: pass through CLI arg
        )
        print(f"Success! Master file updated: {out}")
    except Exception as e:
        print(f"Error: {e}")
