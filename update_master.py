import pandas as pd
import glob
import os

# Paths
master_path = "test_data/master_sheet.xlsx"
group_folder = "test_data/group_sheets/"
output_path = "test_data/master_sheet.xlsx"

# Load master sheet
master_df = pd.read_excel(master_path)

# Create email-to-index map (still using first name as email for now)
email_to_index = {email: idx for idx, email in enumerate(master_df["Email"])}

# Process each group file
group_files = glob.glob(os.path.join(group_folder, "*.xlsx"))

for file in group_files:
    group_df = pd.read_excel(file, header=None)  # No header, just raw rows

    # Extract students string and total grade
    try:
        student_str = group_df.iloc[3, 1]  # B4
        total_str = str(group_df.iloc[10, 1])  # B11
    except Exception as e:
        print(f"Error reading file {file}: {e}")
        continue

    # Clean and parse student names
    students = [s.strip() for s in student_str.split(",")]

    # Parse total grade from "11 / 15.0"
    try:
        total_score = float(total_str.split()[0])
    except Exception as e:
        print(f"Could not parse grade in {file}: {total_str}")
        continue

    # Update each student's grade
    for name in students:
        if name in email_to_index:
            idx = email_to_index[name]
            master_df.at[idx, "Grade"] = total_score
        else:
            print(f"Warning: {name} not found in master sheet.")

# Save updated master sheet
master_df.to_excel(output_path, index=False)
print(f"Updated master sheet saved to {output_path}")
