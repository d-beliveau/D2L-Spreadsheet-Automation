import pandas as pd
import glob
import os

def update_master(master_path, group_folder):
    # Always overwrite the input master file

    master_df = pd.read_excel(master_path)
    email_to_index = {email: idx for idx, email in enumerate(master_df["Email"])}
    group_files = glob.glob(os.path.join(group_folder, "*.xlsx"))

    for file in group_files:
        df = pd.read_excel(file, header=None)
        student_str = df.iloc[3, 1]     # B4
        total_str = str(df.iloc[10, 1]) # B11
        students = [s.strip() for s in student_str.split(",")]
        total_score = float(total_str.split()[0])

        for name in students:
            if name in email_to_index:
                idx = email_to_index[name]
                master_df.at[idx, "Grade"] = total_score
            else:
                print(f"Warning: {name} not found in master sheet.")

    master_df.to_excel(master_path, index=False)
    return master_path

# Add CLI support if run directly
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Import group grades into master sheet.")
    parser.add_argument("master", help="Path to master spreadsheet")
    parser.add_argument("group_folder", help="Path to folder of group spreadsheets")

    args = parser.parse_args()

    try:
        out = update_master(args.master, args.group_folder)
        print(f"✅ Success! Master file updated: {out}")
    except Exception as e:
        print(f"❌ Error: {e}")
