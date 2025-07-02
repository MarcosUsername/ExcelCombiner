import pandas as pd
import os

# Folder containing your Excel files
folder_path = 'D:/Documents/Work/Monica Data Evaluation/5-3'

# Output file
output_file = 'D:/Documents/Work/Monica Data Evaluation/5-3/merged_data.xlsx'

# List for all Data
all_data = []

# Go through folder
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)
        try:
            # Get sheet names
            sheet_names = pd.ExcelFile(file_path, engine='openpyxl').sheet_names
            
            if len(sheet_names) < 2:
                print(f"Skip {filename}: Less than 2 sheets")
                continue

            second_sheet_name = sheet_names[1]
        
            # Use second sheet, and skip first row
            df = pd.read_excel(file_path, sheet_name=second_sheet_name, skiprows=0, engine='openpyxl')
        
            # Set to list
            all_data.append(df)

        except Exception as e:
            print(f"Failed {filename}: {e}")

# Combine
if all_data:
    merged_df = pd.concat(all_data, ignore_index=True)
    merged_df.to_excel(output_file, index=False)
    print(f"\nMerged {len(all_data)} files into: {output_file}")
else:
    print("\nNo data found.")

# Save
merged_df.to_excel(output_file, index=False)

print(f"{len(all_data)} merged into {output_file}")


#=================================================================================================

# Author: Marco Silvestri
# Email: Marco.Silvestri@CabinAir.com
# Date Created: 02-07-25
# Last Modified: 02-07-25
# Description: This script is to merge multiple Excel files from a folder into one Excel file.
# Licence: MIT License
#  __  __  ____  
# |  \/  |/ ___| 
# | |\/| |\___ \ 
# | |  | | ___) |
# |_|  |_||____/ 