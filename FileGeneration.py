import pandas as pd
import os
import subprocess

# Path to the Excel file
excel_file_path = r"C:\Users\QU46838\OneDrive - Chemours\Documents\GitHub\Folder_Automation\Excel_folder_maker.xlsx"

# List of column names for creating nested folders
#folder_columns = ["ICD", "TechID"]
folder_columns = ["TechID"]

# Path to the directory where folders will be created
output_directory = r"Z:\Co-ops\Coop05\Nate Ziegler\Folder Generation\Pumps"

# Path to the virtual environment activation script
venv_activate_script = r"C:\Users\QU46838\OneDrive - Chemours\Documents\Environments\env\Scripts\activate"  # Replace with the actual path to your venv activation script

# Activate the virtual environment
subprocess.run(["source", venv_activate_script], shell=True)

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(excel_file_path)

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    current_path = output_directory
    
    # Iterate over the folder columns
    for folder_column in folder_columns:
        folder_name = str(row[folder_column])
        folder_path = os.path.join(current_path, folder_name)
        
        # Create the folder if it doesn't already exist
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"Created folder: {folder_path}")
        else:
            print(f"Folder already exists: {folder_path}")
        
        # Update the current path for the next iteration
        current_path = folder_path
    
    # Update the Excel sheet with the status
    df.at[index, 'Status'] = 'Created'
    
# Save the updated DataFrame back to the Excel file
df.to_excel(excel_file_path, index=False)
print("Excel file updated.")