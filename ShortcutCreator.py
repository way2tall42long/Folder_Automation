#C:\Users\QU46838\Environment\env\Scripts\python
#ctrl, Shift, P select Interpreter - link to personal venv before running
#pip install pandas winshell openpyxl

import os
import pandas as pd   
import win32com.client

# Replace this with the path to your Excel file
excel_file = r"Z:\Co-ops\Coop05\Nate Ziegler\Folder Generation\Process Measurement Devices Loop and Documentum Links.xlsx"

# Path to the directory where shortcuts will be created
output_directory = r"Z:\Co-ops\Coop05\Nate Ziegler\Folder Generation\ICDs_Loops_etc\Loops"

# Replace these with the column names
sheet_name = "Loop Sheets"
file_path_column = "Completed Link"
shortcut_name_column = "Folder Name"                        #Names cannot contain "/" at this time
is_url_column = "Is URL"

# Function to create shortcuts
def create_shortcut(source_path, target_folder, shortcut_name, is_url=False):
    if is_url:
        shortcut_name = shortcut_name + ".url"
        shortcut_path = os.path.join(target_folder, shortcut_name)

        with open(shortcut_path, "w") as url_file:
            url_file.write("[InternetShortcut]\n")
            url_file.write("URL=" + source_path)
    else:
        shortcut_name = shortcut_name + ".lnk"
        shortcut_path = os.path.join(target_folder, shortcut_name)

        with Shortcut() as s:
            s.targetpath = source_path
            s.workdir = os.path.dirname(source_path)
            s.save(shortcut_path)

# Read the specific Excel tab (worksheet) into a DataFrame
df = pd.read_excel(excel_file, sheet_name)

# Check if the specified columns exist
if file_path_column in df.columns and shortcut_name_column in df.columns:
    file_paths = df[file_path_column]
    shortcut_names = df[shortcut_name_column]
    is_urls = df[is_url_column] if is_url_column in df.columns else [False] * len(df)  # Default to False if column is missing
    # Destination_folder = os.path.expanduser("~\\Desktop") //direct named at beginning 

    for file_path, shortcut_name, is_url in zip(file_paths, shortcut_names, is_urls):
        if is_url:
            create_shortcut(file_path, output_directory, shortcut_name, is_url=True)
        elif os.path.exists(file_path):
            create_shortcut(file_path, output_directory, shortcut_name)
        else:
            print(f"File not found: {file_path}")
else:
    print(f"Columns '{file_path_column}' or '{shortcut_name_column}' not found in the Excel file.")

print("Shortcuts created successfully.")
