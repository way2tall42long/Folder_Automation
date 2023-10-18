import pandas as pd
import os
import shutil

#Not working
# Replace these with your spreadsheet and folder information
spreadsheet_file = r'C:\Users\QU46838\OneDrive - Chemours\Desktop\test\PSM Critical Equipment Database.xlsx'     # Change to your spreadsheet file path
source_folder = r'Z:\Co-ops\Coop05\Nate Ziegler\Folder Generation\ICDs_Loops_etc\Loops'                          # Change to the folder where your files are located
destination_folder = r'Z:\Co-ops\Coop05\Nate Ziegler\Folder Generation\Process Measurement Devices'              # Change to the folder where you want to organize the files
file_extension = '.url'                                                                                          # Specify File extension of files being moved

Filename_column_name = "ICD Number"                                                                              # The name of the File you want to copy
Folder_column_name = "Tech ID Number"                                                                            # Where you want to copy the files to

# Read the spreadsheet using pandas
df = pd.read_excel(spreadsheet_file)        # Change to pd.read_csv for CSV files   

# Get a list of existing destination folders
existing_folders = os.listdir(destination_folder)

# Iterate through each row in the spreadsheet
for index, row in df.iterrows():
    filename = row[Filename_column_name]
    folder_name = row[Folder_column_name]
    
    # Check if filename or folder_name is blank
    if pd.notna(filename) and pd.notna(folder_name):                                                                
        source_path = os.path.join(source_folder, filename + file_extension)
        destination_path = os.path.join(destination_folder, folder_name)
    
        # Create the destination folder if it doesn't exist
        if folder_name in existing_folders:

            # Move the file to the destination folder
            try:
                shutil.copy(source_path, destination_path)
                print(f"Moved '{filename + file_extension}' to '{folder_name}' folder.")
            except FileNotFoundError:
                print(f"File '{filename}' not found in the source folder.")
        else:
            print(f"Destination folder '{folder_name}' does not exist. Skipping file '{filename + file_extension}'.")

    else:
        print(f"Skipping row {index + 1} due to missing filename or folder_name.")


print("File organization complete.")
