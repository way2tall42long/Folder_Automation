#This generates folders based on a column in an Excel File

import pandas as pd
import os

# Path to the Excel file
excel_file_path = r"C:\Users\QU46838\OneDrive - Chemours\Documents\GitHub\Folder_Automation\Excel_folder_maker.xlsx"    # Path to the Excel file
output_directory = r"C:\Users\QU46838\OneDrive - Chemours\Desktop\test\Master_test"                                     # Path to the directory where folders will be created

# List of column names for creating nested folders
#folder_columns = ["ICD", "TechID"]                                                                                     # This will make nested folders but isn't currently used. 
folder_columns = ["Tech ID"]

# Read all sheets from the Excel file
excel_data = pd.ExcelFile(excel_file_path)
# Create a blank data frame for later
final_df = pd.DataFrame()

# Iterate through each sheet and process the data
for sheet_name in excel_data.sheet_names:
    
    # Read the sheet into the excel data frame
    df = pd.read_excel(excel_data, sheet_name)

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        current_path = output_directory
    
        # Iterate over the folder columns
        for folder_column in folder_columns:
            folder_name = str(row[folder_column])
            folder_path = os.path.join(output_directory, folder_name)
        
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

    
    # Save the new data frame
    new_working_df = pd.DataFrame(df)
    # Make a new data frame that has the data of the new data frame and the old one
    current_df = [final_df, new_working_df]
    final_df = pd.concat(current_df)
    # Overwrite the data from the old data frame using the newly created data frame

# Creates a Status excel file and writes the final status of each folder to it   
status_excel_file = os.path.join(output_directory, "Status.xlsx")
final_df.to_excel(status_excel_file, sheet_name = "Completed Items", index=False, engine='xlsxwriter')

# Save the updated DataFrame back to the Excel file
print("Excel file updated.")