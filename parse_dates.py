#!C:\Users\User\Desktop\TACC\venv\Scripts\python.exe

import pandas as pd
# take the path of the excel as argument
import sys
file_path = sys.argv[1] if len(sys.argv) > 1 else 'April_2024_04_28-7-16-59.xlsx'


def escape_special_characters(text):
    # Escape special characters in the text
    text = text.replace('"', "'")
    return text

def extract_data_from_excel(file_path):
    # Load the Excel file, skipping initial rows that don't contain the data table
    df = pd.read_excel(file_path, skiprows=3)

    # Correcting the column names based on the structure observed
    # This step is crucial as the column names in the file might not be standard
    df.columns = ['Date', 'Start', 'End', 'Hours', 'Week', 'Person', 'Work Item', 'Title', 'Team Project', 'Activity Type', 'State', 'Work Item Type', 'Comment']

    # Extract the relevant columns
    extracted_data = df[['Date', 'Title', 'Work Item', 'Hours']]

    return extracted_data


# Call the function and get the extracted data
extracted_data = extract_data_from_excel(file_path)

# Format date column to be in the format of m/d/Y
extracted_data['Date'] = extracted_data['Date'].dt.strftime('%m/%d/%Y')

# Optionally, save the extracted data to a new Excel file
output_path = 'extracted_data.xlsx'  # Define the output file name
extracted_data.to_excel(output_path, index=False)

# Save extracted data date and description to a new text file in the format of a list of lists
    # data := [
    #     ["01/24/2024", "Product Backlog Item 44693: Build five new Promotional Product subcategory pages"],
    #     ["01/25/2024", "Product Backlog Item 44693: Build five new Promotional Product subcategory pages"],
    #     ["01/26/2024", "Product Backlog Item 44693: Build five new Promotional Product subcategory pages"],
    #     ["01/29/2024", "Product Backlog Item 44693: Build five new Promotional Product subcategory pages"],
    #     ["01/31/2024", "Product Backlog Item 44693: Build five new Promotional Product subcategory pages"],
    #     ["01/31/2024", "Product Backlog Item 44693: Build five new Promotional Product subcategory pages"]
    # ]

# Save the extracted data to a text file
with open('extracted_data.txt', 'w') as file:
    file.write('data := [\n')
    for index, row in extracted_data.iterrows():
        hours = row['Hours']
        date = row['Date']
        description = row['Title']
        description = escape_special_characters(description)
        work_item = row['Work Item']

        file.write(f'    ["{date}", "{work_item}: {description}", {hours}], \n')
        
    file.write(']')
    file.close()
print("Data extraction complete. Check the extracted_data.xlsx file.")
