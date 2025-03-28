# Import necessary libraries
import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Set the title of the web app
st.title('Excel to CSV Converter')

# Add a description to the web app
st.markdown('''
This is a simple web app that converts Excel files to CSV files.
''')

# Add a file uploader to the web app
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'], accept_multiple_files=False)

# Function to convert xlsx to csv
def convert_excel_to_csv(excel_file: openpyxl.Workbook) -> list[tuple[str, bytes]]:
    """
    Converts an Excel file to mulitple CSV files, one per sheet.

    Args:
        excel_file (openpyxl.Workbook): The uploaded Excel file object.
    
    Returns:
        list[tuple[str, bytes]]: A list of tuples where each tuple contains:
        - str: The file name of the CSV.
        - bytes: The data of the CSV file in bytes.
    """

    # Create list to store name and content of the CSV files
    csv_files = []
    
    for sheet in enumerate(excel_file.sheetnames): # Iterate over each sheet in the Excel file
        df = pd.read_excel(excel_file, sheet_name=sheet) # Read the sheet into a DataFrame
        csv_file_name = f"{excel_file.name.replace('.xlsx', '').replace('.xls', '')}_{sheet}.csv"
        csv_content = df.to_csv(index=False).encode('utf-8') # Convert the DataFrame to CSV and encode it to bytes
        csv_files.append((csv_file_name, csv_content)) # Append the CSV file name and content to the list as a tuple
    return csv_files

# Check if the file is uploaded
if uploaded_file:
    with st.spinner('Converting the file...'): # Display a spinner while the file is being converted
        excel_file = openpyxl.load_workbook(uploaded_file) # Load uploaded file
        csv_files = convert_excel_to_csv(excel_file) # Convert the Excel file to CSV
        
        st.success('File converted successfully!') # Display a success message
        st.markdown('### Download CSV Files') # Display a header for the download section
        
        # Add a download button for each CSV file
        for file_name, file_content in csv_files:
            st.download_button(
                label = f"Download {file_name}",
                data = file_content,
                file_name = file_name,
                mime = 'text/csv'
            )
        
        

