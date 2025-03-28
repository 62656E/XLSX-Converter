# Import necessary libraries
import streamlit as st
import pandas as pd
from io import BytesIO

# Set the title of the web app
st.title("Excel to CSV Converter")

# Add a description to the web app
st.markdown(
    """
This is a simple web app that converts Excel files to CSV files.
"""
)

# Add a file uploader to the web app
uploaded_file = st.file_uploader(
    "Choose an Excel file", type=["xlsx", "xls"], accept_multiple_files=False
)


# Function to convert xlsx to csv
def convert_excel_to_csv(excel_file: BytesIO) -> list[tuple[str, bytes]]:
    """
    Converts an Excel file to multiple CSV files, one per sheet.

    Args:
        excel_file (BytesIO): The uploaded Excel file object.

    Returns:
        list[tuple[str, bytes]]: A list of tuples where each tuple contains:
        - str: The file name of the CSV.
        - bytes: The data of the CSV file in bytes.
    """

    # Create list to store name and content of the CSV files
    csv_files = []

    excel_data = pd.ExcelFile(excel_file)  # Use pandas to read the uploaded file

    for (
        sheet_name
    ) in excel_data.sheet_names:  # Iterate over each sheet in the Excel file
        df = pd.read_excel(excel_data, sheet_name=sheet_name)
        csv_file_name = f"{uploaded_file.name.replace('.xlsx', '').replace('.xls', '')}_{sheet_name}.csv"
        csv_content = df.to_csv(index=False).encode("utf-8")
        csv_files.append((csv_file_name, csv_content))

    return csv_files


# Check if the file is uploaded
if uploaded_file:
    with st.spinner("Converting the file..."):
        csv_files = convert_excel_to_csv(uploaded_file)

        st.success("File converted successfully!")
        st.markdown("### Download CSV Files")

        # Add a download button for each CSV file
        for file_name, file_content in csv_files:
            st.download_button(
                label=f"Download {file_name}",
                data=file_content,
                file_name=file_name,
                mime="text/csv",
            )
