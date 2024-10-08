import streamlit as st
import pandas as pd
import os

# Set the upload folder
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Function to combine data from Excel sheets
def combine_excel_files(files):
    combined_data_sheet1 = None  # For combining the first sheet
    combined_data_sheet2 = None  # For combining the second sheet

    for file in files:
        filepath = os.path.join(UPLOAD_FOLDER, file.name)
        with open(filepath, "wb") as f:
            f.write(file.getbuffer())

        # Read both sheets from each Excel file
        xls = pd.ExcelFile(filepath)

        # Function to safely read a sheet with a fallback if there are fewer rows
        def safe_read_sheet(sheet_name, expected_header=3, second_sheet=False):
            if second_sheet:
                # Second sheet with header in row 2 (index 1)
                return pd.read_excel(filepath, sheet_name=sheet_name, header=1)
            else:
                df = pd.read_excel(filepath, sheet_name=sheet_name)
                if len(df) > expected_header:
                    return pd.read_excel(filepath, sheet_name=sheet_name, header=expected_header)
                else:
                    return pd.read_excel(filepath, sheet_name=sheet_name, header=0)  # Fallback to first row if fewer rows

        # Combine the first sheet data
        if combined_data_sheet1 is None:
            combined_data_sheet1 = safe_read_sheet(0)  # Initialize with the first file's first sheet
        else:
            combined_data_sheet1 = pd.concat([combined_data_sheet1, safe_read_sheet(0)], ignore_index=True)

        # Combine the second sheet data (ensure header from row 2 is used)
        if combined_data_sheet2 is None:
            combined_data_sheet2 = safe_read_sheet(1, second_sheet=True)  # Initialize with the first file's second sheet
        else:
            combined_data_sheet2 = pd.concat([combined_data_sheet2, safe_read_sheet(1, second_sheet=True)], ignore_index=True)

    return combined_data_sheet1, combined_data_sheet2

# Streamlit app layout
st.title("Excel Files Merger")

# File uploader
uploaded_files = st.file_uploader("Upload Excel files", type=['xls', 'xlsx'], accept_multiple_files=True)

if st.button("Combine"):
    if uploaded_files:
        # Combine the Excel files
        combined_sheet1, combined_sheet2 = combine_excel_files(uploaded_files)

        # Create a download link for the combined data
        combined_filepath = os.path.join(UPLOAD_FOLDER, 'combined_data.xlsx')
        with pd.ExcelWriter(combined_filepath, engine='xlsxwriter') as writer:
            combined_sheet1.to_excel(writer, sheet_name='Combined Sheet 1', index=False)
            combined_sheet2.to_excel(writer, sheet_name='Combined Sheet 2', index=False)

        # Provide a download link for the combined Excel file
        with open(combined_filepath, "rb") as f:
            st.download_button("Download Combined Excel File", f, file_name='combined_data.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        st.error("Please upload at least one Excel file.")
