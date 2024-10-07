from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    files = request.files.getlist('files')
    
    combined_data_sheet1 = None  # For combining the first sheet
    combined_data_sheet2 = None  # For combining the second sheet

    for i, file in enumerate(files):
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

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

    combined_filepath = os.path.join(UPLOAD_FOLDER, 'combined_data.xlsx')

    # Write the combined data into two separate sheets
    with pd.ExcelWriter(combined_filepath, engine='xlsxwriter') as writer:
        combined_data_sheet1.to_excel(writer, sheet_name='Combined Sheet 1', index=False)
        combined_data_sheet2.to_excel(writer, sheet_name='Combined Sheet 2', index=False)

    return send_file(combined_filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
