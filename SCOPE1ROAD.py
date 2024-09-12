import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import xlsxwriter

# Function to load the Excel file and merge specified sheets
def load_and_merge_sheets(file_path, sheet_names):
    merged_data = pd.DataFrame()
    for sheet_name in sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        merged_data = pd.concat([merged_data, df], ignore_index=True)
    return merged_data

# Function to map columns and process data
def map_and_process_data(merged_data, template_columns, column_mapping, specific_date):
    matched_data = pd.DataFrame(columns=template_columns)
    for template_col, client_col in column_mapping.items():
        if client_col in merged_data.columns:
            matched_data[template_col] = merged_data[client_col]
        else:
            st.warning(f"Column '{client_col}' not found in merged_data")

    # Add default values
    matched_data['CF Standard'] = 'IATA'
    matched_data['Fuel Type'] = 'DGO'
    matched_data['GAS Type'] = 'CO2'
    matched_data['Res_Date'] = specific_date

    # Convert date columns
    for col in ['Res_Date', 'Start Date', 'End Date']:
        matched_data[col] = pd.to_datetime(matched_data[col]).dt.date

    # Replace "Not in use" with NaN and drop rows with NaN in "Distance Travelled"
    matched_data["Distance Travelled"].replace("Not in use", np.nan, inplace=True)
    matched_data = matched_data.dropna(subset=["Distance Travelled"])

    return matched_data

# Function to save DataFrame as an Excel file in memory
def save_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Streamlit app layout
st.title('Excel Data Processor for Road Freight')

# Upload Excel file
uploaded_file = st.file_uploader("Choose the Road Freight Excel file", type="xlsx")

# Input for sheet names
specified_sheets = st.text_area("Enter the sheet names (comma-separated)", 
                                value="I 74182 DXB,I 74181 DXB,R 89326 DXB,T 18328 DXB,J 28671 DXB,T 18329 DXB,T 18327 DXB,T 18326 DXB")

# Date input for specific date
specific_date = st.date_input("Select the Res_Date", datetime(2024, 3, 30))

# Define the template columns directly in the code
template_columns = [
    'Country', 'City', 'Facility', 'Vehicle Type', 'Vehicle Number', 
    'Start Date', 'End Date', 'Fuel Consumed', 'Distance Travelled', 
    'CF Standard', 'Fuel Type', 'GAS Type', 'Res_Date'
]

# Column mapping dictionary
column_mapping = {
    'Country': 'Country',
    'City': 'City',
    'Facility': 'Office / Factory / Site / Location',
    'Vehicle Type': 'Vehicle Type',
    'Vehicle Number': 'Vehicle Number',
    'Start Date': 'Start Date',
    'End Date': 'End Date',
    'Fuel Consumed': 'Fuel Consumed (Litres)',
    'Distance Travelled': 'Distance Covered (Km)',
}

# Process if the file is uploaded
if uploaded_file:
    # Load data from the uploaded file
    merged_data = load_and_merge_sheets(uploaded_file, [x.strip() for x in specified_sheets.split(',')])

    # Map and process the data
    final_data = map_and_process_data(merged_data, template_columns, column_mapping, specific_date)

    # Display the processed data
    st.write("Processed Data:")
    st.dataframe(final_data)

    # Create an in-memory Excel file for download
    excel_data = save_to_excel(final_data)

    # Download the processed data as an Excel file
    st.download_button("Download Mapped Data", data=excel_data, file_name="Mapped_Road_Freight_Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Save the processed data to the user-specified path
    output_file_path = st.text_input("Enter the path to save the file", "/content/TW MAPPED DATA/ENVIORNMENT/SCOPE1/FZE/Road.xlsx")
    if st.button("Save to file"):
        final_data.to_excel(output_file_path, index=False)
        st.success(f"Data successfully written to {output_file_path}")

# Go Back button with JavaScript redirect
if st.button("Go Back to home page"):
    st.markdown(
        '''
        <script>
        window.location.href = "https://chatgpt.com/c/66e27a5b-b1c8-800a-bbbd-2bcdd964fd11";
        </script>
        ''', 
        unsafe_allow_html=True
    )
