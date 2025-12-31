import streamlit as st
import pandas as pd
import io

def extract_bl(filename):
    base = filename.rsplit('.', 1)[0]
    parts = base.split('_')
    last_part = parts[-1]
    if '.' in last_part:
        last_part = last_part.rsplit('.', 1)[0]
    return last_part

st.title("XLSX Files Consolidator")

uploaded_files = st.file_uploader("Choose XLSX files", accept_multiple_files=True, type="xlsx")

required_cols = [
    'Sea Waybill No', 'POD', 'POL', 'POR', 'Voyage No', 'Vessel Name', 'Pre-carriage By',
    'Notify Party', 'Consignee', 'Shipper', 'Reference No', 'Booking No', 'Carrier Code',
    'POA', 'Contains CY/DOOR?', 'Freight Prepaid At', 'HTS Code', 'BL Issue Date',
    'Laden Onboard Date', 'Also Notify Party', 'Total Quantity', 'Total Quantity UOM',
    'Total Volume', 'Total Volume UOM', 'Total Actual Gross Weight', 'Total Actual Gross Weight UOM',
    'Service Contract No', 'Yusen Remarks', 'Cargo Received Date', 'Place of BL Issue',
    'BL Prepared By', 'Payment Terms', 'Original BL No', 'Line Number', 'Container Number',
    'Container Type', 'Seal Number', 'Quantity', 'Quantity UOM', 'Gross Weight',
    'Gross Weight UOM', 'Volume', 'Volume UOM', 'File name'
]

if uploaded_files:
    processed_dfs = []
    for file in uploaded_files:
        df = pd.read_excel(file)
        if not df.empty:
            # Assuming the first 33 columns (0 to 32) are the header columns A to AG
            header_cols = df.columns[:33]
            df[header_cols] = df[header_cols].ffill()
            # Add new column with BL from filename, renamed to 'File name'
            bl = extract_bl(file.name)
            df['File name'] = bl
            processed_dfs.append(df)
    
    if processed_dfs:
        consolidated_df = pd.concat(processed_dfs, ignore_index=True)
        
        # Reorder columns according to required sequence
        available_cols = [col for col in required_cols if col in consolidated_df.columns]
        missing_cols = [col for col in required_cols if col not in consolidated_df.columns]
        extra_cols = [col for col in consolidated_df.columns if col not in required_cols]
        
        # Create reordered df with missing columns as empty (pd.NA)
        reordered_df = consolidated_df[available_cols].reindex(columns=required_cols, fill_value=pd.NA)
        
        # Append extra columns at the end if any
        if extra_cols:
            reordered_df = pd.concat([reordered_df, consolidated_df[extra_cols]], axis=1)
        
        # Display preview
        st.subheader("Consolidated Data Preview")
        st.dataframe(reordered_df.head())
        
        # Download button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            reordered_df.to_excel(writer, index=False, sheet_name='Consolidated')
        output.seek(0)
        
        st.download_button(
            label="Download Consolidated XLSX",
            data=output.getvalue(),
            file_name="consolidated_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No data found in uploaded files.")
else:
    st.info("Please upload one or more XLSX files.")