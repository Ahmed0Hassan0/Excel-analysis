import streamlit as st
import pandas as pd
from io import BytesIO
import base64

# Function to load Excel data
def load_data(file, sheet_name):
    data = pd.read_excel(file, sheet_name, engine='openpyxl')
    return data

# Function to generate a download link for the filtered data
def generate_download_link(filter_df, sheet_name, selected_column, selected_item):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    filter_df.to_excel(writer, index=False, sheet_name="Filtered Data")
    writer.save()
    output.seek(0)
    excel_file = output.read()
    b64 = base64.b64encode(excel_file).decode()
    file_name = f"{sheet_name}_{selected_column}_{selected_item}.xlsx"
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">Download Excel file</a>'
    return href

# Function to add a progress bar for data loading
def load_data_with_progress(file, sheet_name):
    with st.spinner("Loading data..."):
        data = load_data(file, sheet_name)
    st.success("Data loaded successfully!")
    return data

# Main function to analyze and display data
def main():
    st.set_page_config(page_title='Excel Sheet Analyzer', page_icon='💹', layout='wide')
    st.title("Excel Sheet Analyzer")

    # File upload
    file = st.file_uploader("Upload Excel file", type=["xlsx"])

    filter_df = None  # Initialize filter_df in the global scope

    if file is not None:
        st.sidebar.title("Select Sheet")
        sheet_name = st.sidebar.selectbox("Select a sheet", pd.ExcelFile(file).sheet_names)

        # Load selected sheet with a progress bar
        df = load_data_with_progress(file, sheet_name)

        # Filter the data
        columns = df.columns.tolist()

        selected_column = st.sidebar.selectbox("Select a column", columns)      

        # List unique items under the selected column
        if selected_column:
            unique_items = df[selected_column].unique().tolist()
            selected_item = st.sidebar.selectbox("Select an item", unique_items)
            filter_df = df[df[selected_column] == selected_item]
        # Display data
        st.write(f"### {sheet_name} Data")

        if filter_df is not None and not filter_df.empty:
            # Use st.expander to create an expandable section
            with st.expander("Dataset"):
                shwdata = st.multiselect("You can choose the columns you want to display in the table below.", filter_df.columns)
                st.dataframe(filter_df[shwdata], use_container_width=True)
                

        else:
            st.warning("No data to display. Please select a column and item.")

        # Add an image for a visual touch

        # Download filtered data with a button
        st.header("Step 3: Download Filtered Data")
        if st.button("Download Filtered Data"):
            if filter_df is not None and not filter_df[shwdata].empty:
                download_link = generate_download_link(filter_df[shwdata], sheet_name, selected_column, selected_item)
                st.markdown(download_link, unsafe_allow_html=True)
            else:
                st.warning("No data to download. Please select a column and item.")

if __name__ == "__main__":
    main()











