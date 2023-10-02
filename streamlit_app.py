import streamlit as st
import pandas as pd
import io

def main():
    st.title("Excel Manipulation App")
    
    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Read the uploaded file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file)

        # Perform data manipulations here (e.g., df manipulation)

        # Create a BytesIO buffer to store the Excel data
        output = io.BytesIO()

        # Use Pandas to write the DataFrame to the BytesIO buffer as an Excel file
        with pd.ExcelWriter(output, engine="opynpyxl") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)

        # Set the cursor to the beginning of the buffer
        output.seek(0)

        # Offer a download link for the manipulated data
        st.download_button(
            label="Download Manipulated Data",
            data=output,
            file_name="manipulated_data.xlsx",
            key="download_button"
        )

if __name__ == "__main__":
    main()
