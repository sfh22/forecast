import streamlit as st
import pandas as pd

def main():
    st.title("Excel Manipulation App")
    
    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Read the uploaded file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file)

        # Perform data manipulations here (e.g., df manipulation)

        # Save the manipulated data to a new Excel file
        manipulated_filename = "manipulated_data.xlsx"
        df.to_excel(manipulated_filename, index=False)

        # Offer a download link for the manipulated data
        st.download_button(
            label="Download Manipulated Data",
            data=manipulated_filename,
            key="download_button"
        )

if __name__ == "__main__":
    main()
