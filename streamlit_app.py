import streamlit as st
import pandas as pd
import io

def main():
    st.title("Manipulation")
    
    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Read the uploaded file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file)

        # Perform data manipulations here (e.g., df manipulation)

        # Offer a download link for the manipulated data
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)

        output.seek(0)
        st.download_button(
            label="Download Manipulated Data",
            data=output,
            file_name="manipulated_data.xlsx",
            key="download_button"
        )

if __name__ == "__main__":
    main()
