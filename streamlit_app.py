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

        # Use Pandas to convert the DataFrame to a CSV string
        csv_string = df.to_csv(index=False, encoding="utf-8")

        # Offer a download link for the manipulated data as a CSV file
        if st.download_button("Download Manipulated Data as CSV", csv_string, "text/csv"):
            st.write("Thanks for downloading!")

if __name__ == "__main__":
    main()
