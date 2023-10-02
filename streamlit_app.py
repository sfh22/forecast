import streamlit as st
import pandas as pd

def main():
    st.title("Excel Manipulation and Download App")
    
    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Read the uploaded file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file, sheet_name=['PM GCC', 'PM Egypt', 'PM Levant', 'PM Iraq'])

        # Create an empty list to store DataFrames for all files
        all_dfs = []

        # List of agency names to look for
        agency_names = ['Starcom', 'Zenith', 'Spark', 'Digitas', 'Publicis']

        for sheet_name, df_sheet in df.items():  # Iterate through specified sheets
            # Read the sheet with 'Client' as the header row
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

            # Find the row with 'Client' and set it as the header
            header_row = df[df.apply(lambda row: row.astype(str).str.contains('Client').any(), axis=1)].index[0]
            df.columns = df.iloc[header_row]

            # Delete rows above the 'Client' row
            df = df.iloc[header_row+1:]

            # Reset the index
            df.reset_index(drop=True, inplace=True)

            agency_name = None

            # Filter out columns containing the word "Total," excluding 'Total MBC'
            columns_to_drop = [col for col in df.columns if 'Total' in str(col) and col != 'Total MBC']
            df = df.drop(columns=columns_to_drop)

            # Filter out rows where 'Client' contains the word "Total" and is not NaN
            df = df[~df['Client'].str.contains('Total', na=False)]

            # Iterate through the rows to find and store agency names
            for index, row in df.iterrows():
                if any(name in str(row['Client']) for name in agency_names):
                    agency_name = row['Client']
                    df.at[index, 'Client'] = None
                elif row['Client'] == 'Total':
                    agency_name = None
                elif agency_name:
                    df.at[index, 'Agency'] = agency_name

            # Drop rows where both 'Client' and 'Agency' are None
            df = df.dropna(subset=['Client', 'Agency'], how='all')

            # Reset the index
            df.reset_index(drop=True, inplace=True)

            # Remove empty rows
            df = df.dropna(how='all')

            # Melt the DataFrame
            df = df.melt(id_vars=['Client', 'Agency'], var_name='Channel', value_name='Value')

            # Add the 'FileName' and 'SheetName' columns
            df['FileName'] = sheet_name
            df['SheetName'] = sheet_name

            # Assign 'Currency' based on sheet name
            df['Currency'] = 'EGP' if sheet_name == 'PM Egypt' else 'USD'

            # Append the DataFrame to the list
            all_dfs.append(df)

        # Concatenate all DataFrames from different sheets into a single DataFrame
        final_result_df = pd.concat(all_dfs, ignore_index=True)

        # Reorder the columns in final_result_df
        final_result_df = final_result_df[['FileName', 'SheetName', 'Agency', 'Client', 'Channel', 'Currency', 'Value']]

        # Group by the specified columns and sum the 'Value' column
        grouped_result_df = final_result_df.groupby(['FileName', 'SheetName', 'Agency', 'Client', 'Channel', 'Currency'])['Value'].sum().reset_index()

        # Drop rows where 'Value' is zero
        grouped_result_df = grouped_result_df[grouped_result_df['Value'] != 0]

        # Offer a download link for the manipulated data as CSV
        if st.download_button("Download Manipulated Data as CSV", grouped_result_df.to_csv(index=False), "text.csv"):
            st.write("Thanks for downloading!")

if __name__ == "__main__":
    main()
