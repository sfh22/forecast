import streamlit as st
import pandas as pd


def main():
    st.title("Excel Manipulation App")
    
    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Read the uploaded file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file)

        # Perform data manipulations here
        all_dfs = []

        # List of agency names to look for
        agency_names = ['Starcom', 'Zenith', 'Spark', 'Digitas', 'Publicis']

        # Specify the sheet names
        sheet_names = ['PM GCC', 'PM Egypt', 'PM Levant', 'PM Iraq']

        for sheet_name in sheet_names:
            # Read the sheet with 'Client' as the header row
            df_sheet = df.parse(sheet_name=sheet_name, header=None)

            # Find the row with 'Client' and set it as the header
            header_row = df_sheet[df_sheet.apply(lambda row: row.astype(str).str.contains('Client').any(), axis=1)].index[0]
            df_sheet.columns = df_sheet.iloc[header_row]

            # Delete rows above the 'Client' row
            df_sheet = df_sheet.iloc[header_row+1:]

            # Reset the index
            df_sheet.reset_index(drop=True, inplace=True)

            agency_name = None

            # Filter out columns containing the word "Total," excluding 'Total MBC'
            columns_to_drop = [col for col in df_sheet.columns if 'Total' in str(col) and col != 'Total MBC']
            df_sheet = df_sheet.drop(columns=columns_to_drop)

            # Filter out rows where 'Client' contains the word "Total" and is not NaN
            df_sheet = df_sheet[~df_sheet['Client'].str.contains('Total', na=False)]

            # Iterate through the rows to find and store agency names
            for index, row in df_sheet.iterrows():
                if any(name in str(row['Client']) for name in agency_names):
                    agency_name = row['Client']
                    df_sheet.at[index, 'Client'] = None
                elif row['Client'] == 'Total':
                    agency_name = None
                elif agency_name:
                    df_sheet.at[index, 'Agency'] = agency_name

            # Drop rows where both 'Client' and 'Agency' are None
            df_sheet = df_sheet.dropna(subset=['Client', 'Agency'], how='all')

            # Reset the index
            df_sheet.reset_index(drop=True, inplace=True)

            # Remove empty rows
            df_sheet = df_sheet.dropna(how='all')

            # Melt the DataFrame
            df_sheet = df_sheet.melt(id_vars=['Client', 'Agency'], var_name='Channel', value_name='Value')

            # Add the 'FileName' and 'SheetName' columns
            df_sheet['FileName'] = uploaded_file.name
            df_sheet['SheetName'] = sheet_name

            # Assign 'Currency' based on sheet name
            df_sheet['Currency'] = 'EGP' if sheet_name == 'PM Egypt' else 'USD'

            # Append the DataFrame to the list
            all_dfs.append(df_sheet)

        # Concatenate all DataFrames into a single DataFrame
        final_result_df = pd.concat(all_dfs, ignore_index=True)

        # Reorder the columns in final_result_df
        final_result_df = final_result_df[['FileName', 'SheetName', 'Agency', 'Client', 'Channel', 'Currency', 'Value']]

        # Group by the specified columns and sum the 'Value' column
        grouped_result_df = final_result_df.groupby(['FileName', 'SheetName', 'Agency', 'Client', 'Channel', 'Currency'])['Value'].sum().reset_index()

        # Drop rows where 'Value' is zero
        grouped_result_df = grouped_result_df[grouped_result_df['Value'] != 0]

        # Offer a download link for the manipulated data as CSV
        if st.download_button("Download Manipulated Data as CSV", grouped_result_df.to_csv(index=False), "text/csv"):
            st.write("Thanks for downloading!")

if __name__ == "__main__":
    main()
