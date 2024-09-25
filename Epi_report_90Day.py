# Import necessary libraries
import pandas as pd
import os
import warnings
import re

# Suppresses all warnings
warnings.filterwarnings('ignore')


def convert_datetime_to_date(df):
        """Convert datetime columns to date-only format."""
        for col in df.select_dtypes(include=['datetime64', 'datetime']):
            df[col] = df[col].dt.date
        return df


def extract_organism_name(filename, run_date):
    """
    Extracts organism name from filename based on a run_date for .csv files.
    
    Parameters:
        filename (str): The name of the file.
        run_date (str): Date of the download in mmddyy format, used to extract organism names.
    
    Returns:
        str: Extracted organism name.
    """
    basename = os.path.splitext(filename)[0]
    if filename.endswith('.csv'):
        parts = basename.split(f'{run_date} PN Export')
        organism_name = parts[1].strip() if len(parts) > 1 else basename
    else:
        parts = basename.split(' Metadata Link BaseSpace')
        organism_name = parts[0]
    return organism_name[:30]
    

def merge_files_to_sheets(csv_directory,xlsx_directory, run_date):
    """
    Merges the first sheet from all .xlsx files and all .csv files from the two different directory into a single dictionary of DataFrames with multiple sheets.
    
    Parameters:
        input_directory (str): Path to the directory containing .xlsx and .csv files.
        run_date (str): Date of the download in mmddyy format, used to extract organism names.
    
    Returns:
        dict: A dictionary of DataFrames, each corresponding to a sheet.
    """
    combined_sheets = {}

    # Merge the first sheet from Excel files
    excel_files = [file for file in os.listdir(xlsx_directory) if file.endswith('.xlsx')]
    for file in excel_files:
        file_path = os.path.join(xlsx_directory, file)
        try:
            sheets = pd.read_excel(file_path, sheet_name=0)  # Read only the first sheet
            sheets = convert_datetime_to_date(sheets)  # Convert datetime columns to date
            # Check if "PulseNet Upload Date" column exists and filter non-empty rows
            if 'PulseNet Upload Date' in sheets.columns:
                sheets = sheets[sheets['PulseNet Upload Date'].notna() & (sheets['PulseNet Upload Date'] != '')]
                
                if not sheets.empty:  # Only proceed if there are valid rows
                    sheet_name = extract_organism_name(file, run_date)  # Get the sheet name based on the organism
                    combined_sheets[sheet_name] = pd.concat([combined_sheets.get(sheet_name, pd.DataFrame()), sheets], ignore_index=True)
            else:
                print(f"'PulseNet Upload Date' column not found in {file}. Skipping.")
        except Exception as e:
            print(f"Error processing {file}: {e}")

    # Merge only recent CSV files based on run_date
    csv_files = [file for file in os.listdir(csv_directory) if file.endswith('.csv') and run_date in file]
    for file in csv_files:
        file_path = os.path.join(csv_directory, file)
        try:
            df = pd.read_csv(file_path)
            df = convert_datetime_to_date(df)  # Convert datetime columns to date
            sheet_name = extract_organism_name(file, run_date) # To get the sheet name based on the organisms
            combined_sheets[sheet_name] = pd.concat([combined_sheets.get(sheet_name, pd.DataFrame()), df], ignore_index=True)
        except Exception as e:
            print(f"Error processing {file}: {e}")
    
    return combined_sheets


def format_df(df):
    """
    Renames columns and rearranges them according to predefined headers,
    while excluding specified columns and including any extra columns specific to each DataFrame.
    
    Parameters:
        df (pd.DataFrame): The DataFrame to be formatted.
    
    Returns:
        pd.DataFrame: Formatted DataFrame with reordered columns.
    """
    rename_col_lst = {
        "Date Modified": "Modified date"
    }
    
    # Columns to be removed
    columns_to_remove = [
        "TAT in Calendar Days",
        "TAT in Workdays (minus weekends/Holidays)",
        "Comment",
        "SequencedDate",
        "PulseNet Upload Date",
        "PatientAgeYears",
        "PatientAgeMonths",
        "PatientAgeDays"
    ]

    # Define the CSV headers
    csv_headers = [
        "Key",
        "Modified date",
        "WGS_id",
        "ReceivedDate",
        "LabReceivedDate"
        "SequencerRun_id",
        "Allele_Code",
        "Outbreak",
        "REP_code",
        "NCBI_ACCESSION",
        "SRR_id",
        "LastName",
        "FirstName",
        "SourceCounty",
        "SourceState",
        "PatientDOB",
        "PATIENTAGEYEARS",
        "PATIENTAGEMONTHS",
        "PATIENTAGEDAYS",
        "SourceSite",
        "PatientSex",
        "IsolatDate",
        "SourceCountry",
        "SourceType",
        "PulseNet_UploadDate",
        "Genus",
        "Species",
        "MLST_ST",
        "LabID",
        "OtherStateIsolate",
    ]
    
    # Rename columns based on the rename mapping
    df = df.rename(columns=rename_col_lst)

    # Drop specified columns if they exist in the DataFrame
    columns_to_drop = [col for col in columns_to_remove if col in df.columns]
    if columns_to_drop:
        df = df.drop(columns=columns_to_drop, errors='ignore')
    else:
        print("No columns to drop from the DataFrame.")

    # Keep only existing columns that are in csv_headers
    existing_csv_headers = [col for col in csv_headers if col in df.columns]
    
    # Include any extra columns that are not in csv_headers
    extra_columns = [col for col in df.columns if col not in existing_csv_headers]
    
    # Combine existing CSV headers with extra columns
    final_columns = existing_csv_headers + extra_columns
    
    # Ensure final DataFrame includes only those columns
    df = df[final_columns]

    return df

def process_df(df):
    """
    Adds columns (PatientDOB, LastName and FirstName) and matches them according to the numeric/integer values in Key column
    with the numeric part of the Key column which starts with KS___ , while including other string Key values specific to each DataFrame.
    
    Parameters:
        df (pd.DataFrame): The DataFrame to be merged or combined.
    
    Returns:
        pd.DataFrame: Combined DataFrame with columns (PatientDOB, LastName and FirstName) added based on the match/mapping of values starting from KS___ with the integer value in the Key column.
    """

    if 'Key' not in df.columns:
        print("Warning: 'Key' column not found in DataFrame. Skipping processing.")
        return format_df(df)  # Format DataFrame before returning

    # Ensure 'Key' column is of object type to handle mixed types
    df['Key'] = df['Key'].astype(object)

    # Identify datatypes
    df['type'] = df['Key'].apply(lambda x: type(x).__name__)

    # Separate columns based on datatype
    df['integers'] = df['Key'].apply(lambda x: x if isinstance(x, int) else None)
    df['strings'] = df['Key'].apply(lambda x: x if isinstance(x, str) else None)
    
    # Convert 'integers' column to numeric, handling errors by coercing them to NaN
    df['integers'] = pd.to_numeric(df['integers'], errors='coerce').astype('Int64')

    # Ensure 'strings' column only contains non-numeric values
    df['strings'] = df['strings'].apply(lambda x: x if isinstance(x, str) and not x.isnumeric() else None)
    
    # Extract numeric part from strings starting with 'KS'
    df['numeric_part'] = df['strings'].apply(lambda x: ''.join(filter(str.isdigit, x)) if x and x.startswith('KS') else None)
    
    # Create new DataFrames for integers and strings
    df_integers = df.dropna(subset=['integers']).drop(columns=['Key', 'type', 'strings', 'numeric_part'])
    df_strings = df.dropna(subset=['strings']).drop(columns=['Key', 'type', 'integers', 'numeric_part'])
    
    # Create DataFrame with numeric parts starting with 'KS'
    df_numeric = df.dropna(subset=['numeric_part'])
    df_numeric['numeric_part'] = pd.to_numeric(df_numeric['numeric_part'], errors='coerce').astype('Int64')
    df_numeric = df_numeric.drop(columns=['Key', 'type', 'integers', 'strings'])

    # Rename columns to avoid conflicts in merge
    df_integers.rename(columns={'integers': 'Key'}, inplace=True)
    df_numeric.rename(columns={'numeric_part': 'Key'}, inplace=True)

    # Set 'Key' as index for both DataFrames
    df_integers.set_index('Key', inplace=True)
    df_numeric.set_index('Key', inplace=True)

    # Combine the DataFrames
    merged_df = df_numeric.combine_first(df_integers)

    # Reset index to turn 'Key' back into a column
    merged_df.reset_index(inplace=True)

    # Drop duplicates in df_strings while keeping the last occurrence
    df_strings = df_strings.drop_duplicates(subset=['strings'], keep='last')
    # Remove rows where 'strings' starts with 'KS'
    df_strings = df_strings[~df_strings['strings'].str.startswith('KS___', na=False)]
    df_strings.rename(columns={'strings': 'Key'}, inplace=True)

    # Set 'Key' as index for both DataFrames
    df_strings.set_index('Key', inplace=True)
    merged_df.set_index('Key', inplace=True)

    final_df = df_strings.combine_first(merged_df)

    # Reset index to turn 'Key' back into a column
    final_df.reset_index(inplace=True)

    # Drop duplicate columns
    final_df = final_df.loc[:, ~final_df.columns.duplicated()]

    # Format DataFrame before returning (This step is necessary)
    return format_df(final_df)


def save_combined_sheets(combined_sheets, output_file):
    """Save combined DataFrames to an Excel file with multiple sheets."""
    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, df in combined_sheets.items():
            df = process_df(df)  # Process DataFrame before saving (This is necessary step)
            df.to_excel(writer, sheet_name=sheet_name, index=False) # Convert the dataframe to excel file
    return output_file  # Return the path to the saved file


if __name__ == "__main__":
    
    # Get the date from the user
    run_date = input("\nPlease enter the date of the download you made in mmddyy format\n--> ")

    # Specify the paths to the directory
    csv_directory = "//kdhe/dfs/LabShared/Molecular_Genomics_Unit/Testing/PulseNet/PulseNet 2.0/PNExports" 
    xlsx_directory = "//kdhe/dfs/LabShared/Molecular_Genomics_Unit/Testing/PulseNet/PulseNet 2.0/WGS_Databases"
    output_path = "//kdhe/dfs/EPI/LAB_OSE/WGS"

    # Combine all sheets from Excel and CSV files
    combined_sheets = merge_files_to_sheets(csv_directory,xlsx_directory, run_date)
    
    # Save the combined sheets to a new Excel file with the name including run_date
    output_file = os.path.join(output_path, f'{run_date} Epi report past 90.xlsx')
    saved_file = save_combined_sheets(combined_sheets, output_file)

print("The output was saved successfully!")
