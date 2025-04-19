"""
    Microscopy Data Organiser
    
    Processes .CSV files from microscopy analysis by categorising them into
    Excel sheets based on binary variables and generating summary statistics.
    
    It supports multiple input files and creates unique output Excel files 
    for each.


"""

#
# Setup
#

import os
import sys
import logging
import argparse
import pandas as pd
from openpyxl import load_workbook
VERSION = "2.0.0"

# Set up logging
logging.basicConfig(level=logging.WARNING)


#
# Parse command line arguments for the script
# 
# Provide one or more input .CSV files
# Specify columns to summarise using the -c option
# 
# Reutrns parsed command line arguments
#

def load_commandline():
    parser = argparse.ArgumentParser(description='Microscopy Data Organiser')

    # Input file argument
    helpstr = 'Input file(s) from microscopy analysis'
    parser.add_argument('input', type=str, nargs='+', help=helpstr)

    # Remove rows where Children_Nuclei_Count is 0
    helpstr = 'Remove rows where Children_Nuclei_Count is 0'
    parser.add_argument('-r', '--remove-zero-nuclei', 
                        action='store_true', help=helpstr)

    # Version information
    helpstr = 'Display version and exit'
    out = f"This is\n Patrick Jane version {VERSION}"
    parser.add_argument('-V', '--version', action='version', help=helpstr,
                        version=out)

    # Select columns for summary
    helpstr = 'Comma-separated list of columns to summarise (default: AreaShape_Area, Children_Nuclei_Count)'
    parser.add_argument(
        '-c', '--columns',
        nargs='+',
        default=["AreaShape_Area", "Children_Nuclei_Count"],
        help="List of columns to summarise (separated by spaces)"
        )

    args = parser.parse_args()

    return (args)


#
# Create unique output names
#


used_sheet_names = set()


def get_unique_sheet_name(name, max_length=31):
    base_name = name[:max_length]
    unique_name = base_name
    counter = 1

    # Avoid overwriting if output with same name exists
    while unique_name in used_sheet_names:
        suffix = f"_{counter}"
        unique_name = (base_name[:max_length - len(suffix)] + suffix)
        counter += 1

    used_sheet_names.add(unique_name)
    return unique_name

#
# Defining the Jane function to make the summary pages
#


def jane(outfile, summary_columns):
    try:
        # Check if there is an output from the Patrick function
        if not os.path.exists(outfile):
            print(f"File '{outfile}' does not exist.")
            return

        file_size = os.path.getsize(outfile)
        print("")
        print(f"File '{outfile}' exists. Size: {file_size} bytes. I know who the killer is I'm just not going to tell you for another 45-50 minutes")
        print("")

        # Load the output from Patrick
        try:
            print("")
            print("Successfully loaded the workbook. Using the powers of the mind to summarise the categories you need.")
            print("")
        except Exception as e:
            print(f"Error loading workbook with openpyxl: {str(e)}")
            return

        xls = pd.ExcelFile(outfile, engine='openpyxl')

        # Summary data containers
        summary_data = {col: [] for col in summary_columns}

        # Iterate through the sheets
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)

            for col in summary_columns:
                if col in df.columns:
                    summary_data[col].append(df[col].rename(sheet_name))

        # Error checking if no data is found for specific columns
        for col, data in summary_data.items():
            if not data:
                print(f"No data found in any sheet for column '{col}'.")

        # Create summary DataFrames
        summary_dfs = {
            col: pd.concat(data, axis=1) if data else None
            for col, data in summary_data.items()
        }

        # Load existing sheets to avoid overwriting them
        with pd.ExcelWriter(outfile, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            for col, summary_df in summary_dfs.items():
                if summary_df is not None:
                    # Truncate the sheet name so openpxyl doesn't get angry
                    sheet_name = get_unique_sheet_name(f"{col} Summary")
                    summary_df.to_excel(writer, sheet_name=sheet_name, index=False)

        print("")
        print("Successfully created 'Area Summary' and 'Nuclei Summary' sheets. I'm going to sit in a derelict attic now instead of doing police work")
        print("")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        sys.exit(1)

#
# The Patrick function
#

# Parse command-line arguments
args = load_commandline()

# Parse the columns to summarise (split by commas)
summary_columns = args.columns if args.columns else ["AreaShape_Area", "Children_Nuclei_Count"]

# Is there an input file in the command?
if len(args.input) < 1:
    logging.error("Please provide the input CSV file name as a command-line argument.")
    sys.exit(1)

for input_file in args.input:  # Loop through all input files

    # Check if the file is real
    if not os.path.exists(input_file):
        logging.error(f"Input file '{input_file}' not found. Skipping.")
        continue

    try:
        # Read the CSV file, take the first row as the header
        df = pd.read_csv(input_file, header=0)

        # Strip whitespace
        df.columns = df.columns.str.strip()
        
        # Remove rows where Children_Nuclei_Count is 0 (if option is specified)
        if args.remove_zero_nuclei:
            if "Children_Nuclei_Count" in df.columns:
                initial_row_count = len(df)
                df = df[df["Children_Nuclei_Count"] != 0]
                final_row_count = len(df)
            print(f"Removed {initial_row_count - final_row_count} rows where Children_Nuclei_Count was 0. They were not of use to me just like how I have no need for concrete evidence in my cases")
        else:
            print("Column 'Children_Nuclei_Count' not found. Skipping row removal.")

        # Required columns
        required_columns = ['Classify_Mononucleated', 'Classify_Infected'] + summary_columns

        # Define the columns to keep in the output sheets
        columns_to_keep = ['ObjectNumber', 'Classify_Mononucleated', 'Classify_Infected'] + summary_columns

        # Check if all columns are present
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise KeyError(f"The following columns are missing: {','.join(missing_columns)}")

        # Categorize the data according to infection and nuclei
        categories = {
            'Mononucleated_Infected': df[(df['Classify_Mononucleated'] == 1) & (df['Classify_Infected'] == 1)][columns_to_keep],
            'Multinucleated_Infected': df[(df['Classify_Mononucleated'] == 0) & (df['Classify_Infected'] == 1)][columns_to_keep],
            'Mononucleated_Uninfected': df[(df['Classify_Mononucleated'] == 1) & (df['Classify_Infected'] == 0)][columns_to_keep],
            'Multinucleated_Uninfected': df[(df['Classify_Mononucleated'] == 0) & (df['Classify_Infected'] == 0)][columns_to_keep]
        }

        # Create the output name
        input_filename = os.path.basename(input_file)
        input_name_no_ext, input_ext = os.path.splitext(input_filename)
        outfile = f"output_{input_name_no_ext}.xlsx"

        # Don't overwrite the output file!
        counter = 1
        original_outfile = outfile
        while os.path.exists(outfile):
            outfile - f"{original_outfile[:-4]}_{counter}.xlsx"
            counter += 1

        # Write categorized data to Excel
        used_sheet_names = set()

        def get_unique_sheet_name(name, max_length=25):
            base_name = name[:max_length]
            unique_name = base_name
            counter = 1

            while unique_name in used_sheet_names:
                suffix = f"_{counter}"
                unique_name = (base_name[:max_length - len(suffix)] + suffix)
                counter += 1

            used_sheet_names.add(unique_name)
            return unique_name

        with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
            for sheet_name, data in categories.items():
                # Truncate and make unique
                sheet_name = get_unique_sheet_name(sheet_name)
                data.to_excel(writer, sheet_name=sheet_name, index=False)

        # Make sure that the file exists
        if not os.path.exists(outfile):
            raise FileNotFoundError(f"Failed to create the output file: {outfile}")

        print("")
        print(f"Writing output data to {outfile}. Did you kill her be honest.")
        print("")

        # Execute the Jane function!
        jane(outfile, summary_columns)

#
# Error messages
#

    except pd.errors.EmptyDataError:
        logging.error(f"The file '{input_file}' is empty :o")
        sys.exit(1)
    except KeyError as e:
        logging.error(f"Column not found in '{input_file}': {e} :S")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An error processing '{input_file}' {e} :|")
        sys.exit(1)
