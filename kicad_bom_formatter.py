#!/usr/bin/python3

import sys
import pandas as pd
import argparse

helptext_formatting = """
Input file does not match expected formatting. 
Input CSV file is expected to follow the format:

    Item, Quantity, Description, Part Number, Part Reference, DEBUGPART

Create this in KiCad using 'Tools > Generate BOM', and selecting the 
'csv_bom_riverloop.py' export script.
"""

description="Creates purchasing BOMs from KiCAD individual BOM."
parser = argparse.ArgumentParser(description=description)

helptext = "A BOM file exported from KiCAD using csv_export_riverloop.py"
parser.add_argument('bom_file', help=helptext)

def file_readin(filename):
    """Reads in a CSV, returns pandas dataframe. 

    Args:
        filename (str): the file name string

    Returns:
        df (pandas dataframe): the input file read into a pandas dataframe object
    
    """
    filename_in  = filename
    df = pd.read_csv(filename_in)
    return df

def rename_columns(df):
    """Renames column headers to format that Arena expects. Sanitizes dataframe fields.

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of renamed, sanitized BOM data

    """
    df.columns = [ 'Level', 'Quantity', 'Description_Ignore', 'Item_Number', 'Reference_Designator', 'DEBUGPART_Ignore']
    df['Item_Number']      = df['Item_Number'].fillna(value='0')
    df['DEBUGPART_Ignore'] = df['DEBUGPART_Ignore'].fillna(value=0)
    try:
        df['DEBUGPART_Ignore'] = df['DEBUGPART_Ignore'].astype(int)
    except:
        print(df[ df['DEBUGPART_Ignore'] != 0 ])
        exit("An unexpected value showed up in the 'DEBUGPART' field. Please check input file.")
    df['Item_Number']      = df['Item_Number'].astype(str)
    return df

def scrub_noload(df):
    """Removes all noload entries from the BOM dataframe

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, with NOLOAD components removed

    """
    return df[~df.Item_Number.str.startswith("N")]

def scrub_dnps(df):
    """Removes all DNP entries from the BOM dataframe

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, with "Pop=DNP" components removed

    """
    return df[df.Pop != ""]

def scrub_mtgholes(df):
    """Removes all mounting holes from the BOM 

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, with mounting holes removed

    """
    return df[~df.RefDes.str.startswith('MTG') ]

def scrub_fiducials(df):
    """Removes all fiducials from the BOM 

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, with fiducials removed

    """
    return df[~df.RefDes.str.startswith('FID') ]

def scrub_testpoints(df):
    """Removes all testpoints from the BOM 

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, with testpoints removed

    """
    return df[~df.RefDes.str.startswith('TP')]
	
def scrub_shortedpads(df):
    """Removes all shorted pad components (item number 157) from the BOM 

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, with shorted pads removed

    """
    return df[~df.Item_Number.str.startswith('157')]

def level_and_sort(df):
    """Sorts BOM dataframe into the proper order for an Arena import

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, organized with 930 first, cmpts second, 127 third, Debug BOM last

    """
    df['Level'] = 1
    df.loc[df.Item_Number.str.startswith('930'), 'Level'] = 0
    df.loc[df.Item_Number.str.startswith('127'), 'Level'] = 2
    df.loc[df.DEBUGPART_Ignore == 1, 'Level'] = 3
    df = df.sort_values(by=['Level', 'Item_Number'])
    df.loc[df.Item_Number.str.startswith('127'), 'Level'] = 1
    df.loc[df.DEBUGPART_Ignore == 1, 'Level'] = 2
    return df

def export_production_bom(df):
    """Changes quantity of the Debug BOM part (PN starting with 127-*) to Zero

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, with Debug BOM part quantity set to Zero

    """
    df.loc[df.Item_Number.str.startswith('127'), 'Quantity'] = 0
    return df

def collate_by_mfgpn(df):
    """Collates dataframe by MfgPN column for purchasing; adds Qty

    Args:
        df (pandas dataframe): dataframe of BOM data

    Returns:
        df (pandas dataframe): dataframe of BOM, with RefDes'es collated by MfgPN 
    """
    df = df.groupby('MfgPN').RefDes.agg(
            [
                ('Qty', 'count'), 
                ('RefDes', ', '.join)
            ])
    print(df)

def collate_by_distributor(distributor, df):
    """Collates dataframe by a given component distributor

    Args:
        distributor (string): string of the distributor (e.g. "Mouser", "Digikey")
        df (pandas dataframe): dataframe of BOM data

    Returns:
        (none - no return function)

    """
    return df

def file_writeout(filename, df):
    """Writes BOM dataframe into an output CSV file, ready to import into Arena

    Args:
        filename     (string): string of the input file, which is modified into the correct output file name  
        df (pandas dataframe): dataframe of BOM data

    Returns:
        (none - no return function)

    """
    filename_out = filename.rstrip('.csv') + '_Scrubbed.' + 'csv'
    print(filename_out)
    df.to_csv(path_or_buf=filename_out, index=False)


if __name__ == '__main__':
    args = parser.parse_args()
    input_bom_file = args.bom_file

    try:
        df = file_readin(input_bom_file)
        df = rename_columns(df)
    except FileNotFoundError:
        sys.exit(f"Input file '{input_bom_file}' not found. Exiting.") 
    except UnicodeDecodeError:
        sys.exit("Input file type must be flat .csv (no Unicode). Exiting.") 
    except ValueError:
        sys.exit(helptext_formatting)
    else:
        df = scrub_dnps(df)
        df = scrub_mtgholes(df)
        df = scrub_testpoints(df)
        df = scrub_shortedpads(df)
        df = level_and_sort(df)
        if prod_export:
            df = export_production_bom(df)
        file_writeout(input_bom_file, df)




