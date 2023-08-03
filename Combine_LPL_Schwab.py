import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime

LPLFile = pd.read_excel('2023Q2_ICA_GROUP_WEALTH_MGT_LLC_13FReportbyCUSIP_RIA_20230705.xlsx')
SchwabFile = pd.read_csv('orion_data.csv')

# Data Cleaning and Manipulation for LPL data
def concat_cusip_with_proxy(row):
    if pd.notnull(row['Proxy Authority']) and row['Proxy Authority'] in ['Y', 'N']:
        return row['CUSIP'] + row['Proxy Authority']
    else:
        return row['CUSIP']

LPLFile['CUSIP'] = LPLFile.apply(concat_cusip_with_proxy, axis=1).astype(str)

LPLFile['Sole'] = LPLFile['Number of Shares/Contracts'].where(LPLFile['CUSIP'].str.endswith('Y'), 0)
LPLFile['None'] = LPLFile['Number of Shares/Contracts'].where(~LPLFile['CUSIP'].str.endswith('Y'), 0)
LPLFile['CUSIP'] = LPLFile['CUSIP'].str.rstrip('YN')
LPLFile['CUSIP'] = LPLFile['CUSIP'].astype(str).str.strip()  # Convert CUSIP to string data type

LPLFile['Name of Issuer'] = LPLFile['Security Name']
LPLFile['Title of Class'] = LPLFile['Security Type']

LPLFile['Aggregate Value'] = LPLFile.groupby('CUSIP')['Aggregate Value (to the nearest $)'].transform('sum')
LPLFile['Sole'] = LPLFile.groupby('CUSIP')['Sole'].transform('sum')
LPLFile['None'] = LPLFile.groupby('CUSIP')['None'].transform('sum')
LPLFile['Number of Shares'] = LPLFile.groupby('CUSIP')['Number of Shares/Contracts'].transform('sum')

columns_to_round = ['Aggregate Value', 'Sole', 'None', 'Number of Shares']
LPLFile[columns_to_round] = LPLFile[columns_to_round].round()

LPLFile.drop_duplicates(subset='CUSIP', keep='first', inplace=True)


# Data Cleaning and Manipulation for Schwab data
SchwabFile = SchwabFile[['Investment Discretion', 'AssetShares', '13FCusip', 'AssetValue', 'Product Description', 'ProductType']]

SchwabFile.rename(columns={'13FCusip': 'CUSIP', 'AssetShares': 'AssetShare'}, inplace=True)

SchwabFile['AssetValue'] = SchwabFile['AssetValue'].round()
SchwabFile['AssetShare'] = SchwabFile['AssetShare'].round()


SchwabFile['Sole'] = SchwabFile['AssetShare'].where(SchwabFile['CUSIP'].str.endswith('Y'), 0)
SchwabFile['None'] = SchwabFile['AssetShare'].where(~SchwabFile['CUSIP'].str.endswith('Y'), 0)
SchwabFile['CUSIP'] = SchwabFile['CUSIP'].str.rstrip('YN')
SchwabFile['CUSIP'] = SchwabFile['CUSIP'].astype(str).str.strip()   # Convert CUSIP to string data type

SchwabFile['Aggregate Value'] = SchwabFile.groupby('CUSIP')['AssetValue'].transform('sum')
SchwabFile['Sole'] = SchwabFile.groupby('CUSIP')['Sole'].transform('sum')
SchwabFile['None'] = SchwabFile.groupby('CUSIP')['None'].transform('sum')
SchwabFile['Number of Shares'] = SchwabFile.groupby('CUSIP')['AssetShare'].transform('sum')

columns_to_round = ['Aggregate Value', 'Sole', 'None', 'Number of Shares']
SchwabFile[columns_to_round] = SchwabFile[columns_to_round].round()

SchwabFile['Name of Issuer'] = SchwabFile['Product Description']
SchwabFile['Title of Class'] = SchwabFile['ProductType']

SchwabFile.drop_duplicates(subset='CUSIP', keep='first', inplace=True)

# Create 'Investment Discretion' column in LPLFile and fill with 'SOLE'
LPLFile['Investment Discretion'] = 'SOLE'

# Combine LPLFile and SchwabFile based on CUSIP
merged_data = pd.concat([LPLFile, SchwabFile], ignore_index=True)

# Check for duplicate CUSIP values in the conjoined DataFrame
duplicates = merged_data[merged_data.duplicated(subset='CUSIP', keep=False)]
if not duplicates.empty:
    # Perform math for duplicate CUSIP values
    merged_data = merged_data.groupby('CUSIP').agg({
        'Name of Issuer': lambda x: x.iloc[0] if pd.notnull(x.iloc[0]) else x.iloc[1] if len(x) > 1 and pd.notnull(x.iloc[1]) else x.iloc[0],
        'Title of Class': lambda x: x.iloc[0] if pd.notnull(x.iloc[0]) else x.iloc[1] if len(x) > 1 and pd.notnull(x.iloc[1]) else x.iloc[0],
        'FIGI': 'first',
        'Aggregate Value': 'sum',
        'Sole': 'sum',
        'None': 'sum',
        'Number of Shares': 'sum',
        'Investment Discretion': 'first'
    }).reset_index()

# Add missing columns with specified values
merged_data['Shares/Principal'] = 'SH'
merged_data['Put/Call'] = ''
merged_data['Other Managers'] = 0
merged_data['Shared'] = 0

# Round all the numeric columns to the nearest whole value again after the math
columns_to_round = ['Aggregate Value', 'Sole', 'None', 'Number of Shares']
merged_data[columns_to_round] = merged_data[columns_to_round].round()

# Select the desired columns
selected_columns = [
    "Name of Issuer", "Title of Class", "CUSIP", "FIGI", "Aggregate Value",
    "Number of Shares", "Shares/Principal", "Put/Call", "Investment Discretion",
    "Other Managers", "Sole", "Shared", "None"
]

# Extract the desired columns from the merged data
final_data = merged_data[selected_columns]
# Write combined and cleaned data to a new CSV file
output_file = 'LPL_and_Schwab_conjoined_13F.xlsx'
final_data.to_excel(output_file, index=False)

print("Data exported to:", output_file)

# Create the root element for the XML with a default namespace
ns = "http://www.sec.gov/edgar/document/thirteenf/informationtable"
root = ET.Element("informationTable", xmlns=ns)

# Create the main infoTable element
info_table = ET.SubElement(root, "infoTable")

# Iterate through the rows of the 'final_data' DataFrame
for index, row in final_data.iterrows():
    # Create an infoTableEntry element for each row
    info_table_entry = ET.SubElement(info_table, "infoTableEntry")

    # Add sub-elements to the 'infoTableEntry' element based on column values
    ET.SubElement(info_table_entry, "nameOfIssuer").text = row["Name of Issuer"]
    ET.SubElement(info_table_entry, "titleOfClass").text = row["Title of Class"]
    ET.SubElement(info_table_entry, "cusip").text = row["CUSIP"]
    ET.SubElement(info_table_entry, "value").text = str(int(row["Aggregate Value"]))

    # Create shrsOrPrnAmt element and add sub-elements for sshPrnamt and sshPrnamtType
    shrs_or_prn_amt = ET.SubElement(info_table_entry, "shrsOrPrnAmt")
    ET.SubElement(shrs_or_prn_amt, "sshPrnamt").text = str(int(row["Number of Shares"]))
    ET.SubElement(shrs_or_prn_amt, "sshPrnamtType").text = row["Shares/Principal"]

    ET.SubElement(info_table_entry, "investmentDiscretion").text = row["Investment Discretion"]
    ET.SubElement(info_table_entry, "otherManager").text = str(int(row["Other Managers"]))

    # Create votingAuthority element and add sub-elements for Sole, Shared, and None
    voting_authority = ET.SubElement(info_table_entry, "votingAuthority")
    ET.SubElement(voting_authority, "Sole").text = str(int(row["Sole"]))
    ET.SubElement(voting_authority, "Shared").text = str(int(row["Shared"]))
    ET.SubElement(voting_authority, "None").text = str(int(row["None"]))

# Create the XML tree
tree = ET.ElementTree(root)

# Get the current month and year
current_month = datetime.now().strftime("%m")
current_year = datetime.now().strftime("%Y")

# Format the filename
xml_output_file = f"SECForm13F_{current_month}-{current_year}.xml"
with open(xml_output_file, "w", encoding="utf-8") as f:
    # Custom XML declaration with standalone attribute
    xml_declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f.write(xml_declaration)
    # Write the tree to the file
    tree.write(f, encoding="unicode")

print("XML data exported to:", xml_output_file)
