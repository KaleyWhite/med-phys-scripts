"""Parse a PDF Daily Charge Report from MOSAIQ (see "T:\Physics - T\Scripts\Data\July Charges" for an example)
Add treatment and simulation charge data to the "Patient Load Statistical Analysis.xlsx" spreadsheet

Assumptions
-----------
No report will be added to the spreadsheet multiple times. In other words, the script does not check for duplicate entries (there isn't a reasonable way to do so).
"""

# Imports
import re  # For regex

import fitz  # PyMuPDF (for parsing PDFs)
import numpy as np
import pandas as pd


# Filepaths
t_path = r"\\vs20filesvr01\groups\CANCER\Physics - T"  # Physics folder on the T drive
input_filepath = r"{}\Scripts\Data\July Charges.pdf".format(t_path)  # PDF to parse
output_filepath = r"{}\Scripts\Output Files\PatientLoadStatisticalAnalysis\Patient Load Statistical Analysis.xlsx".format(t_path)  # Complate filename of output spreadsheet

# Each row starts with all None's
# "Date" and "MD" are required in each row
# For treatment charges, "Tx Type" and "Tx Machine" are filled in, but "Sim Type" and "Sim Machine" are not
# For sim charges, "Sim Type" and "Sim Machine" are filled in, but "Tx Type" and "Tx Machine" are not
# "Date": Date of the charge. "Date" column in PDF
# "MD": Patient's MD. "MD" column in PDF
# "Tx Type": Treatment type. One of "Simple", "Intermediate", "Complex", "IMRT S", "IMRT C", and "SBRT/SRS". Based on "Description" column in PDF
# "Tx Machine": Machine on which the tx occurred. One of "Tomo", "E1", and "E2". Based on "Loc" column in PDF
# "Sim Type": Type of simulation. One of "Boost", "Elekta", and "IMRT". Based on "Description" column in PDF
# "Sim Machine": Machine on which the sim'd patient will be treated. One of "Tomo", "E1", and "E2". This is NOT (necessarily) the machine on which the patient was sim'd. Based on "Description" column in PDF
data = {colname: [] for colname in ["Date", "MD", "Tx Type", "Tx Machine", "Sim Type", "Sim Machine"]}

# Parse PDF
with fitz.open(input_filepath) as pdf:
    for pg in list(pdf)[:-1]:  # Ignore last page
        text = pg.getText().split("\n")[27:-4]  # Remove extraneous info

        # Divide text into rows according to patient names
        # This is necessary because not ecery row contains the same number of values
        # Patient name is in the format "last, first (MR#)"
        pt_name_idxs = [i for i, val in enumerate(text) if re.search("\(000\d{6}\)", val)]

        # Iterate over all rows
        for i, idx in enumerate(pt_name_idxs):
            # We only care abot rows that have a "Loc" value (comes immediately after patient name)
            # These rows contain 12 fields, so we only parse these
            if (i == len(pt_name_idxs) - 1 and idx + 12 == len(text) - 1) or (i != len(pt_name_idxs) - 1 and idx + 12 == pt_name_idxs[i + 1]):
                row = {colname: None for colname in data}  # Start with all None's

                # Fields used for both tx and sim charges
                row["MD"] = text[idx + 2]
                row["Date"] = text[idx + 9]

                loc = text[idx + 1]  # "Loc" field
                chg = text[idx + 7]  # "Description" fiels
                
                # Is it a sim charge?
                if chg in ["Tomo Sim", "IMRT SIM (Tomo)"]:
                    row["Sim Type"] = "IMRT"
                    row["Sim Machine"] = "Tomo"
                elif chg in ["Elekta Sim", "77290 Sim: C", "77285 Sim: I", "77280 Sim: S"]: 
                    row["Sim Type"] = "Elekta"
                    row["Sim Machine"] = "Elekta"
                elif chg == "77290: Sim: Boost":  
                    row["Sim Type"] = "Boost"
                    row["Sim Machine"] = "Elekta"
                else:  # Tx, not sim
                    # Tx machine. Loc is "TOM", "E1", or "E2"
                    if loc == "TOM":
                        row["Tx Machine"] = "Tomo"
                    else:
                        row["Tx Machine"] = "Elekta"
                    
                    # Tx type
                    if chg == "Daily IMRT: S tx Del":
                        row["Tx Type"] = "IMRT S"
                    elif chg == "Daily IMRT: C Tx Del":
                        row["Tx Type"] = "IMRT C"
                    elif chg == "Daily SBRT tx Delive":
                        row["Tx Type"] = "SBRT/SRS"
                    elif chg == "Rad Tx: Simple Daily":
                        row["Tx Type"] = "Simple"
                    elif chg == "Rad Tx Int Daily":
                        row["Tx Type"] = "Intermediate"
                    elif chg == "Rad Tx: Simple Daily":
                        row["Tx Type"] = "Complex"
            
            # Add row data to `data`
            for colname, val in row.items():
                data[colname].append(val)

# Append data to spreadsheet
data = pd.DataFrame(data)  # Convert dictionary to DataFrame
data["Day of Week"] = data["Date"].astype(np.datetime64).dt.day_name  # Add "Day of Week" column (values based on "Date" column)
with pd.ExcelWriter(output_filepath, mode="a", engine="openpyxl") as ew:
    ew.sheets = dict((ws.title, ws) for ws in ew.book.worksheets)  # A "hack" line that lets us append to an existing sheet
    start_row = ew.book["Data"].max_row  # Write new data BELOW, not OVER, existing data
    data.to_excel(ew, sheet_name="Data", startrow=start_row, header=False, index=False) 
    ew.save() 
