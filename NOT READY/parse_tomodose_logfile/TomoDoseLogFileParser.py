"""Write TomoDose logfile data to CSV files

Logfiles are at "T:\Physics - T\QA & Procedures\Tomo\TomoDose\TomoDose Logfile Backup" and organized into year and month folders according to TomoDoseDirectoryStructure script
Data field names are determined from the "dummy" logfile in "T:\Physics - T\Scripts\Data"
If output file already exist, only add data from logfiles with datetime after the latest datetime in the output files

There are three output files
- "TomoDose Log File Data - Other": Data fields in the first ~50 rows of a logfile. E.g., Energy, Cal file, Dose
- "TomoDose Log File Data - Array": Data from the detector arrays, including both X- and Y- coordinates
- "TomoDose Log File Data - Movie Frame": Detector array data at each frame
Each sheet has an added "Datetime" column that associates rows across CSV files
"""

import os
import sys
from datetime import datetime
from pathlib import Path
from re import match
from tkinter import Tk, messagebox
from typing import OrderedDict
from tkinter import Tk

import numpy as np
import pandas as pd
from tkfilebrowser import askopendirname


Tk().withdraw()
input_folder = askopendirname(title='Choose Logfile Location')
output_folder = askopendirname(title='Choose Output File Location')

other_data_path = os.path.join(output_folder, "TomoDose Log File Data - Other.csv")
array_data_path = os.path.join(output_folder, "TomoDose Log File Data - Array.csv")
movie_frame_data_path = os.path.join(output_folder, "TomoDose Log File Data - Movie Frame.csv")

# Get field names from first logfile
logfile_found = False
for file in Path(input_folder).glob('**/*.txt'):
    with open(file) as f:
        lines = f.readlines()

        try:
            ## "Other" DataFrame

            # "Other" data ends at first blank line
            first_blank_line = lines.index("\n")

            # Skip first 3 (info) lines
            # Remove colon from end of field names
            # Ignore "Date" and "Time" fields because we use "Datetime" field instead
            other_colnames = ["Datetime"] + [line.split("\t")[0].replace(":", "") for line in lines[3:first_blank_line] if not line.startswith("Date") and not line.startswith("Time")]
            other_data = OrderedDict([(colname, []) for colname in other_colnames])

            ## "Array" DataFrame

            # "Array" data ends at next blank line
            blank_line = lines.index("\n", first_blank_line + 1)

            # Column names come from 2 rows: first row is name, second row is unit
            # Row below the blank is array "header" info
            # Array has 8 columns
            array_colnames = ["{} ({})".format(first_part, lines[blank_line + 3].split("\t")[1:9][i]) for i, first_part in enumerate(lines[blank_line + 2].split("\t")[1:9])]
            
            # Replace "Coordinate" field with "X" and "Y" fields
            array_colnames.remove("Coordinate (mm)") 
            array_colnames.insert(1, "X (mm)")
            array_colnames.insert(2, "Y (mm)")

            # Add "Datetime" and "Array" (for array number) columns as well
            array_data = OrderedDict([(colname, []) for colname in ["Datetime", "Array"] + array_colnames])

            ## "Movie Frame" DataFrame

            # "Movie Frame" data starts at last blank line
            last_blank_line = len(lines) - lines[::-1].index("\n") - 1

            # Colnames start on line below the blank
            movie_frame_colnames = lines[last_blank_line + 1].split("\t")[1:]
            movie_frame_data = OrderedDict([(colname, []) for colname in ["Datetime"] + movie_frame_colnames])

            logfile_found = True
            break
        except (ValueError, IndexError):
            continue

if not logfile_found:
    Tk().withdraw()
    messagebox.showerror("No Logfiles", "There are no logfiles in the selected input directory.\nClick OK to abort the script.")
    sys.exit(1)

# Read in old data and determine latest year and month in the data so we can ignore folders from earlier than these
try:
    old_other_data = pd.read_csv(other_data_path)
    old_other_data["Datetime"] = old_other_data["Datetime"].astype(np.datetime64)
    
    old_array_data = pd.read_csv(array_data_path)
    old_array_data["Datetime"] = old_array_data["Datetime"].astype(np.datetime64)

    old_movie_frame_data = pd.read_csv(movie_frame_data_path)
    old_movie_frame_data["Datetime"] = old_movie_frame_data["Datetime"].astype(np.datetime64)

    max_dt = max(old_other_data["Datetime"])
    max_yr, max_mo = max_dt.year, max_dt.month
except FileNotFoundError:
    max_dt = datetime(1970, 1, 1)  # Oldest possible datetime value
    max_yr = max_mo = 0  # Every actual year or month will be >0

# Populate DataFrames w/ logfile data
for yr in os.listdir(input_folder):
    if int(yr) < max_yr:  # Assume logfile data has already been added to old data
        continue
    yr_path = r"{}\{}".format(log_file_path, yr)
    for mo in listdir(yr_path):
        if int(yr) == max_yr and int(mo) < max_mo:  # Assume logfile data has already been added to old data
            continue
        mo_path = r"{}\{}".format(yr_path, mo)
        for logfile in listdir(mo_path):
            date = None

            logfile_path = r"{}\{}".format(mo_path, logfile)
            with open(logfile_path, "r") as f:
                lines = f.readlines()

                # Get datetime from logfile
                date = [line for line in lines if line.startswith("Date")][0].strip("\n").split("\t")[1]
                date = "/".join([num.zfill(2) for num in date.split("/")])
                time = [line for line in lines if line.startswith("Time")][0].strip("\n").split("\t")[1]
                time = ":".join([num.zfill(2) for num in time.split(":")])
                dt = "{} {}".format(date, time)
                dt = datetime.strptime(dt, "%m/%d/%Y %H:%M:%S")
                if dt <= max_dt:  # Assume logfile data has already been added to old data
                    continue

                # "Other" data
                other_data["Datetime"].append(dt)
                first_blank_line = lines.index("\n")
                for line in lines[3:first_blank_line]:
                    other_colname, val = line.split("\t")[:2]
                    other_colname = other_colname.replace(":", "")
                    val = val.strip("\n")
                    if other_colname not in ["Date", "Time"]:
                        other_data[other_colname].append(val)

                # "Movie Frame" data
                last_blank_line = len(lines) - lines[::-1].index("\n") - 1
                for line in lines[(last_blank_line + 2):]:
                    movie_frame_data["Datetime"].append(dt)
                    for i, val in enumerate(line.strip("\n").split("\t")[1:]):
                        movie_frame_data[movie_frame_colnames[i]].append(val)

                # "Array" data
                line_idx = first_blank_line
                while line_idx < last_blank_line:  # Array data is next-to-last section in file
                    if lines[line_idx] == "\n":  # Blank line, so starting a new subsection of arrays
                        line_idx += 1  # Skip the blank line

                        array_name_line = lines[line_idx].strip("\n").split("\t")  # Array name and coordinate info
                        col_hdrs_line = lines[line_idx + 1].strip("\n").split("\t")  # Array data column headers
                        
                        for col_idx in range(0, (len(col_hdrs_line) - col_hdrs_line.count("")) // 8):  # There is a blank column after each array in the section. There are 8 columns in an array section.
                            col_idx = 9 * col_idx + 1
                            array_line_idx = line_idx + 3  # "Temp" line_idx variable
                            array_data_line = lines[array_line_idx].strip("\n").split("\t")[col_idx:]  # Array data starts 3 rows below array name and coordinate info
                            
                            # Array 6b's first data row is blank
                            if not array_data_line:
                                array_line_idx += 1
                                array_data_line = lines[array_line_idx].strip("\n").split("\t")[col_idx:]  # Array data starts 3 rows below array name and coordinate info

                            array_hdr = match("Array (\d[a-z]?): Integrated dose on ([XY]) axis: [XY] = ([+\-]?\d+) mm", array_name_line[col_idx])
                            array = array_hdr.group(1)  # Array name ("1", "2", "3a", "3b", "4a", "4b", "5a", "5b", "6a", or "6b")

                            while array_data_line:  # Parse data rows until we hit a blank line in this array
                                array_data["Datetime"].append(dt)
                                array_data["Array"].append(array_hdr.group(1))
                                array_data["Detector # (#)"].append(array_data_line[0].strip())

                                if array_hdr.group(2) == "X":  # "Coordinate" column is y-coordinate
                                    array_data["X (mm)"].append(array_data_line[1])
                                    array_data["Y (mm)"].append(array_hdr.group(3).lstrip("+"))
                                else:  # "Coordinate" column is x-coordinate
                                    array_data["Y (mm)"].append(array_data_line[1])
                                    array_data["X (mm)"].append(array_hdr.group(3).lstrip("+"))

                                for i in range(2, 8):  # Remaining array columns afetr "Coordinate"
                                    array_data[array_colnames[i + 1]].append(array_data_line[i])

                                # Move to next line in the array, but DO NOT move to next line overall (line_idx)
                                array_line_idx += 1  
                                array_data_line = lines[array_line_idx].strip("\n").split("\t")[col_idx:]
                        
                    line_idx = lines.index("\n", line_idx + 1)  # # We've parsed every array that starts at row `line_idx`, so jump to next blank line, below which next arrays start

# Exit script if no new data
if not any(other_data.values()):
    Tk().withdraw()
    messagebox.showinfo("No New Data", "There is no new logfile data. The logfile data spreadsheet was not changed.")
    sys.exit(1)

### Convert dictionaries to DataFrames and clean up

## Other data

# Dictionary -> DataFrame
other_data = pd.DataFrame(other_data)

# Move energy unit from column value to column name
other_data.columns = [colname if colname != "Energy" else "Energy (MV)" for colname in other_data.columns]
other_data["Energy (MV)"] = [val.strip(" (MV)") for val in other_data["Energy (MV)"]]

# Combine with old data
if max_yr != 0:  # `old` DataFrames exist
    other_data = old_other_data.append(other_data, ignore_index=True)

# Sort by datetime
other_data.sort_values("Datetime", inplace=True)

## Array data

# Dictionary -> DataFrame
array_data = pd.DataFrame(array_data)

# Remove "#" unit from "Detector # (#)" column name and "Cnts" word from "Corrected Cnts (Counts)" column name
array_data.columns = ["Detector #" if colname == "Detector # (#)" else "Corrected (Counts)" if colname == "Corrected Cnts (Counts)" else colname for colname in array_data.columns]

# Combine with old data
if max_yr != 0:  # `old` DataFrames exist
    array_data = old_array_data.append(array_data, ignore_index=True)

# Sort by datetime, array name, and detector number
array_data["Detector #"] = array_data["Detector #"].astype(int)
array_data.sort_values(["Datetime", "Array", "Detector #"], inplace=True)
#array_data["Detector #"] = array_data["Detector #"].astype(str)

## Movie frame data

# Dictionary -> DataFrame
movie_frame_data = pd.DataFrame(movie_frame_data)
if max_yr != 0:  # `old` DataFrames exist
    movie_frame_data = old_movie_frame_data.append(movie_frame_data, ignore_index=True)

# Sort by datetime, time, and beam-on time
movie_frame_data[["Current Time", "Beam On Time", "Exposure Period"]] = movie_frame_data[["Current Time", "Beam On Time", "Exposure Period"]].astype(int)
movie_frame_data.sort_values(["Datetime", "Current Time", "Beam On Time"], inplace=True)
#movie_frame_data[["Current Time", "Beam On Time", "Exposure Period"]] = movie_frame_data[["Current Time", "Beam On Time", "Exposure Period"]].astype(str)

# Write data to CSV files
other_data.to_csv(other_data_path, index=False)
array_data.to_csv(array_data_path, index=False)
movie_frame_data.to_csv(movie_frame_data_path, index=False)
