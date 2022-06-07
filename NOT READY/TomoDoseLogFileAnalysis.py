"""Analyze and visualize TomoDose logfile data

For the given logfile, lot dose at X and Z coordinates for all arrays (Y coordinate is always zero).
"""

from re import match
from typing import OrderedDict

import matplotlib.pyplot as plt
import pandas as pd


filename = "7-Jan-2021-A"  # Base filename of the TomoDose log file
data_path = r"\\vs20filesvr01\groups\CANCER\Physics - T\Temp\5498918\2021\01\{}.txt".format(filename)  # May need to change this. Where do we want to keep TomoDose logs?

fields = ["Detector #", "X Coordinate (mm)", "Z Coordinate (mm)", "Dose (cGy)", "Corrected Cnts (Counts)", "Raw (Counts)", "Offset (Counts)", "Background (count/tic)", "Detector (cf)"]  # Column names for array data
data = OrderedDict([(field, []) for field in fields])  # Initialize empty list for each column

axis = other_coord = hdr_row = None
with open(data_path, "r") as f:
    lines = f.readlines()[47:]  # Ignore first 47 lines (info)
    for i, line in enumerate(lines):
        line = [item.strip() for item in line.split("\t")[1:]]  # First column is unnecessary array info (e.g., "Arrays 1 and 2")
        if line == []:  # Ignore blank line
            continue
        if line[0] == "Pulse Count":  # Finished w/ array data
            break
        if line[0].startswith("Array"):  # Starting new array
            hdr_row = i
            m = match("Array \d+[a-z]?: Integrated dose on ([XY]) axis: [XY] = ([\+-]?\d+) mm", line[0])  # X- or Y-axis is specified in array "header"
            axis = m.group(1)
            other_coord = int(m.group(2).lstrip("+"))  # Remove sign from positive number 
        elif line[0] not in ["Detector #", "#"]:  # Array values, not field labels
            data["Detector #"].append(int(line[idx]))  # First value in this array row
            data["X Coordinate (mm)"].append(x[i])  # x-coordinate for this array
            data["Z Coordinate (mm)"].append(int(line[idx + 1]))
            data["Dose (cGy)"].append(float(line[idx + 2]))
            data["Corrected Cnts (Counts)"].append(float(line[idx + 3]))
            data["Raw (Counts)"].append(int(line[idx + 4]))
            data["Offset (Counts)"].append(int(line[idx + 5]))
            data["Background (count/tic)"].append(float(line[idx + 6]))
            data["Detector (cf)"].append(float(line[idx + 7]))

data = pd.DataFrame(data).set_index("Detector #").sort_index()  # Convert dict -> DataFrame for easier manipulation. Row labels are "Detector #". Sort rows by row #.
#data.to_csv(r"\\vs20filesvr01\groups\CANCER\Physics - T\Temp\test.csv")

# Plot data
fig = plt.figure(figsize=(5, 9))  # 5 x 9" image
ax = plt.subplot(111)

ax.set_axisbelow(True)  # Axes below markers
img = ax.scatter(data["X Coordinate (mm)"], data["Z Coordinate (mm)"], c=data["Dose (cGy)"], cmap="rainbow", marker=".", clip_on=False, zorder=10)  # Plot dose at coordinates (approximately same color scheme as Delta4)
ax.set_title("TomoDose Dose: {}".format(filename))
fig.colorbar(img, label="Dose (cGy)")  # Show color legend

# X-axis
# Tick marks at all x-values in data
ax.set_xlabel("X (mm)")
x = sorted(list(set(data["X Coordinate (mm)"])))
ax.set_xticks(x)
ax.set_xticklabels([str(val) for val in x], fontsize="x-small")  # Smaller font so tick mark labels don't overlap
ax.set_xlim(x[0], x[-1])  # Crop x-axis to x-values in data

# Z-axis (Tomo Z is everything else's - incl. matplotlib - Y)
# Tick marks are multiples of 10 (minor ticks at multiples of 5) spanning z-coordinates in data
ax.set_ylabel("Z (mm)")
z = list(set(data["Z Coordinate (mm)"]))
min_z = (min(z) // 10) * 10
max_z = (max(z) // 10) * 10 + 10
z_major = range(min_z, max_z + 1, 10)
z_minor = range(min_z, max_z + 1, 5)
ax.set_yticks(z_major)
ax.set_yticks(z_minor, minor=True)
ax.set_yticklabels([str(val) for val in z_major], fontsize="x-small")
ax.set_ylim(min_z, max_z)  # Crop y-axis to z-values in data

# Save image to file
plot_path = r"\\vs20filesvr01\groups\CANCER\Physics - T\Scripts\Output Files\TomoDoseLogFileAnalysis\{}.png".format(filename)
fig.savefig(plot_path)
