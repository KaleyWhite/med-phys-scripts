"""Parse measured dose profile spreadsheets exported from RayPhysics and analyze which data points and dose values are shared among machines

See Appendix B.3 of "RayPhysics 8B User Manual" for CSV file format

Write four new spreadsheets:
1. "PDDS.csv": PDD ("Depth" in RayPhysics) profile data
   Example row:
   Machine,Field Size [cm],Depth [cm],Dose
   ELEKTA,1,-5,34.35473
2. "X, Y Profiles.csv": Crossline and inline ("X" and "Y" in RayPhysics, resp.) profile data
   Example row:
   Machine,Curve Type,Field Size [cm],Depth [cm],Position [cm],Dose
   ELEKTA,Crossline,1,15,-103,0.185646
3. "Shared PDD Values.csv": For each combination of machines that share each PDD field size, "Y" or "N" to indicate whether each machine's depths in the PDD are the same, and, if so, if the profiles are identical
   Example row:
   Machines,Field Size [cm],Share all depths,Share all doses
   ELEKTA; SBRT 6MV,1,N,N/A
4. "Shared X, Y Profile Values.csv": For each combination of machines that share each combination of field size and depth, "Y" or "N" to indicate whether each machine's coordinates/positions in the profile are the same, and, if so, if the profiles are identical
   Example row:
   Machines,Field Size [cm],Depth [cm],Share all positions,Share all doses
   ELEKTA; SBRT 6MV,1,50,N,N/A

Assumptions
-----------
All profiles have same SSD. SSD info is in input CSV files, but the script has not been modified to account for this because we don't need to use the script again for the time being.
"""

from itertools import combinations
from os import listdir

import pandas as pd


# Filepaths
t_path = r"\\vs20filesvr01\groups\CANCER\Physics - T"  # Physics folder on the T drive
input_filepath = r"{}\Scripts\Data\RayPhysics Measured Dose".format(t_path)  # Folder w/ CSV files exported from RayPhysics
output_filepath = r"{}\Scripts\Output Files\RayPhysicsMeasuredDoseParser".format(t_path)  # Folder to which to export the 4 generated CSV files
pdd_output_filepath = r"{}\PDDs.csv".format(output_filepath)  # Complete filename for "PDDs.csv"
x_y_output_filepath = r"{}\X, Y Profiles.csv".format(output_filepath)  # Complete filename for "X, Y Profiles.csv"
shared_pdd_output_filepath = r"{}\Shared PDD Values.csv".format(output_filepath)  # Complete filename for "Shared PDD Values.csv"
shared_x_y_output_filepath = r"{}\Shared X, Y Profile Values.csv".format(output_filepath)  # Complete filename for "Shared X, Y Profile Values.csv"

# Dict to hold info on each profile parsed from the input file
# Structure:
# {
#    Machine name #1: {
#                           Field size #1: {
#                                               "PDD": {
#                                                           "Depth [cm]": [],
#                                                           "Dose": []    
#                                                      },
#                                               "X": {
#                                                            Depth #1: {
#                                                                           "Position [cm]": [],
#                                                                           "Dose": []
#                                                            }
#                                                    },
#                                               "Y": {
#                                                            Depth #1: {
#                                                                           "Position [cm]": [],
#                                                                           "Dose": []
#                                                            }
#                                                    }
#                                           }
#                      }
# }
profiles_dict = {}

# Read in data from each file in the input directory
for f in listdir(input_filepath):
    f = r"{}\{}".format(input_filepath, f)  #   Absolute path to file
    with open(f) as f:
        profiles = f.read().split("End\n")[:-1]  # "End" indicates end of a profile
        meta = profiles[0].split("\n")[:8]  # First 8 lines of the file start with "#" and apply to all profiles
        profiles[0] = profiles[0][sum(len(elem) + 1 for elem in meta):]  # First profile starts after all meta info

        # Add machine to profiles dictionary, if necessary
        machine = meta[2].split(":")[-1].strip()  # E.g., "#Machine Name: ELEKTA" -> "ELEKTA"
        if machine not in profiles_dict:
            profiles_dict[machine] = {}

        # Parse each profile in the list
        for profile in profiles:
            profile = profile.split("\n")[:-1]  # Last item in list is "", so ignore it
            fs = int(profile[2].split(";")[-1].strip()) // 5  # Convert field size from mm halfway across, to cm all the way across
            # Add field size to this machine in the profiles dictionary, if necessary
            if fs not in profiles_dict[machine]:
                profiles_dict[machine][fs] = {}

            measurements = profile[7:]  # First 7 lines of the profile data are info (e.g., machine), not data points
            pos = [float(measurement.split(";")[0].strip()) for measurement in measurements]  # First value in each data point
            dose = [float(measurement.split(";")[1].strip()) for measurement in measurements]  # Second value in each data point

            curve_type = profile[3].split(";")[-1].strip()  # "Depth", "Crossline", or "Inline"
            if curve_type == "Depth":  # PDD
                # Add "PDD" curve type to machine field size in profiles dictionary, if necessary
                if curve_type not in profiles_dict[machine][fs]:
                    profiles_dict[machine][fs]["PDD"] = {"Depth [cm]": pos, "Dose": dose}
                else:  # Add data points to the PDD data for this machine and field size
                    profiles_dict[machine][fs]["PDD"]["Depth [cm]"].extend(pos)
                    profiles_dict[machine][fs]["PDD"]["Dose"].extend(dose)
            else:  # Crossline (X) or inline (Y)
                depth = float(profile[6].split(";")[-1].strip())  # Z-coordinate of the profile
                # Add curve type to machine field size in profiles dictionary, if necessary 
                if curve_type not in profiles_dict[machine][fs]:
                    profiles_dict[machine][fs][curve_type] = {}
                # Add depth to machine field size curve type in profiles dictionary, if necessary
                if depth not in profiles_dict[machine][fs][curve_type]:
                    profiles_dict[machine][fs][curve_type][depth] = {"Position [cm]": pos, "Dose": dose}
                else:  # Add data points to the data for this machine, field size, depth, and curve type
                    profiles_dict[machine][fs][curve_type][depth]["Position [cm]"].extend(pos)
                    profiles_dict[machine][fs][curve_type][depth]["Dose"].extend(dose)

# Convert profiles dictionary to DataFrames
pdd_df = x_y_df = None
for machine, machine_profiles in profiles_dict.items():
    for fs, fs_profiles in machine_profiles.items():
        for curve_type, curve_type_profiles in fs_profiles.items():
            if curve_type == "PDD":
                df = pd.DataFrame(curve_type_profiles)  # Start w/ "Depth [cm]" and "Dose" columns
                # Add "Machine" and "Field Size [cm]" columns
                df["Machine"] = machine 
                df["Field Size [cm]"] = fs 
                
                if pdd_df is None:
                    pdd_df = df.loc[:]  # Copy the new DataFrame into the PDD DataFrame variable
                else:
                    pdd_df = pd.concat([pdd_df, df])  # Combine the new DataFrame with the existing PDD DataFrame
            else:
                for depth, depth_profiles in curve_type_profiles.items():
                    df = pd.DataFrame(depth_profiles)   # Start w/ "Position [cm]" and "Dose" columns

                    # Add "Machine", "Field Size [cm]", "Curve Type", and "Depth [cm]" columns
                    df["Machine"] = machine
                    df["Field Size [cm]"] = fs
                    df["Curve Type"] = curve_type
                    df["Depth [cm]"] = depth
                    
                    if x_y_df is None:
                        x_y_df = df.loc[:]  # Copy the new DataFrame into the x_y DataFrame variable
                    else:
                        x_y_df = pd.concat([x_y_df, df])  # Combine the new DataFrame with the existing x_y DataFrame

# Reorder columns
pdd_df = pdd_df[["Machine", "Field Size [cm]", "Depth [cm]", "Dose"]]
x_y_df = x_y_df[["Machine", "Curve Type", "Field Size [cm]", "Depth [cm]", "Position [cm]", "Dose"]]

# Sort rows by all column values
pdd_df.sort_values(pdd_df.columns.tolist(), inplace=True)
x_y_df.sort_values(x_y_df.columns.tolist(), inplace=True)

# Write DataFrames to output files (no row numbers)
pdd_df.to_csv(pdd_output_filepath, index=False)
x_y_df.to_csv(x_y_output_filepath, index=False)

# Determine shared PDD data
shared_pdd_data = {colname: [] for colname in ["Machines", "Field Size [cm]", "Share all depths", "Share all doses"]}
machines = pdd_df["Machine"].unique()  # All "Machine" values in the PDD DataFrame
for fs in pdd_df["Field Size [cm]"].unique():
    pdd_df_fs = pdd_df.loc[pdd_df["Field Size [cm]"] == fs]  # Rows with the given field size
    for i in range(2, len(machines) + 1):  # All possible lengths of machine combinations
        machine_combos = combinations(machines, i)  # All machine combinations of the given length
        for combo in machine_combos:
            dfs = [pdd_df_fs.loc[pdd_df_fs["Machine"] == machine] for machine in combo if len(pdd_df_fs.loc[pdd_df_fs["Machine"] == machine]) > 0]  # A DataFrame for each machine that has PDD data for the given field size
            if len(dfs) != i:  # Not all the machines in this combination have data for the given field size
                continue
            shared_pdd_data["Machines"].append("; ".join(list(combo)))  # Machine names are a semicolon-separated list
            shared_pdd_data["Field Size [cm]"].append(fs)
            if all(df_1["Depth [cm]"].equals(df_2["Depth [cm]"]) for df_1, df_2 in combinations(dfs, 2)):  # All machines in this combo share all depths
                shared_pdd_data["Share all depths"].append("Y")
                if len(set(df["Dose"] for df in dfs)) == 1:  # Dose at each coordinate is the same for all machines
                    shared_pdd_data["Share all doses"].append("Y")
                    shared_pdd_data["Share all doses"].append("N")
            else:
                shared_pdd_data["Share all depths"].append("N")
                shared_pdd_data["Share all doses"].append("N/A")
shared_pdd_data = pd.DataFrame(shared_pdd_data)  # Convert dictionary to DataFrame
shared_pdd_data.to_csv(shared_pdd_output_filepath, index=False)  # Write DataFrame to output file (no row numbers)

# Determine shared X, Y profile data
# Same process as for PDD data, except this time we must analyze each profile (defined by the depth)
shared_x_y_data = {colname: [] for colname in ["Machines", "Field Size [cm]", "Depth [cm]", "Share all positions", "Share all doses"]}
machines = x_y_df["Machine"].unique()
for fs in x_y_df["Field Size [cm]"].unique():
    x_y_df_fs = x_y_df.loc[x_y_df["Field Size [cm]"] == fs]
    for depth in x_y_df_fs["Depth [cm]"].unique():
        x_y_df_depth = x_y_df_fs.loc[x_y_df_fs["Depth [cm]"] == depth]
    for i in range(2, len(machines) + 1):
        machine_combos = combinations(machines, i)
        for combo in machine_combos:
            dfs = [x_y_df_depth.loc[x_y_df_depth["Machine"] == machine] for machine in combo if len(x_y_df_depth.loc[x_y_df_depth["Machine"] == machine]) > 0]
            if len(dfs) != i:
                continue
            shared_x_y_data["Machines"].append("; ".join(list(combo)))
            shared_x_y_data["Field Size [cm]"].append(fs)
            shared_x_y_data["Depth [cm]"].append(depth)
            if all(df_1["Position [cm]"].equals(df_2["Position [cm]"]) for df_1, df_2 in combinations(dfs, 2)):
                shared_x_y_data["Share all positions"].append("Y")
                if len(set(df["Dose"] for df in dfs)) == 1:
                    shared_x_y_data["Share all doses"].append("Y")
                    shared_x_y_data["Share all doses"].append("N")
            else:
                shared_x_y_data["Share all positions"].append("N")
                shared_x_y_data["Share all doses"].append("N/A")
shared_x_y_data = pd.DataFrame(shared_x_y_data)
shared_x_y_data.to_csv(shared_x_y_output_filepath, index=False)
