import clr
import random
import re
import sys
from typing import Tuple

from connect import *

import pandas as pd

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import Color
from System.Windows.Forms import MessageBox


# Absolute path to the spreadsheet that CRMC heavily modified from the one from the TG-263 website
TG263_FILEPATH = os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-spreadsheets', 'Structure Names & Colors', 'TG-263 Nomenclature with CRMC Colors.xlsm')


def extract_argb(color: str) -> Tuple[int, int, int, int]:
    """Converts a string color '(A, R, G, B)' into a 4-tuple (A, R, G, B)

    Argument
    --------
    color: The color whose components to extract
           Must be in the format '(A, R, G, B)', with all A, R, G, B integers between 0 and 255, inclusive

    Returns
    -------
    The 4-tuple (A, R, G, B)
    """
    m = re.match(r'\((\d+), (\d+), (\d+), (\d+)\)', color)
    a, r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
    return a, r, g, b


def struct_names_colors() -> None:
    """Renames and recolors ROIs in the current case, according to "TG-263 Nomenclature with CRMC Colors" spreadsheet
    
    Recolor target ROIs according to target type. R, G, and B components come from the structure in the spreadsheet whose name is the target type. A is a random number between 128 and 255, inclusive.
    First, try to match it to an incorrect name in the "Possible Incorrect Names" table. Otherwise, search the main table. 
    When comparing names, copy numbers (which prevent duplicates) and suffixes after a carat are ignored. Copy numbers are removed in the new structure names. Suffixes are added back onto the new ROI name.
    """
    # Get current objects
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort script.', 'No Open Case')
        sys.exit()

    # Read in names and colors data
    tg263 = pd.read_excel(TG263_FILEPATH, sheet_name=['Names & Colors', 'Possible Incorrect Names'])
    names_colors = tg263['Names & Colors'][['Target Type', 'Major Category', 'TG-263 Primary Name', 'Color']]
    incorrect = tg263['Possible Incorrect Names']
    
    # Approved ROI names/colors cannot be changed
    approved_struct_names = []
    for ss in case.PatientModel.StructureSets:
        for approved_ss in ss.ApprovedStructureSets:
            approved_struct_names.extend(geom.OfRoi.Name for geom in approved_ss.ApprovedRoiStructures)
            approved_struct_names.extend(geom.OfPoi.Name for geom in approved_ss.ApprovedPoiStructures)
    approved_struct_names = list(set(approved_struct_names))
    
    # Lists of structures whose names/colors can't be changed because they are approved
    approved_name_incorrect = []  # Non-target incorrect name
    approved_color_incorrect = []  # Incorrect color
    approved_name_color_incorrect = []  # Non-target incorrect name AND color

    no_match = []  # Names in neither Names & Colors nor Possible Incorrect Names

    # Get target types
    target_types = names_colors.loc[names_colors['Target Type'] == 'Target', 'Major Category'].values  # Should be "CTV", "GTV", "ITV", "PTV"
    
    structs = list(case.PatientModel.RegionsOfInterest) + list(case.PatientModel.PointsOfInterest)
    for struct in structs:
        # Target
        if struct.Type.upper() in target_types:
            correct_color = names_colors.loc[names_colors['TG-263 Primary Name'] == struct.Type.upper(), 'Color']  # Color for that target type
            a, r, g, b = extract_argb(correct_color.iloc[0])
            if (struct.Color.R, struct.Color.G, struct.Color.B) != (r, g, b) or a < 128:  # Ignore A component unless it is too low for the color to be recognizable
                if struct.Name in approved_struct_names:
                    approved_color_incorrect.append(struct.Name)
                else:
                    struct.Color = Color.FromArgb(random.randint(128, 255), r, g, b)  # Keep high A so color is recognizable
        # Non-target
        else:
            base_name = struct.Name  # Name without suffix and/or copy number

            # Remove copy number if it exists
            m = re.search(r' \(\d+\)$', struct.Name)
            if m is not None:
                base_name = struct.Name[:m.start()]

            # Remove suffix if it exists
            suffix = ''
            if '^' in base_name and base_name != 'External^NoBox':
                idx = base_name.index('^')
                base_name = base_name[:idx]
                suffix = base_name[(idx + 1):]
         
            correct_name = incorrect.loc[incorrect['Incorrect'] == base_name.lower(), 'Correct']  # Check for name in Possible Incorrect Names
            new_name = correct_name.iloc[0] if len(correct_name) > 0 else base_name  # If not in Possible Incorrect Names, assume correct
            correct_name_color = names_colors.loc[names_colors['TG-263 Primary Name'].str.lower() == new_name.lower(), ['TG-263 Primary Name', 'Color']]  # This will be an empty Series if the name is not correct
            if len(correct_name_color) > 0:
                new_name = correct_name_color['TG-263 Primary Name'].iloc[0]
                a, r, g, b = extract_argb(correct_name_color['Color'].iloc[0])
                if (struct.Color.A, struct.Color.R, struct.Color.G, struct.Color.B) != (a, r, g, b):
                    if struct.Name in approved_struct_names:
                        if base_name != new_name:
                            approved_name_color_incorrect.append(struct.Name)
                        else:
                            approved_color_incorrect.append(struct.Name)
                    else:
                        struct.Color = Color.FromArgb(a, r, g, b)
                        struct.Name = new_name
                elif base_name != new_name:
                    if struct.Name in approved_struct_names:
                        approved_name_incorrect.append(struct.Name)
                    else:
                        if suffix:
                            new_name += '^' + suffix
                        struct.Name = new_name
            elif base_name not in no_match:
                no_match.append(base_name)

    # Display any warnings
    warnings = ''
    if approved_name_incorrect:
        warnings += 'The following structures are incorrectly named, but they can\'t be changed because they are approved:\n  -  ' + '\n  -  '.join(approved_name_incorrect) + '\n\n'
    if approved_color_incorrect:
        warnings += 'The following structures are incorrectly colored, but they can\'t be changed because they are approved:\n  -  ' + '\n  -  '.join(approved_color_incorrect) + '\n\n'
    if no_match:
        warnings += 'The following structure names are in neither the "Names & Colors" nor the "Possible Incorrect Names" table. (Names are shown without copy numbers or suffixes after a carat.)\nIs there information in the name that should be a suffix after a carat? If so, fix this and rerun the script.\nOtherwise, if they follow TG-263 and/or CRMC nomenclature conventions, please add them to the "TG-263 Nomenclature with CRMC Colors" spreadsheet.\n  -  ' + '\n  -  '.join(no_match) + '\n\n'
    if warnings:
        MessageBox.Show(warnings, 'Warnings')
