# This is a hacky way to take a screenshot of the DVH and save it to a file.
# Double-click inside the DVH.
# Push the Scripting window on the left to just to the right of the script name "DVHScreenshot". 
# This ensures that the correct parts of the screen are captured.
# For predictable results, close all other windows except for RayStation.

import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
import os
import sys
sys.path.append(os.path.join("T:", "Physics - T", "Scripts", "Library Files"))

from connect import *
from System.Drawing import *
from System.Windows.Forms import *

from LibraryFunctions import *


def dvh_screenshot():
    start_x = 0.107 * Screen.PrimaryScreen.Bounds.Width  # 205
    bottom_cutoff = 0.07 * Screen.PrimaryScreen.Bounds.Height
    start_y = 0.19 * Screen.PrimaryScreen.Bounds.Height  # 213
    size_x = Screen.PrimaryScreen.Bounds.Width - start_x
    size_y = Screen.PrimaryScreen.Bounds.Height - start_y - bottom_cutoff

    bmp = Bitmap(int(round(size_x)), int(round(size_y)))
    g = Graphics.FromImage(bmp)
    g.CopyFromScreen(int(round(start_x)), int(round(start_y)), 0, 0, bmp.Size)
    g.Dispose()

    patient, plan = get_current_variables("Patient", "Plan")
    patient_name = format_patient_name_display(patient.Name)
    filename = os.path.join(get_output_file_path(__name__), "{} {} DVH {}.png".format(patient_name, plan.Name, get_timestamp()))
    bmp.Save(filename)

    sys.exit()
    