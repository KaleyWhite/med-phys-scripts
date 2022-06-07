import os
import re
import shutil
from tkinter import Tk

from tkfilebrowser import askopendirnames


def ascii_dose_profiles_cm_to_mm() -> None:
    """Corrects coordinates in IBA RFA300 ASCII Measurement Dump electron dose profiles (BDS format) in the user-selected folder(s).

    Dose profiles have the following format:

    :MSR 	10	 # No. of measurement in file
    :SYS BDS 0 # Beam Data Scanner System
    #
    # RFA300 ASCII Measurement Dump ( BDS format )
    #
    # Measurement number 	1
    #
    %VNR 1.0
    %MOD 	RAT
    .
    .
    .
    %STS 	   -9.0	    0.0	    5.0 # Start Scan values in mm ( X , Y , Z )
    %EDS 	    9.0	    0.0	    5.0 # End Scan values in mm ( X , Y , Z )
    #
    #	  X      Y      Z     Dose
    #
    = 	   -9.0	    0.0	    5.0	   16.6
    = 	   -8.9	    0.0	    5.0	   18.6
    .
    .
    .
    :EOM  # End of Measurement
    #
    # RFA300 ASCII Measurement Dump ( BDS format )
    #
    # Measurement number 	2
    .
    .
    .
    :EOF # End of File

    The coordinate units are listed as mm, but the coordinates are clearly in cm. This function converts the coordinates from cm to mm.
    Does not change original files, but writes a new folder called "Corrected Units" in each old folder. If a "Corrected Units" subfolder already exists, it is replaced.
    Assumes all .asc files in all provided subfolders need correcting.
    """
    Tk().withdraw()
    folders = askopendirnames(title='Choose folder(s) containing profiles')

    for folder in folders:
        # Create "Corrected Units" folder
        new_subfolder = os.path.join(folder, "Corrected Units")
        if os.path.isdir(new_subfolder):
            shutil.rmtree(new_subfolder)
        os.mkdir(new_subfolder)
        
        # Correct each ASC file in the folder
        for file in os.listdir(folder):
            if os.path.splitext(file)[-1] == '.asc':
                abs_file = os.path.join(folder, file)
                with open(abs_file) as data:
                    # Multiply coordinates by 10
                    # Leave non-coordinate lines as-are
                    new_data = ""
                    lines = data.readlines()
                    for line in lines:
                        if re.match(r"(=|%STS|%EDS).*", line):  # Line contains coordinates
                            cm = list(re.finditer(r"-?\d+\.\d+", line))[:3]  # First 3 numbers in the line (the X, Y, and Z coordinates)
                            for c in cm[::-1]:  # Work backward so that indices in the line do not change
                                line = line.replace(line[c.start():c.end()], str(float(c.group()) * 10), 1)  # Replace coordinate in cm with that coordinate in mm (multiplied by 10)
                        new_data += line
                        
                    # Write corrected data to new file
                    new_f = os.path.join(new_subfolder, file)  # New file has same name as old (but is, of course, in the new subfolder)
                    with open(new_f, "w") as new_f:
                        new_f.write(new_data)


if __name__ == '__main__':
    ascii_dose_profiles_cm_to_mm()
