"""In the user-selected directory, creates subdirectory structure based on file creation dates.

Creates a subdirectory called "Organized" in thye user-selected directory. If this subdirectiry already exists, it is replaced.
Creates sub-subdirectories for years and months (numbers).
Copies each file in the directory into the appropriate new sub-subdirectory.
Is not recursive.

Example new directory structure:
> <Original directory name>
    > Organized
        > 2015
            > 01
            > 02
        > 2017
            > 05
            > 07
            > 12
        > 2018
            > 01
"""

import os
import shutil
import time
from datetime import datetime
from stat import *

from tkinter import Tk

from tkfilebrowser import askopendirname


def get_ctime(filename: str) -> datetime:
    """Return the file's creation time
    
    Args:
    filename: Absolute filepath.
        
    Returns:
    File creation date, or modification date if on Linux.
    """

    stats = os.stat(filename)
    try:
        ctime = stats.st_ctime
    except:
        ctime = stats.st_mtime
    return datetime.strptime(time.ctime(ctime), '%a %b %d %H:%M:%S %Y')


def organize_files_by_date() -> None:
    Tk().withdraw()
    directory = askopendirname(title='Choose directory with dated files')

    # Create "Organized" directory to house new directory structure
    new_directory = os.path.join(directory, 'Organized')
    if os.path.isdir(new_directory):
        shutil.rmtree(new_directory)
    os.mkdir(new_directory)

    # Copy each file into the correct subdirectory
    # If the appropriate subdirectory doesn't exist, create it
    for f in os.listdir(directory):
        abs_f = os.path.join(directory, f)
        ct = get_ctime(abs_f)
        new_subdir = os.path.join(new_directory, str(ct.year), str('{:02d}'.format(ct.month)))
        if not os.path.isdir(new_subdir):
            os.makedirs(new_subdir)
        shutil.copyfile(abs_f, os.path.join(new_subdir, f))
