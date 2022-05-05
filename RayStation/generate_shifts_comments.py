import clr
from collections import OrderedDict
from datetime import datetime
import os
import re
import sys
from typing import Dict, Optional, Tuple

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox


# Directory to save the TXT file in
# This directory does not have to exist
OUTPUT_DIR = os.path.join('T:', os.sep, 'Physics', 'Scripts', 'Generate Shifts Comments')


def fmt_num(num: float, nearest_x: float) -> str:
    """Gets a nicely formatted string version of the number
    
    Rounds the number to the nearest `nearest_x`
    Removes any trailing zeroes. If nothing remians after the decimal point, removes it also
    
    Arguments
    ---------
    num: The number to format
    nearest_x: The interval to round to

    Returns
    -------
    The string that is the formatted number

    Example
    -------
    fmt_num(1.61, 0.5) -> '1.5'
    """

    num = round(num / nearest_x) * nearest_x
    # Don't count the "0." at the beginning
    num_places = len(str(nearest_x)) - 2
    num = format(num, '.{}f'.format(num_places))
    # Remove trailing zeroes. If there was nothing but zeroes after the decinal point, remove it, too
    num = num.rstrip('0').rstrip('.')
    return num


def setup_beam_ssd(setup_beam: PyScriptObject, regex: str) -> Optional[float]:
    """Gets the SSD of the setup beam if its name or description contains the given regular expression

    The matched substring must not be surrounded by alphabetric characters

    Arguments
    ---------
    setup_beam: The setup beam whose SSD to return
    regex: The regular expression to match a substring in the setup beam's name or description

    Returns
    -------
    The setup beam's SSD, if the regular expression is found, or None otherwise or if the SSD is infinite
    """
    regex = r'(^|[^a-zA-Z])' + regex + r'($|[^a-zA-Z])'  # Matching substring must not be immediately preceded or followed by an alphabetic character
    if re.search(regex, setup_beam.Name, re.IGNORECASE) is None and re.search(regex, setup_beam.Description, re.IGNORECASE) is None:  # No substring match in either the Name or the Description
        return
    setup_beam_ssd = setup_beam.GetSSD()
    if abs(setup_beam_ssd) == float('inf'):  # Infinite SSD
        return
    return setup_beam_ssd


def cb_setup_beams(beam_set: PyScriptObject, pos: str) -> Tuple[Dict[PyScriptObject, float], Dict[PyScriptObject, float]]:
    """Gets the cone-beam (CB) setup beams for a beam set

    A setup beam is considered CB if its name or description contains "CB" or "CBCT", or "AP" (for supine patients) or "PA" (for prone patients)
    Only CB setup beams with SSD are returned

    Arguments
    ---------
    beam_set: The beam set whose CB setup beams to return

    Returns
    -------
    A 2-tuple of dictionaries with setup beams as the keys and SSDs as the values.
    The first element is "CB" setup beams, the second "AP"/"PA" setup beams.
    """
    cbs, aps = OrderedDict(), OrderedDict()
    
    # Check each setup beam to see if it is CB
    for setup_beam in beam_set.PatientSetup.SetupBeams:
        # Is it "CB"?
        ssd = setup_beam_ssd(setup_beam, r'CB(CT)?')
        if ssd is None:
            # Not "CB", so is it "AP"/"PA"?
            ssd = setup_beam_ssd(setup_beam, pos)
            if ssd is not None:
                aps[setup_beam] = ssd
        else:
            cbs[setup_beam] = ssd
    return cbs, aps


def output_filepath(beam_set: PyScriptObject) -> str:
    """Creates an output filename for the shifts TXT file for the beam set

    TXT file is in `OUTPUT_DIR`
    Filename is in the format "<beam set ID> Shifts Comments YYYY-MM-DD HH:MM:SS.txt"
    The colon in the beam set ID is replaced with a hyphen for Windows filename compatibility

    Arguments
    ---------
    beam_set: The beam set to return an output filepath for

    Returns
    -------
    The absolute filepath of the TXT file to create

    Example
    -------
    output_filepath(some_beam_set) -> 'My\Output\Dir\plan name-beam set name 2022-05-04 16_15_30.txt'
    """
    dt = datetime.now().strftime('%Y-%m-%d %H_%M_%S')
    bs_id = '-'.join(beam_set.BeamSetIdentifier().split(':'))
    filename = bs_id + ' Shifts Comments ' + dt + '.txt'
    filename = re.sub(r'[<>:"/\\\|\?\*]', '_', filename)
    return os.path.join(OUTPUT_DIR, filename)


def generate_shifts_comments(include_msq_shifts: Optional[bool] = False) -> None:
    """Creates and opens a TXT file listing patient shifts (couch shifts, MOSAIQ Site Setup shifts) and CB/AP/PA setup SSD if applicable.
    The MOSAIQ Setup Shifts are the same as the couch setup shifts but using the opposite directional term (e.g., "Posterior" if the couch shift is "Anterior").

    If there is no shift, writes "No setup shifts" instead of the couch shift.
    If there is a single CB setup beam, uses it. Otherwise, if there is a single AP (for supine) / PA (for prone) setup beam, uses it. Otherwise, alerts the user and exits the script.

    Assumes that CB setup beam name or description contains "CB" or "CBCT" (case insensitive) and that other AP/PA setup beam names or descriptions contain "AP" or "PA" (case insensitive).

    Arguments
    include_msq_shifts: True to include the MOSAIQ Setup Shifts, False otherwise.
                        Defaults to False.
    """
    # Get current beam set
    try:
        beam_set = get_current('BeamSet')
    except:
        MessageBox.Show('There is no beam set loaded. Click OK to abort the script.', 'No Beam Set Loaded')
        sys.exit()

    # Ensure beam set has beam(s)
    if beam_set.Beams.Count == 0:
        MessageBox('There are no beams in the current beam set.\nClick OK to abort the script.', 'No Beams')
        sys.exit()

    # Localization geometry
    loc_geom = beam_set.GetStructureSet().LocalizationPoiGeometry
    if loc_geom.Point is None or loc_geom.Point.x == float('inf'):
        MessageBox.Show('There is no localization geometry on the planning exam.\nClick OK to abort the script.', 'No Localization Geometry')
        sys.exit()
    loc_geom = loc_geom.Point

    # Beam set isocenter
    isos = [beam.Isocenter.Position for beam in beam_set.Beams]
    # Ensure all beam isos have same coordinates
    if len(set((iso.x, iso.y, iso.z) for iso in isos)) > 1:
        MessageBox.Show('Beams in beam set have different isocenter coordinates.\nClick OK to abort the script.', 'Beam Isos Differ')
        sys.exit()
    iso = isos[0]  # All isos have same position, so doesn't matter which one we use

    # Calculate shifts = localization geometry minus isocenter
    x, y, z = (loc_geom[dim] - iso[dim]
               for dim in 'xyz')

    # ----------------------------- Shift directions ----------------------------- #

    # "Left" or "Right"
    if x < 0:
        couch_r_l = 'Right'
        msq_r_l = 'Left'
    else:
        couch_r_l = 'Left'
        msq_r_l = 'Right'

    # "Superior" or "Inferior"
    if z < 0:
        couch_i_s = 'Inferior'
        msq_i_s = 'Superior'
    else:
        couch_i_s = 'Superior'
        msq_i_s = 'Inferior'

    # "Posterior" or "Anterior"
    if y < 0:
        couch_p_a = 'Anterior'
        msq_p_a = 'Posterior'
    else:
        couch_p_a = 'Posterior'
        msq_p_a = 'Anterior'

    # Round and format shifts
    x, y, z = (fmt_num(abs(shift), 0.1) for shift in (x, y, z))

    # Add couch shifts to comments
    comments = '1. Align to initial CT marks\n'
    if x == y == z == '0':
        comments += 'No setup shifts'
    else:
        comments += f'2. Shift couch so PATIENT is moved: {couch_r_l} {x} cm (patient\'s right/left), {couch_i_s} {z} cm, {couch_p_a} {y} cm'

    # CB setup SSD
    pos = 'AP' if 'Supine' in beam_set.PatientPosition else 'PA'
    cbs, aps = cb_setup_beams(beam_set, pos)  # All CB or AP/PA (depending on pt position) setup beams
    ssd = None
    if len(cbs) > 1:  # Multiple CB setup beams
        if len(aps) > 1:  # Multiple AP/PA setup beams
            MessageBox.Show('There are multiple CB and multiple ' + pos + ' setup beams, and the script doesn\'t know which one to use. Click OK to abort the script.', 'Multiple CB/' + pos + ' Setup Beams')
            sys.exit()
        elif not aps:  # No AP/PA setup beams
            MessageBox.Show('There are multiple CB setup beams and no other ' + pos + ' setup beams, and the script doesn\'t know which one to use. Click OK to abort the script.', 'Multiple CB Setup Beams')
            sys.exit()
        else:  # Exactly 1 AP/PA setup beam
            MessageBox.Show('There are multiple CB and one other ' + pos + ' setup beam, and the script doesn\'t know which one to use. Click OK to abort the script.', 'Multiple CB/' + pos + ' Setup Beams')
            sys.exit()
    elif cbs:  # Exactly 1 CB setup beam, so use it
        ssd = list(cbs.values())[0]
    elif aps:  # At least one AP/PA setup beam
        if len(aps) > 1:  # Multiple AP/PA setup beams
            MessageBox.Show('There are multiple ' + pos + ' setup beams and no other CB setup beams, and the script doesn\'t know which one to use. Click OK to abort the script.', 'Multiple ' + pos + ' Setup Beams')
            sys.exit()
        else:  # Exactly 1 AP/PA setup beam, so use it
            ssd = list(aps.values())[0]
    else:  # No CB or other AP/PA setup beams
        MessageBox.Show('There are no CB or other ' + pos + ' setup beams. Click OK to abort the script.', 'No CB/' + pos + ' Setup Beams')
        sys.exit()
    comments += '\n' + pos + ' setup SSD: ' + fmt_num(ssd, 0.25)

    # Add MOSAIQ site setup shifts to comments
    if include_msq_shifts:
        comments += '\n\nMOSAIQ Site Setup shifts:\n' \
                 + f'{msq_r_l} {x} cm, {msq_i_s} {z} cm, {msq_p_a} {y} cm'

    # Write comments to file
    if not os.path.isdir(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    filepath = output_filepath(beam_set)
    with open(filepath, 'w') as f:
        f.write(comments)

    # Open comments file
    os.system('START /B notepad.exe "' + filepath + '"')
