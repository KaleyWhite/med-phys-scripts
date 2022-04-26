import clr
import os
import sys

from connect import *
import pandas as pd

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # For displaying errors


# Absoluet path to the "TG-263 Nomenclature with CRMC Colors" spreadsheet
TG263_PATH = os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-spreadsheets', 'TG-263 Nomenclature with CRMC Colors.xlsm')


def crmc_color(roi_name: str) -> str:
    """Gets the CRMC conventional color for the provided ROI, according to the "TG-263 Nomenclature with CRMC Colors" spreadsheet

    If the ROI name is not present in the spreadsheet, returns purple

    Arguments
    ---------
    roi_name: The name of the ROI whose conventional color to return

    Returns
    -------
    The ROI color, in the format "A, R, G, B"
    """
    tg263 = pd.read_excel(TG263_PATH, sheet_name='Names & Colors', usecols=['TG-263 Primary Name', 'Color'])
    tg263.set_index('TG-263 Primary Name', drop=True, inplace=True)
    try:
        return tg263.loc[roi_name, 'Color'][1:-1]
    except KeyError:
        return Color.Purple


def dose_grid_box() -> None:
    """Adds a box ROI with geometry that outlines the dose grid of the current beam set"""

    # Get current case
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('No case is open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    # Get current beam set
    try:
        beam_set = get_current('BeamSet')
    except:
        MessageBox.Show('No beam set is loaded. Click OK to abort the script.', 'No Beam Set Loaded')
        sys.exit()

    dose = beam_set.FractionDose
    struct_set = beam_set.GetStructureSet()
    exam = struct_set.OnExamination
    dg = dose.InDoseGrid
    
    # Get DoseGrid ROI name and color
    roi_name = case.PatientModel.GetUniqueRoiName(Name='zDoseGrid')
    color = crmc_color('zDoseGrid')

    # Create dose grid ROI
    box = case.PatientModel.CreateRoi(Name=roi_name, Color=color, Type='Control')
    
    # Box size and center coordinates
    box_sz, box_ctr = {}, {}
    for dim, coord in dg.Corner.items():
        box_sz[dim] = dg.NrVoxels[dim] * dg.VoxelSize[dim]
        box_ctr[dim] = coord + (box_sz[dim] / 2)

    # Create box and update its geometry (must update all geometries at once)
    box.CreateBoxGeometry(Size=box_sz, Examination=exam, Center=box_ctr, VoxelSize=dg.VoxelSize.x)
    if dose.DoseValues is not None:
        dose.UpdateDoseGridStructures()
