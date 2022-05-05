import clr
import os
import sys
from typing import Optional

import pandas as pd

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # For displaying errors


# Absolute path to the "TG-263 Nomenclature with CRMC Colors" spreadsheet
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
        return tg263.loc[roi_name, 'Color'][1:-1]  # Remove parens
    except KeyError:
        return Color.Purple


def beam_set_from_beam(case: PyScriptObject, beam: PyScriptObject) -> PyScriptObject:
    """Gets the beam set that the beam belongs to.

    Args:
        case: The case that the beam belongs to
        beam: The beam

    Returns
    -------
    The beam set containing the beam, or None if the beam is not found in the case
    """
    for plan in case.TreatmentPlans:
        for beam_set in plan.BeamSets:
            for b in beam_set.Beams:
                if b.Equals(beam):
                    return beam_set


def exam_from_dose(case: PyScriptObject, dose: PyScriptObject) -> PyScriptObject:
    """Gets the examination associated with the dose distribution

    Arguments
    ---------
    case: The case that the dose distribution belongs to
    dose: The dose distribution

    Returns
    -------
    The examination associated with the dose distribution
    """
    if hasattr(dose, 'WeightedDoseReferences'):  # Plan dose, dose sum
        return dose.WeightedDoseReferences[0].DoseDistribution.ForBeamSet.GetStructureSet().OnExamination
    if hasattr(dose, 'ForBeamSet'):  # Beam set dose, perturbed dose, dose on additional set
        return dose.ForBeamSet.GetPlanningExamination()
    if hasattr(dose, 'ForBeam'):  # Beam dose
        return beam_set_from_beam(case, dose.ForBeam).GetPlanningExamination()
    if hasattr(dose, 'OfDoseDistribution'):  # Deformed dose
        return dose.OfDoseDistribution.ForBeamSet.GetPlanningExamination()


def dose_grid_box(case: PyScriptObject, dose_dist: Optional[PyScriptObject]) -> PyScriptObject:
    """Adds a box ROI with geometry that outlines the dose grid

    Arguments
    ---------
    case: The case that the dose distribution belongs to
    dose_dist: The dose distribution whose dose grid to outline
               If not provided, the current beam set is used
    
    Returns
    -------
    The box ROI
    """
    # Get current case
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('No case is open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    # Get dose grid
    if dose_dist is None:
        # Get current beam set
        try:
            beam_set = get_current('BeamSet')
            dose_dist = beam_set.FractionDose
        except:
            MessageBox.Show('No beam set is loaded. Click OK to abort the script.', 'No Beam Set Loaded')
            sys.exit()
    dg = dose_dist.InDoseGrid
    exam = exam_from_dose(case, dose_dist)
    
    # Get DoseGrid ROI name and color
    roi_name = case.PatientModel.GetUniqueRoiName(DesiredName='zDoseGrid')
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
    if dose_dist.DoseValues is not None:
        dose_dist.UpdateDoseGridStructures()

    return box
