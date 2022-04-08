import clr
import sys

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # For displaying errors


def dose_grid_box() :
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
    
    # Get DoseGrid ROI if it exists; create it otherwise
    try:
        box = case.PatientModel.RegionsOfInterest['DoseGrid']
        # Exit script if geometry is approved
        for approved_ss in struct_set.ApprovedStructureSets:
            for approved_geom in approved_ss.ApprovedRoiStructures:
                if approved_geom.Name == 'DoseGrid':
                    MessageBox.Show('DoseGrid geometry is approved, so it cannot be changed. Click OK to abort the script.', 'DoseGrid Is Approved')
                    sys.exit()
    except:
        box = case.PatientModel.CreateRoi(Name='DoseGrid', Color='purple', Type='Control')
    
    # Box size and center coordinates
    box_sz, box_ctr = {}, {}
    for dim, coord in dg.Corner.items():
        box_sz[dim] = dg.NrVoxels[dim] * dg.VoxelSize[dim]
        box_ctr[dim] = coord + (box_sz[dim] / 2)

    # Create box and update its geometry (must update all geometries at once)
    box.CreateBoxGeometry(Size=box_sz, Examination=exam, Center=box_ctr, VoxelSize=dg.VoxelSize.x)
    if dose.DoseValues is not None:
        dose.UpdateDoseGridStructures()
