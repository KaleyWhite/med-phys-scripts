import clr
import sys

from connect import *  # Interact w/ RS

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox


def delete_empty_rois() -> None:
    """Deletes all unapproved ROIs in the current case that are empty on all exams
    
    Alerts the user of any empty ROIs that could not be deleted because they are approved
    """
    # Get current case
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]  # All ROI names in current case
    approved_roi_names = [geom.OfRoi.Name for ss in case.PatientModel.StructureSets for approved_ss in ss.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures]  # ROI names whose geometries are approved in any structure set in current case
    approved_roi_names.extend(geom.OfRoi.Name for plan in case.TreatmentPlans for bs in plan.BeamSets if bs.DependentApprovedStructureSet is not None for geom in bs.DependentApprovedStructureSet.ApprovedRoiStructures)  # ROI names whose are part of any structure set that any beam set in the current case depends on
    empty_approved_roi_names = []  # Names of ROIs that could not be deleted because they are approved
    
    # Attempt to delete all empty ROIs
    with CompositeAction('Delete empty ROIs'):
        for roi_name in roi_names:
            if not any(ss.RoiGeometries[roi_name].HasContours() for ss in case.PatientModel.StructureSets):  # ROI has no geometries in any structure set in the current case
                if roi_name in approved_roi_names:
                    empty_approved_roi_names.append(roi_name)  # We will alert user that ROI is approved so could not be deleted 
                else:
                    case.PatientModel.RegionsOfInterest[roi_name].DeleteRoi()
    
    # Alert user if any empty ROIs were not deleted
    if empty_approved_roi_names:
        msg = 'The following ROIs could not be deleted because they are part of approved structure set(s): ' + ', '.join('"' + name + "'" for name in empty_approved_roi_names)
        MessageBox.Show(msg, 'Approved Geometry(ies)')
