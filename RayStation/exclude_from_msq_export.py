import clr
import os
import re
import sys

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox


def name_matches(base: str, name: str) -> bool:
    """Determines whether a name matches the given name (case insensitive), possibly with a copy number in parentheses

    Arguments
    ---------
    base: The "base name" (without a copy number) to compare the name to
    name: The name to compare to the "base" name

    Returns:
    True if the name exactly matches the base or is the base plus a copy number, False otherwise

    Examples
    --------
    name_matches("SpinalCord", "Spinalcord (1)") -> True
    name_matches("SpinalCord", "Spinal Cord") -> False
    name_matches("SpinalCord", "Spinal Cord (1)") -> False
    """
    regex = base + r'( \(\d+\))?'
    m = re.match(regex, name, re.IGNORECASE)
    return m is not None


def hi_max_dose(dose_dist, geom):
    """Determines whether Dmax (0.035) for the geometry in the dose distribution is at least 10 Gy

    Arguments
    ---------
    dose_dist: The dose distribution that the geometry belongs to
    geom: The ROI geometry

    Returns
    -------
    True if D0.035 >= 10 Gy, False otherwise
    """
    if dose_dist is None:
        return False
    rel_vol = 0.035 / geom.GetRoiVolume()
    max_dose = dose_dist.GetDoseAtRelativeVolumes(RoiName=geom.OfRoi.Name, RelativeVolumes=[rel_vol])[0]
    return max_dose >= 1000


def exclude_from_msq_export() -> None:
    """Includes/excludes the current case's ROIs from export.
    
    Makes visible all included ROIs and makes invisible all excluded ROIs.
    
    Includes SpinalCord, Targets, External and ROIs with materials defined.
    Included ROIs must also be nonempty (planning exam if a plan is open, ecurrent exam otherwise. If neither plan nor exam is open, ignore empty/nonempty rule).
    If the open plan has dose, included ROIs mcut also have max dose >= 10 Gy.

    Excludes "Box" and Control.

    Assumes all ROI names are TG-263 compliant. ROI names with a copy number are considered matches (e.g., "SpinalCord (1)" is "SpinalCord").
    
    Upon script completion, the user should review the included/excluded ROIs in ROI/POI Details and make any necessary changes.
    """
    ## Get current variables
    # Patient
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    # Case
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    # Dose distribution (for checking max geometry dose) and ROI geometries
    dose_dist = roi_geoms = None
    try:
        # If a plan is open, there may be dose, and use the plan's structure set
        plan = get_current('Plan')
        if plan.TreatmentCourse.TotalDose.DoseValues is not None:
            dose_dist = plan.TreatmentCourse.TotalDose
        roi_geoms = plan.GetTotalDoseStructureSet().RoiGeometries
    except:
        # Use current exam's structure set if an exam is open
        try:
            roi_geoms = case.PatientModel.StructureSets[get_current('Examination').Name].RoiGeometries
        except:
            pass
    
    # Lists of ROI names to include/exclude
    include, exclude = [], []
    
    # Add each ROI name to include or exclude list
    for roi in case.PatientModel.RegionsOfInterest:
        # Include if material override
        if roi.RoiMaterial is not None:
            include.append(roi.Name)
        # Exclude "Box" or Control
        elif name_matches('box', roi.Name) or roi.Type == 'Control':
            exclude.append(roi.Name)
        # Include anything else if there is no dose or structure set (so we can't check is nonempty or has high max dose)
        elif roi_geoms is None:
            include.append(roi.Name)
        # Check structure set and/or dose distribution to determine whether to include
        else:
            geom = roi_geoms[roi.Name]
            if geom.HasContours():  # Nonempty
                # Include SpinalCord, target, external, if no dose (so can't check max dose), or if max dose is high
                if name_matches('SpinalCord', roi.Name) or roi.OrganData.OrganType == 'Target' or roi.Type == 'External' or dose_dist is None or hi_max_dose(dose_dist, geom):
                    include.append(roi.Name)
                else:
                    exclude.append(roi.Name)
            # Exclude empty
            else:
                exclude.append(roi.Name)
 
    # Include/exclude from export and set visibility
    with CompositeAction('Apply ROI changes'):
        case.PatientModel.ToggleExcludeFromExport(ExcludeFromExport=False, RegionOfInterests=include)
        for roi_name in include:
            patient.SetRoiVisibility(RoiName=roi_name, IsVisible=True)
        case.PatientModel.ToggleExcludeFromExport(ExcludeFromExport=True, RegionOfInterests=exclude)
        for roi_name in exclude:
            patient.SetRoiVisibility(RoiName=roi_name, IsVisible=False)
