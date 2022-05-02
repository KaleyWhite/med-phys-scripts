import clr
import os
import re
import sys

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox


def name_matches(base, name):
    regex = base + r' \(\d+\)'
    m = re.match(regex, name, re.IGNORECASE)
    return m is not None


def exclude_from_mosaiq_export() -> None:
    """Includes/excludes the current case's ROIs from export
    
    Makes visible all included ROIs and makes invisible all excluded ROIs

    Includes and makes visible:
    - SpinalCord
    - Targets
    - Support ROIs
    - External
    - If a plan is open and has dose:
        - Any otherwise excluded ROIs with max dose >= 10 Gy in the current plan

    Exclude and make invisible:
    - ROI named "box" (case insensitive)
    - Control ROIs
    - If a plan is open:
        - ROIs with empty geometries on planning exam of current plan
        - If the current plan has dose:
            - Any otherwise included ROIs with max dose < 10 Gy in the current plan

    Assumes all ROI names are TG-263 compliant
    
    Upon script completion, the user should review the included/excluded ROIs in ROI/POI Details and make any necessary changes
    """
    # Get current variables
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        plan = get_current('Plan')
        dose_dist = plan.TreatmentCourse.TotalDose
        if dose_dist.DoseValues is None:
            dose_dist = None
        struct_set = plan.GetTotalDoseStructureSet()
    except:
        dose_dist = None
    try:
        roi_geoms = plan.GetTotalDoseStructureSet().RoiGeometries
    except:
        try:
            roi_geoms = case.PatientModel.StructureSets[get_current('Examination').Name].RoiGeometries
        except:
            roi_geoms = None
    
    # Lists of OARs to include, other structures to include, and all structures to exclude
    # Separate OARs and other structures so we can remove OAR dependencies 
    include_organs, include_other, exclude = [], [], []  
    
    for roi_name in case.PatientModel.RegionsOfInterest:
        if name_match('box', roi_name) or roi.Type == 'Control':
            exclude.append(roi_name)
        elif roi_geoms is not None:
            geom = roi_geoms[roi_name]

        
        elif  or not geom.HasContours():  # ROI is a control or geometry has no contours
            exclude.append(roi_name)
        elif names_match(roi_name, 'SpinalCord'):
            include_organs.append(roi_name)
        elif geom.OfRoi.OrganData.OrganType == 'Target' or geom.OfRoi.Type in ['Support', 'External']:
            include_other.append(roi_name)
        else:  # include/exclude determined by max dose @ appr vol
            rel_vol = 0.035 / geom.GetRoiVolume()
            max_dose = dose_dist.GetDoseAtRelativeVolumes(RoiName=roi_name, RelativeVolumes=[rel_vol])[0]
            if max_dose >= 1000:
                include_organs.append(roi_name)
            else:
                exclude.append(roi_name)

    # Remove dependent ROIs  
    dependent_rois = [struct_set.RoiGeometries[geom_name].GetDependentRois() for geom_name in include_organs]
    include_organs = [geom for geom in include_organs if geom not in dependent_rois]
    include = include_organs + include_other
 
    # Include/exclude from export and set visibility
    with CompositeAction('Apply ROI changes'):
        case.PatientModel.ToggleExcludeFromExport(ExcludeFromExport=False, RegionOfInterests=include)
        for roi in include:
            patient.SetRoiVisibility(RoiName=roi, IsVisible=True)
        case.PatientModel.ToggleExcludeFromExport(ExcludeFromExport=True, RegionOfInterests=exclude)
        for roi in exclude:
            patient.SetRoiVisibility(RoiName=roi, IsVisible=False)
