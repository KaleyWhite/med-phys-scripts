import clr
import os
import sys

from connect import *   # Interact with RayStation

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # For displaying errors/warnings

sys.path.append(os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-scripts', 'RayStation'))
from center_geometries import center_geometries


def center_couch():
    """Centers all couch geometries on the current exam

    Checks all couch names in the Elekta couch structure templates ("Elekta Couch" and "Elekta Prone Couch") and centers existing couch geometries that are unapproved and nonempty on the current exam

    Alerts the user if any geometries could not be added because they are approved or empty
    """
    # Get current objects
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        exam = get_current('Examination')
    except:
        MessageBox.Show('There are no exams in the open case. Click OK to abort the script.', 'No Exams')
        sys.exit()
    patient_db = get_current('PatientDB')
    
    struct_set = case.PatientModel.StructureSets[exam.Name]
    
    all_roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
    approved_roi_names = list(set(geom.OfRoi.Name for approved_ss in struct_set.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures))

    approved_couch, empty_couch = [], []  # Names of couch structures that can't be added because they're approved or empty, respectively, on the current exam
    ok_couch = []  # Names of couch geometries to center

    # For each couch name in the Elekta couch templates, add to the appropriate list
    for template_name in ['Elekta Couch', 'Elekta Prone Couch']:
        template = patient_db.LoadTemplatePatientModel(templateName=template_name)
        for roi in template.PatientModel.RegionsOfInterest:
            if roi.Name not in all_roi_names:  # Ignore ROIs not in current case
                continue
            if roi.Name in approved_roi_names:
                approved_couch.append(roi.Name)
            elif not struct_set.RoiGeometries[roi.Name].HasContours():
                empty_couch.append(roi.Name)
            else:
                ok_couch.append(roi.Name)
    
    # Remove duplicate couch names in lists
    approved_couch = list(set(approved_couch))
    empty_couch = list(set(empty_couch))
    ok_couch = list(set(ok_couch))

    # Exit script if no couch ROIs or centerable geometries
    if not ok_couch:
        if approved_couch or empty_couch:
            MessageBox.Show('There are no unapproved, nonempty couch geometries on the current exam. Click OK to abort the script.', 'No Couch Geometries')   
        else:
            MessageBox.Show('There are no couch ROIs in the current case. Click OK to abort the script.', 'No Couch ROIs') 
        sys.exit()
    
    # Center each couch geometry in the R-L direction
    center_info = {couch: [True, False, False] for couch in ok_couch}
    center_geometries(center_info)

    # Display warnings if any couch geometries could not be centered
    if approved_couch or empty_couch:
        warnings = 'The following couch geometries could be centered on the current exam:\n'
        if approved_couch:
            warnings += '\t-  Because they are approved: ' + ', '.join('"' + couch + '"' for couch in approved_couch)
        if empty_couch:
            warnings += '\n\t-  Because they are empty: ' + ', '.join('"' + couch + '"' for couch in empty_couch)
