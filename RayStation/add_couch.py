import clr
import os
import re
import shutil
import sys

from connect import *  # Interact w/ RS
from pydicom import dcmread  # Read DICOM from file
from scipy.signal import find_peaks  # Find "peak" values in an array

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # I use to display errors


def add_couch():
    """Adds template couch structures at correct location on current exam

    Correct location is R-L center, I-S center, and aligned with sim couch P-A
    Alert user of any couch geometries in the incorrect position that can't be moved because the geometries are approved
    Set default dose grid for any non-initial sim plan on the current exam. An initial sim plan name contains 'initial sim' or 'trial_1' (case insensitive).
    
    Assumes patient position of current exam is HFS, HFP, or FFS
    """

    # Get current variables
    try:
        case = get_current('Case')
        try:
            exam = get_current('Examination')
        except:
            MessageBox.Show('There are no exams in the current case.', 'Click OK to abort the script.', 'No Examinations')
            sys.exit()
    except:
        MessageBox.Show('There is no case loaded. Click OK to abort the script.', 'No Case Loaded')
        sys.exit()
    patient_db = get_current('PatientDB')
    patient = get_current('Patient')
    struct_set = case.PatientModel.StructureSets[exam.Name]
    
    # Patient position
    pt_pos = exam.PatientPosition
    if pt_pos not in ['HFS', 'HFP', 'FFS']:
        MessageBox.Show('This script does not support ' + pt_pos + ' exams. Click OK to abort the script.', 'Unsupported Patient Position')
        sys.exit()
    is_supine = exam.PatientPosition in ['HFS', 'FFS']

    img_stack = exam.Series[0].ImageStack

    # Later, we will alert user of geometries that couldn't be moved
    approved_roi_names = [geom.OfRoi.Name for ss in struct_set.ApprovedStructureSets for geom in ss.ApprovedRoiStructures]
    all_roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
    approved_couch_incorrect = []  # Geometry names that cannot be moved to correct location because they are approved

    # Density values are cleared and dose values are invalidated, so recompute dose on unapproved beam sets that have dose and that are on the current examination
    to_recompute = [bs for p in case.TreatmentPlans for bs in p.BeamSets if bs.GetPlanningExamination().Equals(exam) and (p.Review is None or p.Review.ApprovalStatus != 'Approved') and bs.FractionDose.DoseValues is not None and bs.FractionDose.DoseValues.IsClinical]

    # Couch template and structure names
    template_name = 'Elekta Couch' if is_supine else 'Elekta Prone Couch'
    template = patient_db.LoadTemplatePatientModel(templateName=template_name)
    couch_structs = template.PatientModel.RegionsOfInterest
    outer_couch, inner_couch = couch_structs

    # Attempting to apply a template to a structure set that is part of an approved plan gives an error
    # So ensure that, if either of the couch ROIs does not exist, 
    if outer_couch.Name in approved_roi_names or inner_couch.Name in approved_roi_names:
        MessageBox.Show('One or more couch structures is approved on the current exam. Click OK to abort the script.', 'Caouch Geometry(ies) Are Approved')
        sys.exit()
    
    # Apply template
    # Doesn't matter which source exam we use, so just use the first
    case.PatientModel.CreateStructuresFromTemplate(SourceTemplate=template, SourceExaminationName=template.StructureSetExaminations[0].Name, SourceRoiNames=[outer_couch.Name, inner_couch.Name], SourcePoiNames=[], TargetExamination=exam, AssociateStructuresByName=True, InitializationOption='AlignImageCenters')

    ## Export exam so we can use pydicom to extract pixel data
    # Create export folder
    export_path = r'T:\Physics\Temp\AddCouchScript'
    if os.path.isdir(export_path):
        shutil.rmtree(export_path)
    os.makedirs(export_path)

    # Export exam
    # No RS functionality to export a single CT slice
    patient.Save()  # Error if you don't save before export
    try:
        case.ScriptableDicomExport(ExportFolderPath=export_path, Examinations=[exam.Name], IgnorePreConditionWarnings=False)
    except:
        case.ScriptableDicomExport(ExportFolderPath=export_path, Examinations=[exam.Name], IgnorePreConditionWarnings=True)

    ## Get y-coordinate of top of sim couch
    # If entire sim couch is in scan, the couch top is the fourth 'bright' peak in the R-L middle

    # Get pixel intensities array
    dcm = dcmread(export_path + r'\\' + os.listdir(export_path)[0])
    intensities = dcm.pixel_array

    # Select intensities in R-L center
    y_axis_intensities = [row[intensities.shape[0] // 2] for row in intensities]

    # Select 'bright peaks' on y-axis
    # These are white areas on image
    # The `prominence` argument was obtained through trial and error
    # `find_peaks` returns a tuple whose first element is the indices of the peaks
    y_axis_peaks = find_peaks(y_axis_intensities, height=(0, 700), prominence=180)[0]

    # Select 4th peak from the bottom
    # Supine: 4th from the end
    # Prone: 4th from the beginning
    second_peak = y_axis_peaks[-1] if is_supine else y_axis_peaks[-1]

    # Convert pixel y-coordinate to exam coordinate
    #correct_y = img_stack.Corner.y + second_peak * img_stack.PixelSize.y
    if is_supine:
        correct_y = img_stack.Corner.y + second_peak * img_stack.PixelSize.y - 2.25
    else:
        correct_y = img_stack.Corner.y - second_peak * img_stack.PixelSize.y + 2.25
    
    # The TOP of the couch goes at the 4th peak, so compute where the P-A CENTER of the couch should go
    # Center is half a couch height down
    outer_couch_bounds = struct_set.RoiGeometries[outer_couch.Name].GetBoundingBox()
    outer_couch_ht = outer_couch_bounds[1].y - outer_couch_bounds[0].y
    if is_supine:
        correct_y += outer_couch_ht / 2
    else:
        correct_y -= outer_couch_ht / 2

    # Correct x- and z-coordinates
    # x is R-L center (always zero)
    # z is I-S center
    correct_x = 0
    img_bounds = img_stack.GetBoundingBox()
    correct_z = (img_bounds[0].z + img_bounds[1].z) / 2

    # Center each couch geometry
    for roi in (inner_couch, outer_couch):  # Iterate over couch ROIs in the applied couch template
        geom = struct_set.RoiGeometries[roi.Name]  # That couch ROI's geometry on the exam

        # Current position
        geom_ctr = geom.GetCenterOfRoi()

        # If geometry is approved, add to list that user will be alerted about
        # 'Correct position' is within 1 mm. A tolerance is especially important due to Python's weird floating point errors
        if roi.Name in approved_roi_names and (round(correct_x, 2) != round(geom_ctr.x, 2) or round(correct_y, 2) != round(geom_ctr.y, 2) or round(correct_z, 2) != round(geom_ctr.z, 2)):
            approved_couch_incorrect.append(roi.Name)
        elif roi.Name not in approved_roi_names:  # Geometry is unapproved, so we can move it
            # Transformation matrix
            # Each x, y, and z transform is the difference between the correct and current coordinates
            mat = {'M11': 1, 'M12': 0, 'M13': 0, 'M14': correct_x - geom_ctr.x,
                'M21': 0, 'M22': 1, 'M23': 0, 'M24': correct_y - geom_ctr.y,
                'M31': 0, 'M32': 0, 'M33': 1, 'M34': correct_z - geom_ctr.z, 
                'M41': 0, 'M42': 0, 'M43': 0, 'M44': 1}

            # Reposition couch structure
            geom.OfRoi.TransformROI3D(Examination=exam, TransformationMatrix=mat)

            # Show structure
            patient.SetRoiVisibility(RoiName=roi.Name, IsVisible=True)

    # For any plan on this exam that has no dose, set default dose grid
    for p in case.TreatmentPlans:
        if not re.search(r'(initial )?sim|trial.?1', p.Name, re.IGNORECASE) and p.GetTotalDoseStructureSet().OnExamination.Equals(exam) and p.TreatmentCourse.TotalDose.DoseValues is None:
            voxel_sz = p.GetTotalDoseGrid().VoxelSize
            if any(val == 0 for val in voxel_sz.values()):
                voxel_sz = {coord: 0.3 for coord in 'xyz'}
            for bs in p.BeamSets:
                bs.SetDefaultDoseGrid(VoxelSize=voxel_sz)
                bs.FractionDose.UpdateDoseGridStructures()

    # Recompute invalidated beam sets
    for bs in to_recompute:
        bs.ComputeDose(ComputeBeamDoses=True, DoseAlgorithm=bs.AccurateDoseAlgorithm.DoseAlgorithm)

    # Alert user if any approved geometries (can't be moved) are in incorrect positions
    if len(approved_couch_incorrect) == 1:
        MessageBox.Show('The geometry for "' + approved_couch_incorrect[0] + '" is approved, so it cannot be moved to the correct position.', 'Cannot Move Geometry')
    elif len(approved_couch_incorrect) == 2:
        MessageBox.Show('The couch geometries are approved, so they cannot be moved to the correct positions.', 'Cannot Move Geometries')

    # Remove unnecessary exported files
    shutil.rmtree(export_path)
