import clr
import os
import sys

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # For displaying errors


INNER_COUCH_NAME = 'Elekta Couch Foam Core'
OUTER_COUCH_NAME = 'Elekta Carbon Fiber Shell'


def add_box_to_external():
    """Adds a box to the external geometry on the current exam.

    Box is as wide as the inner couch geometry, its height extends from the top of the couch to the loc point, and its depth is the depth of the external geometry.
    Useful for including the Vac-Lok for SBRT lung plans.
    """

    # Ensure a case is open
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('No case is open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    # Ensure the case has an exam
    try:
        exam = get_current('Examination')
    except:
        MessageBox.Show('There are no exams in the open case. Click OK to abort the script.', 'No Exams')
        sys.exit()
    
    struct_set = case.PatientModel.StructureSets[exam.Name]

    # Get external ROI
    try:
        ext_geom = next(geom for geom in struct_set.RoiGeometries if geom.OfRoi.Type == 'External')
        ext = ext_geom.OfRoi
    except StopIteration:
        MessageBox.Show('There is no external geometry on the current exam. Click OK to abort the script.', 'No External Geometry')
        sys.exit()

    # Will need to check whether various ROIs are approved
    approved_roi_names = []
    for approved_ss in struct_set.ApprovedStructureSets:
        for approved_roi_geom in approved_ss.ApprovedRoiStructures:
            if approved_roi_geom.OfRoi.Name not in approved_roi_names:
                approved_roi_names.append(approved_roi_geom.OfRoi.Name)

    # Ensure external is unapproved
    if ext.Name in approved_roi_names:
        MessageBox.Show('External geometry is approved on the current exam, so the geometry cannot be changed. Click OK to abort the script.', 'External Geometry Is Approved')
        sys.exit()

    # Get couch ROIs
    try:
        inner_couch = struct_set.RoiGeometries[INNER_COUCH_NAME]
        outer_couch = struct_set.RoiGeometries[OUTER_COUCH_NAME]
    except:
        MessageBox.Show('One or both of ' + INNER_COUCH_NAME + ' and ' + OUTER_COUCH_NAME + ' ROIs do not exist in the current case. Click OK to abort the script.', 'Missing Couch ROI(s)')
        sys.exit()
    # Get couch geometries
    if not inner_couch.HasContours() or not outer_couch.HasContours():
        MessageBox.Show('One or both of ' + INNER_COUCH_NAME + ' and ' + OUTER_COUCH_NAME + ' geometries do not exist on the current exam. Click OK to abort the script.', 'Missing Couch Geometry(ies)')
        sys.exit()

    # RL width = exam width minus a px on each side (RS throws error if external is as wide as image)
    px_sz = exam.Series[0].ImageStack.PixelSize.x
    exam_bb = exam.Series[0].ImageStack.GetBoundingBox()
    exam_min_x, exam_max_x = exam_bb[0].x + px_sz, exam_bb[1].x - px_sz
    couch_bb = inner_couch.GetBoundingBox()
    couch_min_x, couch_max_x = couch_bb[0].x, couch_bb[1].x
    x = min(couch_max_x, exam_max_x) - max(couch_min_x, exam_min_x)

    # PA height
    loc_y = struct_set.LocalizationPoiGeometry
    if loc_y is None or loc_y.Point is None or abs(loc_y.Point.x) == float('Inf'):  # No localization point, or localization geometry has infinite coordinates
        MessageBox.Show('There is no localization geometry on the current exam. Click OK to abort the script.', 'No Localization Geometry')
        sys.exit()
    loc_y = loc_y.Point.y  # Loc point
    couch_y = outer_couch.GetBoundingBox()[exam.PatientPosition.endswith('P')].y  # Top of couch is larger y for prone patient, smaller y for supine
    y = abs(couch_y - loc_y)

    # IS height = height of External
    ext_min_z = ext_geom.GetBoundingBox()[0].z
    ext_max_z = ext_geom.GetBoundingBox()[1].z
    z = ext_max_z - ext_min_z

    # Box center (x center = 0)
    z_ctr = (ext_max_z + ext_min_z) / 2
    y_ctr = (couch_y + loc_y) / 2  

    # Create box geometry
    # Get box ROI, or create it if it doesn't exist
    try:
        box = case.PatientModel.RegionsOfInterest['Box']
        # Ensure box is unapproved
        for box.Name in approved_roi_names:
            MessageBox.Show('Box geometry is approved on the current exam, so the geometry cannot be changed. Click OK to abort the script.', 'Box Geometry Is Approved')
            sys.exit()
    except:
        box = case.PatientModel.CreateRoi(Name='Box', Type='Control')
    box.CreateBoxGeometry(Size={'x': x, 'y': y, 'z': z}, Examination=exam, Center={'x': 0, 'y': y_ctr, 'z': z_ctr})
    
    # Copy external into new ROI 'External^NoBox'
    # Get External^NoBox, or create it if it doesn't exist
    try:
        ext_no_box = case.PatientModel.RegionsOfInterest['External^NoBox']
        # Ensure External^NoBox is unapproved
        for ext_no_box.Name in approved_roi_names:
            MessageBox.Show('External^NoBox geometry is approved on the current exam, so the geometry cannot be changed. Click OK to abort the script.', 'External^NoBox Geometry Is Approved')
            sys.exit()
    except:
        ext_no_box = case.PatientModel.CreateRoi(Name='External^NoBox', Type='Organ', Color='white')
    ext_no_box.CreateMarginGeometry(Examination=exam, SourceRoiName=ext.Name, MarginSettings={'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})
    
    # Add box to external
    ext.CreateAlgebraGeometry(Examination=exam, ExpressionA={'Operation': 'Union', 'SourceRoiNames': [ext.Name], 'MarginSettings': {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ExpressionB={'Operation': 'Union', 'SourceRoiNames': [box.Name], 'MarginSettings': {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ResultOperation='Union', ResultMarginSettings={'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

    # Delete unnecessary ROI
    box.DeleteRoi()

    # Due to new geometries, update dose structures for all unapproved plans on the current exam
    for plan in case.TreatmentPlans:
        if (plan.Review is None or plan.Review.ApprovalStatus != 'Approved') and plan.GetTotalDoseStructureSet().OnExamination.Equals(exam):
            plan.TreatmentCourse.TotalDose.UpdateDoseGridStructures()
