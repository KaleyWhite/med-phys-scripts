import clr
import random

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # To display warnings and errors

LUNG_EXPANSION_FOR_CHESTWALL = 3


def contour_chestwall():
    """Creates Chestwall_L and Chestwall_R geometries on the current examination

    Chestwall geometries are based on lung expansions, so lung geometries must exist.
    If an unapproved chestwall ROI exists, overwrite its geometry. Otherwise, create a new chestwall ROI.

    Assumptions
    -----------
    Left and right lung ROI names are TG-263 compliant.
    """
    try:
        case = get_current('Case')
        try:
            exam = get_current('Examination')
        except:
            MessageBox.Show('The current case has no exams. Click OK to abort the script.', 'No Exams')
            sys.exit()
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    struct_set = case.PatientModel.StructureSets[exam.Name]

    # Get external geometry
    try:
        ext_name = next(roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type == 'External')
    except StopIteration:
        MessageBox.Show('There is no external ROI in the current case. Click OK to abort the script.', 'No External ROI')
        sys.exit()

    approved_roi_names = [geom.OfRoi.Name for ss in struct_set.ApprovedStructureSets for geom in ss.ApprovedRoiStructures]

    missing_lung_names = []
    approved_chestwall_names = []
    added_chestwall_geom_names = []
    # Create chestwall contours
    for side in ['L', 'R']:
        # Select lung
        lung_name = 'Lung_' + side
        try:
            lung_geom = struct_set.RoiGeometries[lung_name]
        except KeyError:
            missing_lung_names.append(lung_name)
            continue

        # Chestwall name and color
        chestwall_name = 'Chestwall_' + side  # 'Chestwall_L' or 'Chestwall_R'
        
        # Create/get chestwall ROI
        if chestwall_name in approved_roi_names:
            approved_chestwall_names.append(chestwall_name)
            continue
        try:
            chestwall = struct_set.RoiGeometries[chestwall_name].OfRoi
        except:
            chestwall = case.PatientModel.CreateRoi(Name=chestwall_name, Type='Organ')

        # Create chestwall geometry based on lung
        margin_a = {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': LUNG_EXPANSION_FOR_CHESTWALL, 'Posterior': LUNG_EXPANSION_FOR_CHESTWALL}  # Expand lung posteriorly, anteriorly, and in the `side` direction
        margin_b = {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0}  # Expand lung oppsite the `side` direction
        if side == 'L':  # Left chestwall
            margin_a['Left'] = margin_b['Right'] = LUNG_EXPANSION_FOR_CHESTWALL
            margin_a['Right'] = margin_b['Left'] = 0
        else:  # Right chestwall
            margin_a['Right'] = margin_b['Left'] = LUNG_EXPANSION_FOR_CHESTWALL
            margin_a['Left'] = margin_b['Right'] = 0
        chestwall.CreateAlgebraGeometry(Examination=exam, ExpressionA={'Operation': 'Union', 'SourceRoiNames': [lung_name], 'MarginSettings': margin_a}, ExpressionB={'Operation': 'Union', 'SourceRoiNames': [lung_name], 'MarginSettings': margin_b}, ResultOperation='Subtraction', ResultMarginSettings={'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})
 
        # Remove parts of chestwall that extend outside the external
        chestwall.CreateAlgebraGeometry(Examination=exam, ExpressionA={'Operation': 'Union', 'SourceRoiNames': [chestwall.Name], 'MarginSettings': {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ExpressionB={'Operation': 'Union', 'SourceRoiNames': [ext_name], 'MarginSettings': {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ResultOperation='Intersection', ResultMarginSettings={'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})
        added_chestwall_geom_names.append(chestwall.Name)
    
    if added_chestwall_geom_names:
        struct_set.SimplifyContours(RoiNames=added_chestwall_geom_names, RemoveHoles3D=True, RemoveSmallContours=True, AreaThreshold=0.01, ResolveOverlappingContours=True)
        # For some reason, the above does not remove chestwall overlap, so subtract one chestwall from the other
        if len(added_chestwall_geom_names) == 2:
            minuend, subtrahend = random.shuffle(added_chestwall_geom_names)
            case.PatientModel.RegionsOfInterest[minuend].CreateAlgebraGeometry(Examination=exam, Algorithm='Auto', ExpressionA={'Operation': 'Union', 'SourceRoiNames': [minuend], 'MarginSettings': {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 }}, ExpressionB={'Operation': 'Union', 'SourceRoiNames': [subtrahend], 'MarginSettings': {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 }}, ResultOperation='Subtraction', ResultMarginSettings={'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})
    
    warnings = ''

    if len(missing_lung_names) == 1:
        missing_lung_name = missing_lung_names[0]
        warnings += '\nThere is no ' + missing_lung_name + ' ROI in the current case, so no Chestwall_' + missing_lung_name[:-2] + ' was added.'
    elif len(missing_lung_names) == 2:
        warnings += '\nNeither Lung_L nor Lung_R is present in the current case, so no chestwalls were added.'
    
    if len(approved_chestwall_names) == 1:
        approved_chestwall_name = approved_chestwall_names[0]
        warnings += '\nThe ' + approved_chestwall_name + ' geometry is approved on the current exam, so it could not be changed.'
    elif len(approved_chestwall_names) == 2:
        warnings += 'Both Chestwall_L and Chestwall_R geometries are approved on the current exam, so no chestwall geometries were changed.'

    if warnings:
        MessageBox.Show(warnings)
