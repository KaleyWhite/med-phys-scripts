import clr
from collections import OrderedDict
from datetime import datetime
import math
import os
import re
import sys
from typing import Dict, List, Optional

from connect import *
from connect.connect_cpython import PyScriptObject

import reportlab.lib.colors
from reportlab.lib.colors import Blacker, Whiter, blue, green, red, yellow
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, KeepTogether
from reportlab.lib.units import inch

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import *


class PlanChkConstants(object):
    """Class that defines several useful constants for this remainder of the script"""

    # ---------------------------------------------------------------------------- #
    #                           Clinic-specific constants  
    #                                CHANGE THESE!
    # ---------------------------------------------------------------------------- #
    
    # Absolute path to the directory in which to create the PDF report
    # This directory does not have to already exist
    OUTPUT_DIR = os.path.join('T:', os.sep, 'Physics', 'Scripts', 'Output Files', 'Plan Check')

    # Desired imaging system name for exams
    IMG_SYS = 'HOST-7307'

    # ---------------------------- Regular expressions --------------------------- #

    DATE_REGEX = r'(\d{1,2}[/\-. ]\d{1,2}[/\-. ](\d{4}|\d{2}))|(\d{1,2}[/\-. ](Jan(uary)?|Feb(uary)?|Mar(ch)?|Apr(il)?|May|June?|July?|Aug(ust)?|Sep(t(ember)?)?|Oct(ober)?|Nov(ember)?|Dec(ember)?)[/\-. ](\d{4}|\d{2}))|(\d{9})'  # For checking for date in exam names
    MRN_REGEX = r'0{3}\d{6}'  # Patient ID format
    PROS_PLAN_REGEX = r'(^|[^A-Za-z])(pros|pb|bed|fossa)([^A-Za-z]|$)'  # Should be present in case name, case comments, plan name, beam set name, etc. of prostate plans
    PT_NAME_REGEX = r'[A-Z][A-Za-z\']*(\^+[A-Z][A-Za-z\']*)+'  # Patient name format. Probably won't need changing as it's just the basic DICOM Patient Name attribute format
    IMG_FOR_TEMPLATES_REGEX = MRN_REGEX + ' IMAGE FOR TEMPLATES'  # Exams used only for structure templates. These are ignored in several exam checks
    INI_SIM_PLAN_REGEX = r'(initial sim)|(trial_1)'  # An "initial sim" plan name contains this
    CBCT_REGEX = r'CB(CT)?'  # A CBCT setup beam name contains this

    # ---------------------------------------------------------------------------- #
    #                             No changes necessary                             #
    # ---------------------------------------------------------------------------- #

    # Paths to Adobe Reader on RS servers
    ADOBE_READER_PATHS = [os.path.join('C:', os.sep, 'Program Files (x86)', 'Adobe', 'Reader 11.0', 'Reader', 'AcroRd32.exe'), os.path.join('C:', os.sep, 'Program Files (x86)', 'Adobe', 'Acrobat Reader DC', 'Reader', 'AcroRd32.exe')]

    # --------------------------------- ReportLab -------------------------------- #

    # Colors for each check "state"
    PASS = 'green'
    WARN = 'yellow'
    FAIL = 'red'
    MANUAL = 'blue'
    COLORS = [PASS, WARN, FAIL, MANUAL]

    # Styles
    STYLES = getSampleStyleSheet()  # Base styles (e.g., 'Heading1', 'Normal')
    for color in COLORS:
        reportlab_color = getattr(reportlab.lib.colors, color)
        STYLES.add(ParagraphStyle(name=color, parent=STYLES['Normal'], backColor=Whiter(reportlab_color, 0.25), borderPadding=7, borderWidth=1, borderColor=reportlab_color, borderRadius=5))

    # Spacers to avoid overlap of checks in the report
    _, width = letter  # Need the width (8.5') for Spacer objects
    SPCR_SM = Spacer(width, 0.1 * inch)  # Small
    SPCR_LG = Spacer(width, 0.3 * inch)  # Large

    # HTML bulleted list formatting for ReportLab Paragraph text
    NEWLINE_AND_BULLET = '<br/>' + '&nbsp;' * 4 + '&bull;' + '&nbsp;' * 2


def distance(a: Dict[str, float], b: Optional[Dict[str, float]] = {'x': 0, 'y': 0, 'z': 0}) -> float:
    """Computes the Euclidean distance between two points a and b in 3D space

    Arguments
    ---------
    a: Coordinates of the first point
       A dictionary with dimensions as keys and coordinates as values
    b: Coordinates of the second point.
       A dictionary with dimensions as keys and coordinates as values
       Defaults to the origin
       `a` and `b` must have the same keys

    Returns
    -------
    The distance between the two points

    Examples
    --------
    distance({'x': 2, 'y': 3, 'z': -4}, {'x': 0, 'y': 1, 'z': 1}) -> 5.74...
    distance({'x': 2, 'y': 3, 'z': -4}) -> 5.38...
    """

    return math.sqrt(sum((val - b[coord]) ** 2 for coord, val in a.items()))


def get_contour_coords(case: PyScriptObject, exam: PyScriptObject, geom: PyScriptObject) -> List[Dict[str, float]]:
    """Gets a list of the contour coordinates of an ROI geometry

    Arguments
    ---------
    case: The case to which the geometry belongs
    exam: The exam to which the geometry belongs
    geom: The geometry

    Returns
    -------
    A list of the 3D points that comprise the geometry
    Each point is a dictionary with dimensions as keys and coordinates as values
    If the geometry is empty, returns an empty list

    Example
    -------
    get_contour_coords(some_case, some_exam, some_geom) -> {'x': 0, 'y': 1, 'z': 2}, ...]
    """

    if not geom.HasContours():  # Empty geometry
        return []
    
    # If has contour representation, just flatten contour coords array and return
    # Otherwise, copy ROI, set copy's representation to contours, delete the copy, and return the copy's flattened contours array
    if hasattr(geom.PrimaryShape, 'Contours'):  # Contour representation
        coords = [c for coord in geom.PrimaryShape.Contours for c in coord]  # Flatten contours array
    else:
        copy_name = case.PatientModel.GetUniqueRoiName(DesiredName=geom.OfRoi.Name)
        copy = case.PatientModel.CreateRoi(Name=copy_name, Color=geom.OfRoi.Color, Type=geom.OfRoi.Type)  # Create ROI w/ same color and type as geom's ROI
        copy.CreateAlgebraGeometry(Examination=exam, ExpressionA={ 'Operation': 'Union', 'SourceRoiNames': [geom.OfRoi.Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation='None', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})  # Copy has same geometry
        copy_geom = case.PatientModel.StructureSets[exam.Name].RoiGeometries[copy_name]
        copy_eom.SetRepresentation(Representation='Contours')  # Convert to contour representation
        coords = [c for coord in geom.PrimaryShape.Contours for c in coord]  # Flatten contours array
        copy.DeleteRoi()  # We no longer need the copy
    
    return coords


def goal_str(goal: PyScriptObject) -> str:
    """Creates a string representation of the clinical goal, including its value in parentheses after the main goal text

    Attempts to mimic whatever RayStation does behind the scenes

    Arguments
    ---------
    goal: The goal to format as a string

    Returns
    -------
    The goal string

    Raises
    ------
    ValueError: If the dose type is not recognized

    Example
    -------
    Given some_goal V2000cGy < 0.25cm^3 for ROI "Esophagus", with value 0.2567 cm^3:
    goal_str(some_goal) -> "Esophagus: At most 0.25 cm^3 volume at 2000 cGy dose (0.26 cm^3)"
    """
    goal_type = goal.PlanningGoal.Type

    goal_criteria = goal.ForRegionOfInterest.Name + ': '
    goal_criteria += 'At least' if goal.PlanningGoal.GoalCriteria == 'AtLeast' else 'At most'
    accept_lvl = goal.PlanningGoal.AcceptanceLevel
    param_val = goal.PlanningGoal.ParameterValue
    goal_val = goal.GetClinicalGoalValue()

    goal_txt = goal_criteria + ' '

    if goal_type == 'DoseAtAbsoluteVolume':
        goal_txt += f'{accept_lvl:.0f} cGy dose at {param_val:.2f} cm^3 volume ({goal_val:.0f} cGy)'
    elif goal_type == 'DoseAtVolume':
        goal_txt += f'{accept_lvl:.0f} cGy dose at {param_val * 100:.2f}% volume ({goal_val:.0f} cGy)'
    elif goal_type == 'AbsoluteVolumeAtDose':
        goal_txt += f'{accept_lvl:.2f} cm^3 volume at {param_val:.0f} cGy dose ({accept_lvl:.2f} cm^3)'
    elif goal_type == 'VolumeAtDose':
        goal_txt += f'{accept_lvl * 100:.2f}% volume at {param_val:.0f} cGy dose ({goal_val * 100:.2f}%)'
    elif goal_type == 'AverageDose':
        goal_txt += f'{param_val:.0f} cGy average dose ({goal_val:.0f} cGy)'
    elif goal_type == 'ConformityIndex':
        goal_txt += f'a conformity index of {accept_lvl:.2f} at {param_val:.0f} cGy dose ({goal_val:.2f})'
    elif goal_type == 'HomogeneityIndex':
        goal_txt += f'a homogeneity index of {accept_lvl:.2f} at {param_val * 100:.2f}% volume ({goal_val:.2f})'
    elif goal_type == 'DoseAtPoint':  # DoseAtPoint
        goal_txt += f'{accept_lvl:.0f} cGy dose at point ({goal_val:.2f})'
    else:
        raise ValueError('Unrecognized goal type "' + goal_type + '"')

    return goal_txt


def format_coords(point: Dict[str, float]) -> str:
    """Formats IEC 61217 patient coordinates to DICOM patient coordinates and nicely formats for display

    Argument
    --------
    point: The coordinates to format
           A dictionary with dimensions 'x', 'y', and 'z' as keys, and coordinates as values

    Returns
    -------
    The DICOM patient coordinates, fomatted as an ordered triplet (in parentheses)

    Example
    -------
    format_coords({'x': 1, 'y': 2, 'z': 3}) -> '(1, 3, -2)'
    """
    return f'({format_num(point["x"])}, {format_num(point["z"])}, {format_num(-point["y"])})'


def format_name(name: str) -> str:
    """Converts the patient's Name attribute into a better format for display

    Arguments
    ---------
    pt: The patient whose name to format

    Returns
    -------
    The formatted patient name

    Example
    -------
    Given some_pt with Name '^Jones^Bill^^M':
    format_pt_name(some_pt) -> 'Jones, Bill M'
    """
    parts = [part for part in re.split(r'\^+', name) if part != '']
    name = parts[0]
    if len(parts) > 0:
        name += ', ' + ' '.join(parts[1:])
    return name


def format_num(num: float) -> str:
    """Formats a number as a display string

    If number between 0 and 0.005 (exclusive), or greater than 100000, format in scientific notation
    Number itself or base is stripped of leading zeroes and trailing zeroes after the decimal place
    If no numbers after the decimal place, removes the decimal point. Otherwise, rounds to 2 decimal places
    Superscripts are HTML <sup> elements
    
    If there are no digits after the decimal place, return an int. Otherwise, strip all zeroes from the end and display a maximum of 2 decimal places

    Argument
    --------
    num: The number to format

    Returns
    -------
    The formatted number string
    """
    if 0 < num < 0.005 or num > 100000:  # Small or large numbers should be in scientific notation
        num = f'{num:.2E}'  # Scientific notation w/ base rounded to 2 decimal places
        coef, exp = num.split('E')  # We will format base and exponent separately
        coef = coef.rstrip('0').rstrip('.')  # Remove trailing zeroes and trailing decimal place from coefficient
        if exp.startswith('+'):
            exp = exp[1:].lstrip('0')  # Remove sign and leading zeroes from positive exponent
        else:
            exp = f'-{exp[1:].lstrip("0")}'  # Remove leading zeroes from negative exponent
        if coef == '1':
            return f'10<sup>{exp}</sup>' # Remove coefficient of 1
        if coef == '-1':
            return f'-10<sup>{exp}</sup>' # Remove "1" from coefficient of -1
        return f'{coef} &times; 10<sup>{exp}</sup>'  # Format as "coefficient × 10<sup>exponent</sup>"
        
    return str(round(num, 2)).rstrip('0').rstrip('.')  # For number that doesn't need scientific formatting, round to 2 decimal places, remove trailing zeroes, and remove trailing decimal point


def will_gantry_collide(cyl: PyScriptObject, struct_set: PyScriptObject, r: float, iso: Dict[str, float], likelihood: str, couch_bounds: Optional[List[Dict[str, float]]] = None, ext_bounds: Optional[List[Dict[str, float]]] = None) -> Optional[str]:
    """Determines whether the couch and/or external geometries will likely collide with the gantry, using radius `r`

    A collision of a geometry with the gantry is likely if the geometry is more than `r` cm from the central z-axis through the iso, in any direction. 
    15 cm above and below (on the z-axis) the iso are checked

    Arguments
    ---------
    cyl: The ROI to use for the cylinder geometry
    struct_set: The structure set that the geometry belongs to
    r: The cylinder radius to use
    iso: The coordinates of the isocenter. The center of the cylinder geometry to create
    likelihood: String describing the "likelihood" if gantry collision. E.g., "likely", "very likely"
    couch_counds: The bounding box of the (outer) couch geometry in the structure set
                  Defaults to None
    ext_bounds: The bounding box of the external geometry in the structure set
                Defaults to None

    Returns
    -------
    A message dsecrbing the likelihood of collision with couch and/or external (patient)
    Returns None if collision is unlikely
    """
    # Create cylinder geometry
    cyl.CreateCylinderGeometry(Radius=r, Axis={'x': 0, 'y': 0, 'z': 1}, Length=30, Examination=struct_set.OnExamination, Center=iso)  # Cylinder w/ 30 cm height in inf-sup direction, centered at the iso
    cyl_bounds = struct_set.RoiGeometries[cyl.Name].GetBoundingBox()  # Cylinder bounding box
    
    # Do couch and/or external geometries extend outside the cylinder?
    if couch_bounds is not None:
        couch_collide = couch_bounds[0].x <= cyl_bounds[0].x or couch_bounds[0].y <= cyl_bounds[0].y or couch_bounds[1].x >= cyl_bounds[1].x or couch_bounds[1].y >= cyl_bounds[1].y
    else:
        couch_collide = False
    if ext_bounds is not None:
        ext_collide = ext_bounds[0].x <= cyl_bounds[0].x or ext_bounds[0].y <= cyl_bounds[0].y or ext_bounds[1].x >= cyl_bounds[1].x or ext_bounds[1].y >= cyl_bounds[1].y
    else:
        ext_collide = False

    # Returns the appropriate message
    if couch_collide or ext_collide:  # Collision is likely
        if couch_collide and ext_collide:  # Likely to collide with both couch and external
            return f'The couch and the patient are each >{r} cm away from the gantry. Collision may be {likelihood}.'
        if couch_collide:  # Likely to collide with couch but not external
            return f'The couch is >{r} cm away from the gantry. Collision may be {likelihood}.'
        return f'The patient is >{r} cm away from the gantry. Collision may be {likelihood}.'  # Likely to collide with external but not couch   


def plan_check() -> None:
    """Performs an "Initial Physics Review" plan check on the current plan

    SAVES PATIENT!

    Writes and opens a PDF report named "<patient name> <YYYY-MM-DD HH_MM_SS>.pdf" to the specified output directory
    Report is divided into several sections:
    - Errors (red): Things that really should be fixed
    - Warnings (yellow): Things that should be fixed but aren't just dire
    - Passing (green): Things that don't need to be fixed. This section is divided into subsections for Case, Plan, and each Beam Set.
    - Manual Checks (blue): Things the script can't check, so the user should check manually. These mainly relate to MOSAIQ.
    
    The following checks are performed:
    - Case
        * External ROI exists.
        * If external ROI exists: External ROI is named "External".
        * The ROI named "External", if it exists, is of type "External".
        * Case information is filled in: body site, diagnosis, physician name.
        * If physician name is present: Physician name includes "MD" suffix.
        * All exams have imaging system name "HOST-7307".
        * All exam names include a date (rough check - month, day, and year numbers are not validated).
        * For plans that are not VMAT H&N: Both couch ROIs exist.
        * Initial sim plan exists.
    - Plan
        * Planner is specified.
        * External is contoured on planning exam.
        * For prostate plans: The following ROIs exist: Bladder, Rectum, Colon_Sigmoid, Bag_Bowel.
        * There are no empty geometries on planning exam.
        * If external geometry exists and plan has dose: All contours are contained within external, within the bounds of the minimum of the dose grid and the planning exam.
        * SBRT plans with couch geometries: External and couch meet with no gap or overlap (tolerance of 3 mm).
        * If plan contains bolus: There is no gap between adjacent boli (overlap is okay).
        * Dose grid does not extend outside image.
        * All contours are contained within dose grid.
        * Dose grid voxel sizes are uniform.
        * Dose grid voxel sizes are <=2 mm for SBRT, <=3 mm otherwise.
        * Plan has dose.
        * Plan has Clinical Goals.
        * All Clinical Goals pass. (Ignore clinical goals for ROIs whose planning geometry is empty or has been updated since last voxel volume computation.)
        * Localization point is defined.
        * For SBRT plans with a localization point: Gantry and couch/patient are unlikely to collide.
    - Each beam set
        * Beam names are the same as their numbers.
        * Beam, including setup beam, names and numbers are unique across all cases, with the exception of initial sim plans.
        * Either a CT, or AP/PA and lat setup beam exist.
        * Machine is "SBRT 6MV" for SBRT, "ELEKTA" otherwise.
        * All beams have isocenter z-coordinate <=100 cm.
        * Beam set has Rx.
        * Beam set has dose.
        * For beam sets with Rx and dose:
            - For Rx to volume, all DSPs are within PTV.
            - For Rx to point, all DSPs are within 80% isodose line.
            - Each beam MU is at least 10% more than the beam Rx.
        * If initial sim plan exists and beam set has beam(s): No beam isocenters coordinates have been changed from initial sim.
        * Dose algorithm is Collapsed Cone (CCDose).
        * Autoscale to Rx is enabled.
        * For VMAT plans: All beam energies are 6 MV.
        * For VMAT plans: Optimization tolerance <=10^-5.
        * For VMAT lung plans: "Compute intermediate dose" is checked.
        * For VMAT plans: Each beam has gantry spacing <=3 cm.
        * For VMAT plans: Each beam has max delivery time <=120 s.
        * For VMAT plans: Each beam delivery time <=120 s for SBRT, 180 for SBRT (with a 10% margin).
        * For VMAT plans: Constraint on max leaf travel distance per degree is enabled.
        * For VMAT plans with this constraint enabled: Max leaf travel distance per degree <=0.5 cm.
        * If beam set has dose: Max dose (to external) <125 (or 140) for SBRT, <108 (or 110) for other VMAT, or <110 (or 115) for non-VMAT.
    - Manual Checks:
        * Admission date is filled in in MOSAIQ.
        * If beam set has Rx: Rx name in MOSAIQ matches Rx plan name in RS.
        * Any dose sums that the MD requested, are present in RS.
        * For VMAT plans: View MLC movie.
        * If current beam set Rx has non-100% isodose, this isodose is specified in MOSAIQ.
        * Are the proper structures excluded from MOSAIQ export and invisible?

    Assumptions
    -----------
    Current beam set is the 'main' beam set that will be exported to MOSAIQ.
    Each plan optimization optimizes a single beam set.
    Initial sim plan corresponding to this plan is the plan on the planning exam whose name contains "Initial Sim" or "Trial_1" (case insensitive).
    Initial sim plan, if it exists, has just one beam set, and all beams in that beam set share the same isocenter.

    It is best practice to run the plan check before anything is approved.
    Iteratively make changed according to Plan Check's errors/warnings and run Plan Check again.
    """
    # Get current variables
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()  # Exit script
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
        sys.exit()
    try:
        get_current('BeamSet')
    except:
        MessageBox.Show('There are no beam sets in the current plan. Click OK to abort the script.', 'No Beam Sets')
        sys.exit()
    struct_set = plan.GetTotalDoseStructureSet()  # Structure set on planning exam
    exam = struct_set.OnExamination  # Planning exam

    # Need to determine plan types now so that we know if ANY plan types are ____
    # Example: Minimum dose grid voxel size depends on whether ANY beam set is SBRT
    # OrderedDict to retain original order of beam sets, since this dict is used later to iterate over beam sets
    # Dict elements are beam set : plan type ('SRS', 'SBRT', 'VMAT', 'IMRT', or '3D')
    plan_types = OrderedDict()  
    for beam_set in plan.BeamSets:
        if beam_set.Modality == 'Photons':  # Ignore beam sets that are not photons
            fx = beam_set.FractionationPattern
            if fx is not None:
                fx = fx.NumberOfFractions
            if beam_set.PlanGenerationTechnique == 'Imrt':
                if beam_set.DeliveryTechnique == 'DynamicArc':
                    if fx in [1, 3]:
                        plan_types[beam_set] = 'SRS'
                    elif fx == 5:
                        plan_types[beam_set] = 'SBRT'
                    else:
                        plan_types[beam_set] = 'VMAT'
                else:
                    plan_types[beam_set] = 'IMRT'
            else:
                plan_types[beam_set] = '3D'
    
    # Are there any photon plans?
    if not plan_types:
        MessageBox.Show('This is not a photon plan. Click OK to abort script.', 'Incorrect Treatment Modality')
        sys.exit()

    # Is beam set supine or prone?
    if re.search(r'(Supine|Prone)$', beam_set.PatientPosition) is None:
        MessageBox.Show('Only supine or prone beam sets are supported. Click OK to abort the script.', 'Unsupported Patient Position')
        sys.exit()

    ## Prepare output file

    # Get couch names
    template_name = 'Elekta Couch' if 'Supine' in beam_set.PatientPosition else 'Elekta Prone Couch'
    template = get_current('PatientDB').LoadTemplatePatientModel(templateName=template_name)
    couch_names = outer_couch_name, inner_couch_name = [roi.Name for roi in template.PatientModel.RegionsOfInterest]  # [outer couch name, inner couch name]

    # Create PDF file
    if not os.path.isdir(PlanChkConstants.OUTPUT_DIR):
        os.makedirs(PlanChkConstants.OUTPUT_DIR)
    pt_name = format_name(patient.Name)
    filename = pt_name + ' ' + datetime.now().strftime('%Y-%m-%d %H_%M_%S') + '.pdf'
    filepath = os.path.join(PlanChkConstants.OUTPUT_DIR, re.sub(r'[<>:"/\\\|\?\*]', '_', filename))
    pdf = SimpleDocTemplate(filepath, pagesize=letter, bottomMargin=0.2 * inch, leftMargin=0.25 * inch, rightMargin=0.2 * inch, topMargin=0.2 * inch)  # 8.5 x 11" w/ 0.25" left & right margins, & 0.2" top & bottom margins

    # Headings
    hdg = Paragraph(pt_name, style=PlanChkConstants.STYLES['Heading1'])  # e.g., 'Jones, Bill'
    mrn = Paragraph('MRN: ' + patient.PatientID, style=PlanChkConstants.STYLES['Heading2'])  # e.g., 'MRN: 000123456'
    plan_chk = Paragraph('Plan Check: ' + plan.Name, style=PlanChkConstants.STYLES['Heading2'])  # e.g., 'Plan Check: Prostate'
    elems = [hdg, mrn, plan_chk]

    is_vmat_hn = set(plan_types.values()) == {'VMAT'} and case.BodySite == 'Head and Neck'  # Only VMAT H&N plans may lack couch

    rois = case.PatientModel.RegionsOfInterest
    roi_names = [roi.Name for roi in rois]
    
    # Lists of messages
    fail_msgs = []  # Errors
    warn_msgs = []  # Warnings
    pass_msgs = OrderedDict()  # No problems. Dict instead of list so that categories can be used ('Case', 'Plan', and a section for each beam set)
    # Manual checks come at the end of the report (after green messages)

    ## 'Patient:' section
    pass_msgs_section = []

    # Patient ID in correct format
    if re.match(PlanChkConstants.MRN_REGEX, patient.PatientID) is None:
        msg = 'Patient ID is in incorrect format: "' + patient.PatientID + '".'
        warn_msgs.append(msg)
    else:
        msg = 'Patient ID is in correct format: "' + patient.PatientID + '".'
        pass_msgs_section.append(msg)

    # Patient name in correct format
    if patient.Name.isupper() or re.match(PlanChkConstants.PT_NAME_REGEX, patient.Name) is None:
        msg = 'Patient name may be incorrectly formatted: "' + patient.Name + '".'
        warn_msgs.append(msg)
    else:
        msg = 'Patient name appears to be correctly formatted: "' + patient.Name + '".'
        pass_msgs_section.append(msg)

    # Patient DOB not in future or very distant past
    dob = patient.DateOfBirth
    dob = datetime(dob.Year, dob.Month, dob.Day)
    today = datetime.today()
    if dob > today:
        msg = 'Patient birthday, ' + dob.strftime('%m/%d/%Y'), ', is in the future.'
        fail_msgs.append(msg)
    elif (today - dob).days / 365.25 > 125:
        msg = 'Patient birthday, ' + dob.strftime('%m/%d/%Y') + ', appears incorrect.'
        fail_msgs.append(msg)
    else:
        msg = 'Patient birthday, ' + dob.strftime('%m/%d/%Y') + ', is not obviously incorrect.'
        pass_msgs_section.append(msg)

    ## 'Case:' section
    pass_msgs_section = []

    # External ROI exists
    try:
        ext = next(roi for roi in rois if roi.Type == 'External')  # There will never be more than one external ROI
    except StopIteration:
        ext = None
        msg = 'There is no external ROI.'
        fail_msgs.append(msg)
    else:
        msg = 'External ROI exists.'
        pass_msgs_section.append(msg)

        # External is named "External"
        if ext.Name != 'External':
            msg = 'External ROI is named "' + ext.Name + '", not "External".'
            fail_msgs.append(msg)
        else:
            msg = 'External ROI is named "External".'
            pass_msgs_section.append(msg)

    # ROI named "External" is of type "External"
    try:
        external_type = next(roi.Type for roi in rois if roi.Name == 'External' and roi.Type != 'External')
    except StopIteration:
        msg = 'The ROI named "External" is of type "External".'
        pass_msgs_section.append(msg)
    else:
        msg = 'The ROI named "External" is of type "' + external_type + '", not "External".'
        fail_msgs.append(msg)

    # Case information is filled in
    case_attrs = {'Body site': case.BodySite, 'Diagnosis': case.Diagnosis, 'Physician name': case.Physician.Name}
    attrs = [attr for attr, case_attr in sorted(case_attrs.items()) if case_attr == '']
    if attrs:
        msg = 'Case information is missing:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(attrs)
        warn_msgs.append(msg)
    else:
        case_attrs = [attr + ': ' + case_attr for attr, case_attr in sorted(case_attrs.items())]  # e.g., 'Body site: Thorax'
        msg = 'All case information is filled in:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(case_attrs)
        pass_msgs_section.append(msg)

    # MD name includes "MD" suffix
    if case.Physician.Name is not None:
        print(case.Physician.Name)
        if re.search(r'\^MD($|\^)', case.Physician.Name) is None:
            msg = 'Physician name is missing "MD" suffix: "' + case.Physician.Name + '".'
            fail_msgs.append(msg)
        else:
            msg = 'Physician name includes "MD" suffix: "' + case.Physician.Name + '".'
            pass_msgs_section.append(msg)

    # Exams: Imaging system name is HOST-7307
    wrong_img_sys = [e.Name for e in case.Examinations if e.EquipmentInfo.ImagingSystemReference is None or e.EquipmentInfo.ImagingSystemReference.ImagingSystemName != 'HOST-7307']
    if wrong_img_sys:
        msg = 'The imaging system is incorrect for the following exams:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}'.format('<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;'.join(wrong_img_sys))
        fail_msgs.append(msg)
    else:
        msg = 'All exams have imaging system HOST-7307.'
        pass_msgs_section.append(msg) 

    # Exam names include date
    missing_date = [e.Name for e in case.Examinations if re.search(fr'({PlanChkConstants.DATE_REGEX})|(^{PlanChkConstants.IMG_FOR_TEMPLATES_REGEX}$)', e.Name, re.IGNORECASE) is None]  # Very crude date regex: does not validate month, day, or year numbers. Also, ignore "IMAGE FOR TEMPLATES" exams
    if missing_date:
        msg = 'The following exam names are missing a date:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(missing_date)
        warn_msgs.append(msg)
    else:
        msg = 'All exam names include a date.'
        pass_msgs_section.append(msg)

    # Couch ROIs exist
    missing_couch_rois = [couch_name for couch_name in couch_names if couch_name not in roi_names]
    if not is_vmat_hn:
        if missing_couch_rois:
            msg = 'Couch ROI(s) are missing:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(missing_couch_rois)
            if set(plan_types.values()) == {'3D'}:  # All plans are 3D
                warn_msgs.append(msg)
            else:  # There are IMRT plans
                fail_msgs.append(msg)
        else:
            msg = 'Couch ROIs exist:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(couch_names)
            pass_msgs_section.append(msg)

    # "Initial sim" plan exists
    try:
        ini_sim_plan = next(p for p in case.TreatmentPlans if re.match(PlanChkConstants.INI_SIM_PLAN_REGEX, p.Name, re.IGNORECASE) is not None and ('SBRT' in plan_types.values() or p.GetTotalDoseStructureSet().OnExamination.Equals(exam)))  # "Initial sim" plan must be on planning exam unless this is an SBRT plan
    except StopIteration:
        ini_sim_plan = None
        msg = 'There is no "Initial Sim" plan.'
        warn_msgs.append(msg)
    else:
        msg = '"Initial Sim" plan is present.'
        pass_msgs_section.append(msg)

    # If there are any green case messages, add them to `pass_msgs` dict
    if pass_msgs_section:
        pass_msgs['Case:'] = pass_msgs_section

    ## 'Plan:' section
    pass_msgs_section = []

    dose_dist = plan.TreatmentCourse.TotalDose

    # Plan information is filled in
    if plan.PlannedBy == '':
        msg = 'Planner is not specified.'
        warn_msgs.append(msg)
    else:
        msg = 'Planner is specified: ' + plan.PlannedBy
        pass_msgs_section.append(msg)

    if ext is not None:
        # External has geometry on planning exam
        if not struct_set.RoiGeometries[ext.Name].HasContours():
            msg = 'There is no external geometry on the planning exam.'
            fail_msgs.append(msg)
        else:
            msg = 'External is contoured on planning exam.'
            pass_msgs_section.append(msg)

    # For prostate plans, ensure certain ROIs exist
    # We know it's a prostate plan if any of certain prostate-related keywords is in certain case/plan/beam set info fields
    chk_for_body_site = [case.BodySite, case.CaseName, case.Comments, case.Diagnosis, plan.Comments, plan.Name] + [beam_set.DicomPlanLabel for beam_set in plan.BeamSets]  # Fields to check for prostate keywords
    if any(re.search(PlanChkConstants.PROS_PLAN_REGEX, attr) is not None for attr in chk_for_body_site):  # It's a prostate plan
        pros_rois = ['Bladder', 'Rectum', 'Colon_Sigmoid', 'Bag_Bowel']  # ROIs that must be present if this is a prostate plan
        missing_pros_rois = [pros_roi for pros_roi in pros_rois if pros_roi not in roi_names]
        if missing_pros_rois:
            msg = 'Important prostate plan ROI(s) are missing:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(missing_pros_rois)
            fail_msgs.append(msg)
        else:
            msg = 'Important prostate plan ROIs exist:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(pros_rois)
            pass_msgs_section.append(msg)

    # Empty geometries on planning exam
    empty_geom_names = [geom.OfRoi.Name for geom in struct_set.RoiGeometries if not geom.HasContours() and not geom.OfRoi.Equals(ext)]  # No external should be error (taken care of above), not warning
    if empty_geom_names:
        msg = 'The following ROIs are empty on the planning exam:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(empty_geom_names)
        warn_msgs.append(msg)
    else:
        msg = 'There are no empty geometries on the planning exam.'
        pass_msgs_section.append(msg)

    # ROIs that have been updated since last voxel volume computation: have contours but no volume in dose grid
    dose_stats_missing = [geom.OfRoi.Name for geom in struct_set.RoiGeometries if geom.OfRoi.Name not in couch_names and geom.HasContours() and dose_dist.GetDoseGridRoi(RoiName=geom.OfRoi.Name).RoiVolumeDistribution is None]
    if dose_stats_missing:
        msg = 'Dose statistics need updating.'
        warn_msgs.append(msg)
    else:
        msg = 'All dose statistics are up to date.'
        pass_msgs_section.append(msg)
    
    dg = plan.GetTotalDoseGrid()
    dg_min = dg.Corner
    dg_max = {dim: coord + dg.NrVoxels[dim] * dg.VoxelSize[dim] for dim, coord in dg_min.items()}

    img_stack = exam.Series[0].ImageStack
    exam_min, exam_max = img_stack.GetBoundingBox()
    exam_px_sz = img_stack.PixelSize
    
    # Dose grid contained inside planning exam, ideally with some "padding" between them
    if all(couch_name in missing_couch_rois or not struct_set.RoiGeometries[couch_name].HasContours() for couch_name in couch_names):
        if any(dg_min[dim] < coord for dim, coord in exam_min.items()) or any(dg_max[dim] > coord for dim, coord in exam_max.items()):
            msg = 'Dose grid extends outside planning exam.'
            warn_msgs.append(msg)
        elif any(dg_min[dim] - exam_min[dim] < sz for dim, sz in exam_px_sz.items()) or dg_min['z'] - exam_min['z'] < dg.VoxelSize.z or any(exam_max[dim] - dg_max[dim] < sz for dim, sz in exam_px_sz.items()) or exam_min['z'] - dg_min['z'] < dg.VoxelSize.z:
            msg = 'There should be at least 1 px between each dose grid and planning exam edge.'
            warn_msgs.append(msg)
        else:
            msg = 'There is at least 1 px padding between each dose grid and planning exam edge.'
            pass_msgs_section.append(msg)

    #There are no contours "in the air" (outside external)
    # A contour extends outside external if any of its min coords are less than external min coordinates, or any of its max coordinates are greater than external max coordinates
    # Ignore coordinates outside the planning exam or dose grid, whichever is stricter
    if ext is not None and dose_dist.DoseValues is not None and len(set(dg.VoxelSize.values())) == 1:  # External exists, dose grid is defined, and dose grid voxel sizes are uniform
        ext_geom = struct_set.RoiGeometries[ext.Name]
        if ext_geom.HasContours():
            vox_sz = dg.VoxelSize.x 

            # Min and max coordinates in dose grid, defined by a box geometry
            box_name = case.PatientModel.GetUniqueRoiName(DesiredName='zDoseGrid')
            box = case.PatientModel.CreateRoi(Name=box_name, Type='Control')
            box_min = {dim: max(coord, exam_min[dim]) for dim, coord in dg_min.items()}
            box_max = {dim: min(coord, exam_max[dim]) for dim, coord in dg_max.items()}
            box_ctr = {dim: (coord + box_max[dim]) / 2.0 for dim, coord in box_min.items()}
            box_sz = {dim: box_max[dim] - coord for dim, coord in box_min.items()}
            box.CreateBoxGeometry(Size=box_sz, Examination=exam, Center=box_ctr, VoxelSize=vox_sz)

            # Voxel indices of planning exam, and external w/ 3 mm margin
            ext_prv_name = case.PatientModel.GetUniqueRoiName(DesiredName=f'External_PRV{str(int(vox_sz * 10)).zfill(2)}')
            ext_prv = case.PatientModel.CreateRoi(Name=ext_prv_name, Type='Control')
            ext_prv.SetMarginExpression(SourceRoiName=ext.Name, MarginSettings={ 'Type': 'Expand', 'Superior': vox_sz, 'Inferior': vox_sz, 'Anterior': vox_sz, 'Posterior': vox_sz, 'Right': vox_sz, 'Left': vox_sz })
            ext_prv.UpdateDerivedGeometry(Examination=exam)
            dose_dist.UpdateDoseGridStructures()
            box_vi = set(dose_dist.GetDoseGridRoi(RoiName=box_name).RoiVolumeDistribution.VoxelIndices)  # Voxel indices of box geometry ("voxel indices" of image)
            ext_vi = set(dose_dist.GetDoseGridRoi(RoiName=ext_prv_name).RoiVolumeDistribution.VoxelIndices).intersection(box_vi)  # Voxel indices of external ROI that are inside the image / dose grid

            # Delete unnecessary ROIs
            box.DeleteRoi()  # Box ROI no longer needed
            ext_prv.DeleteRoi()

            stray_contours = []  # Geometries that extend outside external
            for geom in struct_set.RoiGeometries:
                if geom.HasContours() and geom.OfRoi.Type not in ['Bolus', 'Control', 'External', 'FieldOfView', 'Fixation', 'Support'] and geom.OfRoi.RoiMaterial is None:  # Ignore the external contour, any other external contours, FOV or support (e.g., couch) contours, and ROIs with a material defined
                    geom_vi = set(dose_dist.GetDoseGridRoi(RoiName=geom.OfRoi.Name).RoiVolumeDistribution.VoxelIndices).intersection(box_vi)  # Voxel indices of the geometry that are inside the image
                    if not geom_vi.issubset(ext_vi):
                        stray_contours.append(geom.OfRoi.Name)
            
            if stray_contours:
                msg = 'The following contours extend outside the external:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(stray_contours)
                fail_msgs.append(msg)
            else:
                msg = 'All contours are contained inside the external.'
                pass_msgs_section.append(msg)

        # External extends to couch w/o gap or overlap
        if 'SBRT' in plan_types.values() and not missing_couch_rois and struct_set.RoiGeometries[outer_couch_name].HasContours() and struct_set.RoiGeometries[inner_couch_name].HasContours():
            ext_bottom = struct_set.RoiGeometries[ext.Name].GetBoundingBox()[1].y  # Bottom of external
            couch_top = struct_set.RoiGeometries[outer_couch_name].GetBoundingBox()[0].y
            diff = round(ext_bottom - couch_top, 2)
            if diff < -0.3:
                msg = f'External and couch overlap by {-diff} cm.'
                fail_msgs.append(msg)
            elif diff > 0.3:
                msg = f'There is a {diff}-cm gap between external and couch.'
                fail_msgs.append(msg)
            else:
                msg = 'There is no overlap or gap between external and couch.'
                pass_msgs_section.append(msg)

    # No gap between adjacent boli
    boli = [geom for geom in struct_set.RoiGeometries if geom.HasContours() and geom.OfRoi.Type == 'Bolus' and geom.OfRoi.DerivedRoiExpression is None]  # Non-derived bolus geometries
    if len(boli) > 1:  # Plan has bolus
        ext_coords = get_contour_coords(struct_set.RoiGeometries[ext.Name])  # All coordinates in External geometry
        gap_btwn_boli = []  # List of tuples of adjacent boli w/ a gap; e.g., [('Bolus 1', 'Bolus 2'), ('Bolus 3', 'Bolus 4')]
        
        # Iterate over each bolus, finding the adjacent (closest) bolus, as determined by smallest distance between closest two points that are also part of External
        for bolus in boli:
            bolus_coords = [coords for coords in get_contour_coords(bolus) if coords in ext_coords]  # All coordinates shared by bolus and External geometries

            # All other boli, from which to find adjacent bolus
            other_boli = boli[:]
            other_boli.remove(bolus)  # Bolus can't be adjacent to itself

            # Find adjacent bolus
            min_dist = float('Inf')  # Start min at largest possible so we'll be sure to find an adjacent bolus
            adj_bolus = None  # We don't yet have an adjacent bolus
            gap = False  # Assume no gap between bolus and its adjacent bolus
            for other_bolus in other_boli:
                other_bolus_coords = [coords for coords in get_contour_coords(bolus) if coords in ext_coords]  # All coordinates shared by possible adjacent bolus, and External geometries
                # Compare each pair of coordinates from the two boli
                for bolus_coord in bolus_coords:
                    for other_bolus_coord in other_bolus_coords:
                        dist = distance(bolus_coord, other_bolus_coord)
                        if dist < min_dist:  # We found another bolus that it's closer to
                            min_dist = dist
                            adj_bolus = other_bolus
                            gap = bolus_coord not in other_bolus_coords  # There is a gap if there is no overlap; IOW, the two boli don't share the coordinate
            if gap and (adj_bolus, bolus) not in gap_btwn_boli:  # Prevent duplicates in gap list (e.g., ('Bolus 1', 'Bolus 2') and ('Bolus 2', 'Bolus 1'))
                gap_btwn_boli.append((bolus, adj_bolus))
        
        if gap_btwn_boli:
            gap_btwn_boli = [f'{pair[0]} and {pair[1]}' for pair in gap_btwn_boli]  # e.g., 'Bolus 1 and Bolus 2'
            msg = 'There is a gap between each of the follwing pair(s) of adjacent boli:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}'.format('<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;'.join(gap_btwn_boli))
            fail_msgs.append(msg)
        else:
            msg = 'There are no gaps between adjacent boli.'
            pass_msgs_section.append(msg)

    # Dose grid includes all contours (except perhaps FOV)
    # A contour extends outside dose grid if any of its min coords are less than dose grid min coordinates, or any of its max coordinates are greater than dose grid max coordinates
    outside_dg = []  # Geometries that extend outside dose grid
    for geom in struct_set.RoiGeometries:  # Ignore empty geometries
        if geom.HasContours() and geom.OfRoi.Type != 'FieldOfView' and not (geom.OfRoi.Type in ['Bolus', 'Fixation', 'Support'] and geom.OfRoi.RoiMaterial is not None):  # Ignore FOV, and bolus/fixation/support with material override
            bounds = geom.GetBoundingBox()  # [{'x': min x-coord, 'y': min y-coord, 'z': min z-coord}, {'x': max x-coord, 'y': max y-coord, 'z': max z-coord}]
            if any(bounds[0][coord] < val for coord, val in dg_min.items()) or any(bounds[1][coord] > val for coord, val in dg_max.items()):
                outside_dg.append(geom.OfRoi.Name)
    if outside_dg:
        msg = 'Dose grid does not include all of the following geometries:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(outside_dg)
        warn_msgs.append(msg)
    else:
        msg = 'All necessary contours are contained inside the dose grid.'
        pass_msgs_section.append(msg)

    # Uniform dose grid
    voxel_szs = ['{} = {:.0f} mm'.format(coord, sz * 10) for coord, sz in sorted(dg.VoxelSize.items())]
    if not dg.VoxelSize.x == dg.VoxelSize.y == dg.VoxelSize.z:
        msg = 'Dose grid voxel sizes are not uniform:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(voxel_szs)
        fail_msgs.append(msg)
    else:
        msg = 'Dose grid voxel sizes are uniform: x = y = z = {:.0f} mm.'.format(dg.VoxelSize.x * 10)
        pass_msgs_section.append(msg)

    # Dose grid voxel sizes are small enough
    max_sz = 2 if 'SBRT' in plan_types.values() or 'SRS' in plan_types.values() else 3  # 3 mm dose grid for non-SBRT, 2 mm for SBRT (incl. SRS)
    lg_voxels = ['{} = {:.0f} mm'.format(coord, sz * 10) for coord, sz in dg.VoxelSize.items() if sz > max_sz]  # Coordinates whose voxel sizes are too large. Convert from cm to mm and display as integer, not float
    if lg_voxels:
        msg = f'The following voxel sizes >{max_sz} mm:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(lg_voxels)
        fail_msgs.append(msg)
    else:
        voxel_szs = ['{} = {:.0f} mm'.format(coord, sz * 10) for coord, sz in sorted(dg.VoxelSize.items())]
        msg = f'All voxel sizes &leq;{max_sz} mm:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(voxel_szs)
        pass_msgs_section.append(msg)

    # Plan has Clinical Goals
    if plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count == 0:
        msg = 'Plan has no Clinical Goals.'
        warn_msgs.append(msg)
    else:
        msg = 'Plan has Clinical Goals.'
        pass_msgs_section.append(msg)

    # Lists of clinical goals that fail due to empty geometries, geometries updated since last voxel volume computation (need to run UpdateDoseGridStructures or click 'Dose statistics missing' in GUI)
    # Display format uses similar logic to what RS presumably uses to display Clinical Goal text
    # E.g., 'External: At most 6250 cGy dose at 0.03 cm^3 volume (6652 cGy)'
    if dose_dist.DoseValues is None:  # Plan has no dose
        msg = 'Plan has no dose.'
        fail_msgs.append(msg)
    else:
        failing_goals = []
        for func in plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions:
            roi = func.ForRegionOfInterest
            ok = roi.Name not in dose_stats_missing and not func.EvaluateClinicalGoal()
            try:
                ok = ok and struct_set.RoiGeometries[roi.Name].HasContours()
            except KeyError:
                ok = ok and struct_set.PoiGeometries[roi.Name].Point is not None and struct_set.PoiGeometries[roi.Name].Point.x != float('inf')
            if ok:
                goal_val = func.GetClinicalGoalValue()  # nan if empty or out-of-date geometry
                failing_goals.append(goal_str(func))

        if failing_goals:
            msg = 'The following Clinical Goals fail:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(failing_goals)
            fail_msgs.append(msg)
        else:
            msg = 'All evaluable Clinical Goals pass.'
            pass_msgs_section.append(msg)

    # Gantry does not collide with couch or patient
    # Max distance between plan isocenter and couch/Skin should be <40 cm, at worst <41.5 cm
    # Create cylinder geometry w/ radius 40 and 41.5, and height the diameter of the collimator
    # Min and max couch/skin coordinates must not be outside cylinder min and max coordinates

    # Plan isocenter coordinates
    iso = struct_set.LocalizationPoiGeometry
    if iso is None or iso.Point is None or any(abs(coord) == float('inf') for coord in iso.Point.values()):
        msg = 'Plan has no localization geometry.'
        fail_msgs.append(msg)
    else:
        iso = iso.Point
        msg = 'Localization geometry is defined.'    
        pass_msgs_section.append(msg)

        # Bounds of couch and external
        if not missing_couch_rois and struct_set.RoiGeometries[outer_couch_name].HasContours():
            couch_bounds = struct_set.RoiGeometries[outer_couch_name].GetBoundingBox()
        else:
            couch_bounds = None
        if ext and struct_set.RoiGeometries[ext.Name].HasContours():
            ext_bounds = struct_set.RoiGeometries[ext.Name].GetBoundingBox()
        else:
            ext_bounds = None

        # Create cylinder for collision checking
        cyl_name = case.PatientModel.GetUniqueRoiName(DesiredName='zCylinder')
        cyl = case.PatientModel.CreateRoi(Name=cyl_name, Type='Control')
        
        msg = will_gantry_collide(cyl, struct_set, 40, iso, 'very likely', couch_bounds, ext_bounds,)  # 40 cm cylinder
        if msg is not None:
            fail_msgs.append(msg)
        else:
            msg = will_gantry_collide(cyl, struct_set, 41.5, iso, 'likely', couch_bounds, ext_bounds,)  # 41.5 cm cylinder
            if msg is not None:
                warn_msgs.append(msg)
            else:
               pass_msgs_section.append('Distance from gantry to patient/couch &GreaterEqual;41.5 cm. Collision is unlikely.') 
        
        cyl.DeleteRoi()

    # If there are any green case messages, add them to `pass_msgs` dict
    if pass_msgs_section:
        pass_msgs['Plan:'] = pass_msgs_section

    ## Beam set sections

    # List of all beam names (including setup beams), for preventing duplicate beam names
    beam_names, beam_nums = [], []
    for c in patient.Cases:
        for p in c.TreatmentPlans:
            if re.match(PlanChkConstants.INI_SIM_PLAN_REGEX, p.Name, re.IGNORECASE) is not None:
                continue
            for bs in p.BeamSets:
                beam_names.extend(b.Name for b in bs.Beams)
                beam_nums.extend(b.Number for b in bs.Beams)
                beam_names.extend(sb.Name for sb in bs.PatientSetup.SetupBeams)
                beam_nums.extend(sb.Number for sb in bs.PatientSetup.SetupBeams)

    # Iterate over beam sets in `plan_types` dict instead of plan.BeamSets b/c we only want photon beam sets
    for i, beam_set in enumerate(plan_types):
        pass_msgs_section = []

        beam_set_name = beam_set.DicomPlanLabel
        plan_type = plan_types[beam_set]
        try:
            bs_rx = beam_set.Prescription.PrimaryPrescriptionDoseReference
        except:
            bs_rx = None

        # Beam name = beam number
        bad_names = [f'{beam.Name} (#{beam.Number})' for beam in beam_set.Beams if beam.Name != str(beam.Number)]  # e.g., 'CCW (#2)'
        if bad_names:
            msg = f'The following beam names in beam set "{beam_set_name}" are not the same as their numbers:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(bad_names)
            warn_msgs.append(msg)
        else:
            msg = 'Beam names are the same as their numbers.'
            pass_msgs_section.append(msg)

        # Duplicate beam names
        dup_names = [beam.Name for beam in beam_set.Beams if beam_names.count(beam.Name) > 1]
        if dup_names:
            msg = f'The following beam names in beam set "{beam_set_name}" exist in other cases or plans or as setup beam names in the current plan:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(dup_names)
            fail_msgs.append(msg)
        else:
            msg = 'No beam name already exists.'
            pass_msgs_section.append(msg)

        # Duplicate setup beam names
        sbs = beam_set.PatientSetup.SetupBeams
        dup_names = [sb.Name for sb in sbs if beam_names.count(sb.Name) > 1]
        if dup_names:
            msg = f'The following setup beam names in beam set "{beam_set_name}" exist in other cases or plans or as beam names in the current plan:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(dup_names)
            fail_msgs.append(msg)
        else:
            msg = 'No setup beam name already exists.'
            pass_msgs_section.append(msg)

        # Duplicate beam numbers
        dup_nums = [str(beam.Number) for beam in beam_set.Beams if beam_nums.count(beam.Number) > 1]
        if dup_nums:
            msg = f'The following beam numbers in beam set "{beam_set_name}" exist in other cases or plans or as setup beam numbers in the current plan:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(dup_nums)
            fail_msgs.append(msg)
        else:
            msg = 'No beam number already exists.'
            pass_msgs_section.append(msg)

        # Duplicate setup beam numbers
        dup_nums = [str(sb.Number) for sb in sbs if beam_nums.count(sb.Number) > 1]
        if dup_nums:
            msg = f'The following setup beam names in beam set "{beam_set_name}" exist in other cases or plans or as beam numbers in the current plan:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(dup_nums)
            fail_msgs.append(msg)
        else:
            msg = 'No setup beam number already exists.'
            pass_msgs_section.append(msg)

        # Setup beams present: either AP + Lat kV, or CBCT (preferred)
        gantry_angles = [sb.GantryAngle for sb in sbs]
        if sbs.Count == 0:
            msg = f'There are no setup beams for beam set "{beam_set_name}". There should be either a CBCT, or an AP/PA and a lat.'
            warn_msgs.append(msg)
        elif any(re.search(PlanChkConstants.CBCT_REGEX, sb.Name, re.IGNORECASE) is not None or re.search(PlanChkConstants.CBCT_REGEX, sb.Description, re.IGNORECASE) is not None for sb in sbs):
            msg = 'CBCT setup beam exists.'
            pass_msgs_section.append(msg)
        elif (0 in gantry_angles or 180 in gantry_angles) and (90 in gantry_angles or 270 in gantry_angles):
            msg = 'AP/PA and lat setup beams exist.'
            pass_msgs_section.append(msg)
        else:
            msg = f'Beam set "{beam_set_name}" contains neither a CBCT setup beam, nor both an AP/PA and a lat setup beam.'
            warn_msgs.append(msg) 

        # Machine is ELEKTA or SBRT 6MV
        machine = 'SBRT 6MV' if plan_type in ['SRS', 'SBRT'] else 'ELEKTA'
        if beam_set.MachineReference.MachineName != machine:
            msg = f'Machine for beam set "{beam_set_name}" should be "{machine}", not "{beam_set.MachineReference.MachineName}".'
            fail_msgs.append(msg)
        else:
            msg = f'Machine is "{machine}".'
            pass_msgs_section.append(msg)
           
        # z-coordinate of beam isos between -100 and 100
        lg_z = [f'{beam.Name} ({format_num(beam.Isocenter.Position.z)} cm)' for beam in beam_set.Beams if abs(beam.Isocenter.Position.z) > 100]  # e.g., '2 (105 cm)'
        if lg_z:
            msg = f'The following beams in beam set "{beam_set_name}" have isocenter z-coordinate > 100 cm:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(lg_z)
            fail_msgs.append(msg)
        else:
            z = [f'{beam.Name} ({format_num(beam.Isocenter.Position.z)} cm)' for beam in beam_set.Beams]
            msg = 'All beams have isocenter z-coordinate &leq; 100 cm:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(z)
            pass_msgs_section.append(msg)        

        # Beam set has Rx
        if bs_rx is None:
            msg = f'There is no prescription for beam set "{beam_set_name}".'
            fail_msgs.append(msg)
        else:
            msg = 'Beam set has a prescription.'
            pass_msgs_section.append(msg)

        # Beam set has dose
        if beam_set.FractionDose.DoseValues is None:
            msg = f'Beam set "{beam_set_name}" has no dose.'
            fail_msgs.append(msg)
        else:
            msg = 'Beam set has dose.'
            pass_msgs_section.append(msg)

            # DSPs are near target
            # For Rx to volume, DSP is within PTV
            # For Rx to point, DSP is within 80% isodose line (create a geometry from dose to determine this)
            if bs_rx is not None:  # There is an Rx, and it is to volume or to point
                if bs_rx.PrescriptionType == 'DoseAtVolume':  # Rx to volume
                    roi = bs_rx.OnStructure  # PTV
                    dsp_roi = 'PTV'
                else:  # Rx to point
                    roi_name = case.PatientModel.GetUniqueRoiName(DesiredName='IDL_80%')
                    roi = case.PatientModel.CreateRoi(Name=roi_name, Type='Control')  # Create ROI
                    roi.CreateRoiGeometryFromDose(DoseDistribution=dose_dist, ThresholdLevel=0.8 * bs_rx.DoseValue)  # Set geometry to isodose line for 80% of the Rx
                    dsp_roi = '80% isodose line'

                roi_bounds = struct_set.RoiGeometries[roi.Name].GetBoundingBox()
                bad_dsps = []
                for dsp in beam_set.DoseSpecificationPoints:
                    if any(dsp.Coordinates[coord] < val for coord, val in roi_bounds[0].items()) or any(dsp.Coordinates[coord] > val for coord, val in roi_bounds[1].items()):
                        bad_dsps.append(f'{dsp.Name}: {format_coords(dsp.Coordinates)}')

                # Delete IDL ROI
                if dsp_roi == '80% isodose line':  # Delete the IDL ROI if it exists
                    roi.DeleteRoi() 

                if bad_dsps:
                    msg = f'The following DSPs for beam set "{beam_set_name}" are outside the {dsp_roi}:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(bad_dsps)
                    fail_msgs.append(msg)
                else:
                    msg = f'All DSPs are inside the {dsp_roi}.'
                    pass_msgs_section.append(msg)         

        # Iso has not been changed from initial sim
        if ini_sim_plan is not None and ini_sim_plan.BeamSets[0].Beams.Count > 0:  # Assume a single beam set in initial sim plan
            ini_sim_iso = ini_sim_plan.BeamSets[0].Beams[0].Isocenter  # Assume all beams in ini sim beam set have same iso
            loc_pt = beam_set.GetStructureSet().LocalizationPoiGeometry
            if loc_pt is None or loc_pt.Point is None or any(abs(coord) == float('inf') for coord in loc_pt.Point.values()):
                iso_chged = [f'{b.Name}: {b.Isocenter.Annotation.Name} {format_coords(b.Isocenter.Position)}'for b in beam_set.Beams if format_coords(b.Isocenter.Position) != format_coords(ini_sim_iso.Position)]
                if iso_chged:
                    msg = f'Isocenter coordinates for the following beams in beam set "{beam_set_name}" were changed from initial sim:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(iso_chged)
                    warn_msgs.append(msg)
                else:
                    msg = f'No isocenter coordinates were changed from initial sim {format_coords(ini_sim_iso.Position)}.'
                    pass_msgs_section.append(msg)
            else:
                loc_pt = loc_pt.Point
                if format_coords(loc_pt) != format_coords(ini_sim_iso.Position):
                    msg = f'Loc point coordinates for beam set "{beam_set_name}" {format_coords(loc_pt)} were changed from initial sim iso coordinates {format_coords(ini_sim_iso.Position)}'
                    warn_msgs.append(msg)
                else:
                    msg = f'Loc point coordinates and initial sim iso coordinates {format_coords(ini_sim_iso.Position)} are the same.'
                    pass_msgs_section.append(msg)
          
        # Dose algorithm is Collapsed Cone (CCDose)
        if beam_set.AccurateDoseAlgorithm.DoseAlgorithm != 'CCDose':
            msg = f'Dose algorithm for beam set "{beam_set_name}" is {beam_set.AccurateDoseAlgorithm.DoseAlgorithm}. It should be CCDose.'
            fail_msgs.append(msg)
        else:
            msg = f'Dose algorithm is CCDose.'
            pass_msgs_section.append(msg)

        # Plan optimization settings (there is a different plan optimization object for each beam set)
        opt = plan.PlanOptimizations[i]
        if not opt.AutoScaleToPrescription:
            msg = f'Autoscale to prescription is disabled for beam set "{beam_set_name}".'
            fail_msgs.append(msg)
        else:
            msg = 'Autoscale to prescription is enabled.'
            pass_msgs_section.append(msg)

        # The following checks are for VMAT (incl. SRS, SBRT) only
        if plan_type in ['VMAT', 'SRS', 'SBRT']:
            # Beam energy = 6 MV
            bad_energy = [f'{beam.Name} ({beam.BeamQualityId} MV)' for beam in beam_set.Beams if beam.BeamQualityId != '6']  # e.g., 'CCW (18 MV)'
            if bad_energy:
                msg = f'The following beam energies in beam set "{beam_set_name}" should be 6 MV:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(bad_energy)
                fail_msgs.append(msg)
            else:
                msg = 'All beam energies are 6 MV.'
                pass_msgs_section.append(msg)    

            # Optimization tolerance <=10^-5
            if opt.OptimizationParameters.Algorithm.OptimalityTolerance > 0.0001:
                msg = f'Optimization tolerance for beam set "{beam_set_name}" = {format_num(opt.OptimizationParameters.Algorithm.OptimalityTolerance)} > 10<sup>-5</sup>.'
                fail_msgs.append(msg)
            else:
                msg = f'Optimization tolerance = {format_num(opt.OptimizationParameters.Algorithm.OptimalityTolerance)} &leq; 10<sup>-5</sup>.'
                pass_msgs_section.append(msg)

            # Compute Intermediate Dose is checked
            if any(re.search(r'(^|[^A-Za-z])lung([^A-Za-z]|$)', attr, re.IGNORECASE) is not None for attr in [beam_set_name, plan.Name, case.CaseName, case.BodySite]):
                if not opt.OptimizationParameters.DoseCalculation.ComputeIntermediateDose:
                    msg = f'"Compute intermediate dose" is unchecked for beam set "{beam_set_name}".'
                    warn_msgs.append(msg)
                else:
                    msg = '"Compute intermediate dose" is checked.'
                    pass_msgs_section.append(msg)

            # Gantry spacing <=3 cm
            tss = opt.OptimizationParameters.TreatmentSetupSettings[0]
            bad_gantry_spacing = [f'{beam.Name} ({format_num(tss.BeamSettings[j].ArcConversionPropertiesPerBeam.FinalArcGantrySpacing)}&deg;)' for j, beam in enumerate(beam_set.Beams) if tss.BeamSettings[j].ArcConversionPropertiesPerBeam.FinalArcGantrySpacing > 3]
            if bad_gantry_spacing:
                msg = f'The following beams in beam set "{beam_set_name}" have gantry spacing >3&deg;:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(bad_gantry_spacing)
                fail_msgs.append(msg)
            else:
                gantry_spacing = [f'{beam.Name} ({format_num(tss.BeamSettings[j].ArcConversionPropertiesPerBeam.FinalArcGantrySpacing)}&deg;)' for j, beam in enumerate(beam_set.Beams)]
                msg = 'All beams have gantry spacing &leq; 3&deg;:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(gantry_spacing)
                pass_msgs_section.append(msg)

            # Max delivery time <=120 or 180 s
            max_del_time = 180 if plan_type in ['SRS', 'SBRT'] else 120
            bad_max_del = [f'{beam.Name} ({format_num(tss.BeamSettings[j].ArcConversionPropertiesPerBeam.MaxArcDeliveryTime)} s)' for j, beam in enumerate(beam_set.Beams) if tss.BeamSettings[j].ArcConversionPropertiesPerBeam.MaxArcDeliveryTime > max_del_time]
            if bad_max_del:
                msg = f'The following beams in beam set "{beam_set_name}" have max delivery time >{max_del_time} s:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(bad_max_del)
                fail_msgs.append(msg)
            else:
                max_del = [f'{beam.Name} ({format_num(tss.BeamSettings[j].ArcConversionPropertiesPerBeam.MaxArcDeliveryTime)} s)' for j, beam in enumerate(beam_set.Beams)]
                msg = f'All beams have max delivery time &leq;{max_del_time} s:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(max_del)
                pass_msgs_section.append(msg)

            # Actual delivery time w/in 10% of what max should be
            bad_del_time, ok_del_time = [], []
            for b in beam_set.Beams:
                del_time = int(round(60 * b.BeamMU * sum(s.RelativeWeight / s.DoseRate for s in b.Segments if s.DoseRate != 0))) 
                del_time_str = f'{b.Name} ({format_num(del_time)} s)'
                if del_time > max_del_time * 1.1:
                    bad_del_time.append(del_time_str)
                else:
                    ok_del_time.append(del_time_str)
            if bad_del_time:
                msg = f'The following beams in beam set "{beam_set_name}" have delivery time >{max_del_time} s + 10%:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(bad_del_time)
                fail_msgs.append(msg)
            else:
                msg = f'All beams have delivery time &leq;{max_del_time} s + 10%:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(ok_del_time)
                pass_msgs_section.append(msg)

            # Constraint on max leaf distance per degree is enabled
            if not tss.SegmentConversion.ArcConversionProperties.UseMaxLeafTravelDistancePerDegree:
                msg = f'Constraint on leaf motion per degree is disabled for beam set "{beam_set_name}".'
                fail_msgs.append(msg)
            else:
                msg = 'Constraint on leaf motion per degree is enabled.'
                pass_msgs_section.append(msg)

                # Max distance per degree <=0.5 cm (only applies if this constraint is enabled)
                if tss.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree > 0.5:
                    msg = f'Max leaf motion per degree = {format_num(tss.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree)} > 0.5 cm/deg for beam set "{beam_set_name}".'
                    fail_msgs.append(msg)
                else:
                    msg = f'Max leaf motion per degree = {format_num(tss.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree)} &leq; 0.5 cm/deg.'
                    pass_msgs_section.append(msg)

        if bs_rx is not None and beam_set.FractionationPattern is not None and beam_set.FractionDose.DoseValues is not None:
            # Max dose
            dose_per_fx = float(bs_rx.DoseValue) / beam_set.FractionationPattern.NumberOfFractions
            max_dose = int(round(beam_set.FractionDose.GetDoseStatistic(RoiName=ext.Name, DoseType='Max') / dose_per_fx * 100))
            if plan_type in ['SRS', 'SBRT']:
                if max_dose > 125:
                    if max_dose > 140:
                        msg = f'Max dose for beam set "{beam_set_name}" = {max_dose}% > 140% Rx'
                        fail_msgs.append(msg)
                    else:
                        msg = f'Max dose for beam set "{beam_set_name}" = {max_dose}% > 125% Rx. This may be okay since it is below 140.'
                        warn_msgs.append(msg)
                else:
                    msg = f'Max dose = {max_dose}% &LessEqual; 125% Rx.'
                    pass_msgs_section.append(msg)
            elif plan_type == 'VMAT':
                if max_dose > 108:
                    if max_dose > 110:
                        msg = f'Max dose for beam set "{beam_set_name}" = {max_dose}% > 110% Rx.'
                        fail_msgs.append(msg)
                    else:
                        msg = f'Max dose for beam set "{beam_set_name}" = {max_dose}% Rx. Ideal is 107&ndash;108%, but {max_dose}% may be okay since it &leq; 110.'
                        warn_msgs.append(msg)
                else:
                    msg = f'Max dose = {max_dose}% &leq; 108% Rx.'
                    pass_msgs_section.append(msg)
            else:
                if max_dose > 110:
                    if max_dose > 118:
                        msg = f'Max dose for beam set "{beam_set_name}" = {max_dose}% > 118% Rx.'
                        fail_msgs.append(msg)
                    else:
                        msg = f'Max dose for beam set "{beam_set_name}" = {max_dose}% > 110% Rx. This may be okay since it &leq; 118.'
                        warn_msgs.append(msg)
                else:
                    msg = f'Max dose = {max_dose}% &leq; 110% Rx.'
                    pass_msgs_section.append(msg)
            
            if plan_type in ['VMAT', 'SRS', 'SBRT']:
                # Beam MU >= 110% beam dose
                too_modulated, modulation_ok = [], []
                for i, beam in enumerate(beam_set.Beams):
                    mu = beam.BeamMU
                    dose = beam_set.FractionDose.BeamDoses[i].DoseAtPoint.DoseValue
                    msg = f'{beam.Name} ({mu:.0f} MU, {dose:.0f} cGy dose)'
                    if mu < 1.1 * dose:
                        too_modulated.append(msg)
                    elif not too_modulated:
                        modulation_ok.append(msg)
                if too_modulated:
                    msg = f'The following beams in beam set "{beam_set_name}" may be too modulated:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(too_modulated)
                    warn_msgs.append(msg)
                else:
                    msg = 'Modulation is appropriate for all beams:' + PlanChkConstants.NEWLINE_AND_BULLET + PlanChkConstants.NEWLINE_AND_BULLET.join(modulation_ok)
                    pass_msgs_section.append(msg)

        if pass_msgs_section:
            pass_msgs[f'Beam Set "{beam_set_name}":'] = pass_msgs_section

    # Add red messages
    if fail_msgs:
        elems.extend([Paragraph('Errors:', style=PlanChkConstants.STYLES['Heading3']), PlanChkConstants.SPCR_SM])
        for msg in fail_msgs:
            elems.extend([Paragraph(msg, style=PlanChkConstants.STYLES[PlanChkConstants.FAIL]), PlanChkConstants.SPCR_LG])

    # Add yellow messages
    if warn_msgs:
        elems.extend([Paragraph('Warnings:', style=PlanChkConstants.STYLES['Heading3']), PlanChkConstants.SPCR_SM])
        for msg in warn_msgs:
            elems.extend([Paragraph(msg, style=PlanChkConstants.STYLES[PlanChkConstants.WARN]), PlanChkConstants.SPCR_LG])

    # Add green messages
    if pass_msgs:
        elems.extend([Paragraph('Passing:', style=PlanChkConstants.STYLES['Heading3']), PlanChkConstants.SPCR_SM])
        for heading, msgs in pass_msgs.items():
            elems.extend([Paragraph(heading, style=PlanChkConstants.STYLES['Heading4']), PlanChkConstants.SPCR_SM])
            
            for msg in msgs:
                elems.extend([Paragraph(msg, style=PlanChkConstants.STYLES[PlanChkConstants.PASS]), PlanChkConstants.SPCR_LG])

    ## 'Manual Checks:' section

    # Section header
    elems.extend([Paragraph('Manual Checks:', style=PlanChkConstants.STYLES['Heading3']), PlanChkConstants.SPCR_SM])
    
    # Messages to display in blue
    manual_msgs = ['Is admission date filled in in MOSAIQ? If not, ask the front desk to enter it.', 'Did the MD request any dose sums? Are they present in RS?', 'Are structures excluded from MOSAIQ export, and invisible? You may run script "Exclude from MOSAIQ Export".']
    
    # If any VMAT (incl. SRS, SBRT) plans, view MLC movie
    if any(plan_type in plan_types.values() for plan_type in ['VMAT', 'SRS', 'SBRT']):  
        manual_msgs.append('View MLC movie. There should be no weird/unexpected MLC positions, and MLC should approximately conform to PTV size.')
    
    # If Rx isodose is not 100%, remind user to double check MOSAIQ for this change
    rx = get_current('BeamSet').Prescription
    if rx is not None:
        rx = rx.PrimaryPrescriptionDoseReference
        if rx.PrescriptionType == 'DoseAtVolume' and rx.DoseVolume != 100:
            manual_msgs.extend([f'Is the Rx in MOSAIQ titled "{plan.Name}" to match the RS plan name?', f'The current beam set\'s primary Rx is to {format_num(rx.DoseVolume)}% volume, not 100%. Does this match in D and I in MOSAIQ?'])
    
    # Create blue check for each message
    for msg in manual_msgs:
        elems.extend([Paragraph(msg, style=PlanChkConstants.STYLES[PlanChkConstants.MANUAL]), PlanChkConstants.SPCR_LG])

    # Build PDF.
    pdf.build([KeepTogether(elem) for elem in elems])
    
    # Open report
    for reader_path in PlanChkConstants.ADOBE_READER_PATHS:
        try:
            os.system(fr'START /B "{reader_path}" "{filepath}"')
            break
        except:
            continue
