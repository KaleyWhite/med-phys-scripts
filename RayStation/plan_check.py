# For GUI
import clr
clr.AddReference("System.Windows.Forms")

import os
import re
import sys
from collections import OrderedDict
from math import sqrt

from connect import *
from reportlab.lib.colors import Blacker, Whiter, blue, green, red, yellow
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, KeepTogether
from reportlab.lib.units import inch
from System.Windows.Forms import *


case = plan = exam = struct_set = None


# ReportLab Paragraph styles
styles = getSampleStyleSheet()  # Base styles (e.g., "Heading1", "Normal")
h1 = ParagraphStyle(name="h1", parent=styles["Heading1"], fontName="Helvetica", fontSize=24, alignment=TA_CENTER)  # For patient name
h2 = ParagraphStyle(name="h2", parent=styles["Normal"], fontName="Helvetica", fontSize=18, alignment=TA_CENTER)  # For MR# and plan name
h3 = ParagraphStyle(name="h3", parent=styles["Normal"], fontName="Helvetica", fontSize=16)  # For headers "Errors:", "Warnings:", "Passing:", "Manual Checks:"
h4 = ParagraphStyle(name="h4", parent=styles["Normal"], fontName="Helvetica", fontSize=12)
normal = ParagraphStyle(name="normal", parent=styles["Normal"], fontName="Helvetica", fontSize=10)  # For checks
green = ParagraphStyle(name="green", parent=normal, backColor=Whiter(green, 0.25), borderPadding=7, borderWidth=1, borderColor=green, borderRadius=5)  # For passing checks: green background
yellow = ParagraphStyle(name="yellow", parent=green, backColor=Whiter(yellow, 0.25), borderColor=Blacker(yellow, 0.75))  # For warnings: yellow background
red = ParagraphStyle(name="red", parent=green, backColor=Whiter(red, 0.25), borderColor=red)  # For failing checks: red background
blue = ParagraphStyle(name="blue", parent=green, backColor=Whiter(blue, 0.25), borderColor=blue)  # For manual checks: blue background

# ReportLab Spacer objects to reuse for nice formatting
_, width = letter  # Need the width (8.5") for Spacer objects
spcr_sm = Spacer(width, 0.1 * inch)  # Small
spcr_lg = Spacer(width, 0.3 * inch)  # Large


def distance(a, b={"x": 0, "y": 0, "z": 0}):
    # Helper function that returns the distance between two points a and b
    # b defaults to the 3D origin

    return sqrt(sum((val - b[coord]) ** 2 for coord, val in a.items()))


def get_contour_coords(geom):
    # Return list of contour coordinates of a geometry
    # E.g., [{"x": 0, "y": 1, "z": 2}, ...]
    # Return the empty list if geometry is empty

    if not geom.HasContours():  # Empty geometry
        return []
    
    # If has contour representation, just flatten contour coords array and return
    # Otherwise, copy ROI, set copy's representation to contours, delete the copy, and return the copy's flattened contours array
    if hasattr(geom.PrimaryShape, "Contours"):
        coords = [c for coord in geom.PrimaryShape.Contours for c in coord]  # Flatten contours array
    else:  
        copy_name = name_item(geom.OfRoi.Name, [roi.Name for roi in case.PatientModel.RegionsOfInterest])  # Unique ROI name
        copy = case.PatientModel.CreateRoi(Name=copy_name, Color=geom.OfRoi.Color, Type=geom.OfRoi.Type)  # Create ROI w/ same color and type as geom's ROI
        copy.CreateAlgebraGeometry(Examination=exam, ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [geom.OfRoi.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="None", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})  # Copy is same as geom's ROI
        geom = case.PatientModel.StructureSets[exam.Name].RoiGeometries[copy_name]  # Change geom value so that correct contours are returned
        geom.SetRepresentation(Representation="Contours")  # Convert to contour representation
        coords = [c for coord in geom.PrimaryShape.Contours for c in coord]  # Flatten contours array
        copy.DeleteRoi()  # We no longer need the copy
    
    return coords


def format_coords(point):
    # Helper function that returns coordinmats formatted nicely for display
    # point: dictionary or ExpandoObject
    
    return "({}, {}, {})".format(format_num(point["x"]), format_num(point["z"]), format_num(-point["y"]))
    

def format_num(num):
    # Helper function to format a number for display
    # If there are no digits after the decimal place, return an int. Otherwise, strip all zeroes from the end and display a maximum of 2 decimal places

    # If number is not zero but rounds to 0, or number is greater than 100000, format in scientific notation
    if 0 < num < 0.005 or num > 100000:
        num = "{:.2E}".format(num)
        coef, power = num.split("E")
        coef = coef.rstrip("0").rstrip(".")
        if power.startswith("+"):
            power = power[1:].lstrip("0")
        else:
            power = "-{}".format(power[1:].lstrip("0"))
        if coef == "1":
            return "10<sup>{}</sup>".format(power)
        if coef == "-1":
            return "-10<sup>{}</sup>".format(power)
        return "{} &times; 10<sup>{}</sup>".format(coef, power)
        
    return str(round(num, 2)).rstrip("0").rstrip(".")


def name_item(item, l, max_len=sys.maxsize):
    # Helper function that generates a unique name for `item` in list `l` (case insensitive)
    # Limit name to `max_len` characters
    # E.g., name_item("Isocenter Name A", ["Isocenter Name A", "Isocenter Na (1)", "Isocenter N (10)"]) -> "Isocenter Na (2)"

    l_lower = [l_item.lower() for l_item in l]
    copy_num = 0
    old_item = item
    while item.lower() in l_lower:
        copy_num += 1
        copy_num_str = " ({})".format(copy_num)
        item = "{}{}".format(old_item[:(max_len - len(copy_num_str))].strip(), copy_num_str)
    return item[:max_len]


def will_gantry_collide(r, couch_bounds, ext_bounds, iso, likelihood):
    # Helper function that returns whether a geometry will likely collide with the gantry, using radius `r`
    cyl_name = name_item("Cylinder_{}".format(r), [roi.Name for roi in case.PatientModel.RegionsOfInterest], 16)
    cyl = case.PatientModel.CreateRoi(Name=cyl_name, Type="Control")
    cyl.CreateCylinderGeometry(Radius=r, Axis={"x": 0, "y": 0, "z": 1}, Length=30, Examination=exam, Center=iso)
    cyl_bounds = struct_set.RoiGeometries[cyl_name].GetBoundingBox()
    if couch_bounds:
        couch_collide = couch_bounds[0].x <= cyl_bounds[0].x or couch_bounds[0].y <= cyl_bounds[0].y or couch_bounds[1].x >= cyl_bounds[1].x or couch_bounds[1].y >= cyl_bounds[1].y
    else:
        couch_collide = False
    if ext_bounds:
        ext_collide = ext_bounds[0].x <= cyl_bounds[0].x or ext_bounds[0].y <= cyl_bounds[0].y or ext_bounds[1].x >= cyl_bounds[1].x or ext_bounds[1].y >= cyl_bounds[1].y
    else:
        ext_collide = False
    cyl.DeleteRoi()

    if couch_collide or ext_collide:
        if couch_collide and ext_collide:
            return "The couch and the patient are each >{} cm away from the gantry. Collision may be {}.".format(r, likelihood)
        if couch_collide:
            return "The couch is >{} cm away from the gantry. Collision may be {}.".format(r, likelihood)
        return "The patient is >{} cm away from the gantry. Collision may be {}.".format(r, likelihood)    


def plan_check():
    """Perform an "Initial Physics Review" plan check on the current plan

    Write a report to "T:\Physics\Scripts\Output Files\PlanCheck"
    Report is divided into several sections:
    - Errors (red): Things that really should be fixed
    - Warnings (yellow): Things that should be fixed but aren't just dire
    - Passing (green): Things that don't need to be fixed. This section is divided into subsections for Case, Plan, and each Beam Set.
    - Manual Checks (blue): Things the script can't check, so the user should check manually. These mainly relate to MOSAIQ.
    
    PlanCheck checks the following:
    - Case
        * External ROI exists.
        * If external ROI exists: External ROI is named 'External'.
        * The ROI named 'External', if it exists, is of type 'External'.
        * Case information is filled in: body site, diagnosis, physician name.
        * If physician name is present: Physician name includes "MD" suffix.
        * All exams have imaging system name "HOST-7307".
        * All exam names include a date (rough check - month, day, and year numbers are not validated).
        * For plans that are not VMAT H&N: Both couch ROIs exist.
        * If current plan is not an initial sim plan: Initial sim plan exists.
    - Plan
        * Planner is specified.
        * External is contoured on planning exam.
        * For prostate plans: The following ROIs exist: "Bladder", "Rectum", "Colon_Sigmoid", "Bag_Bowel".
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
        * Beam names are unique across all cases, with the exception of initial sim plans.
        * Setup beam names are unique across all cases.
        * Either a CT, or AP/PA and lat setup beam exist.
        * Machine is "SBRT6MV" for SBRT, "ELEKTA" otherwise.
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
        * Rx name in MOSAIQ matches Rx plan name in RS.
        * Any dose sums that the MD requested, are present in RS.
        * For VMAT plans: View MLC movie.
        * For VMAT plans: Non-100% isodose is specified in Rx in MOSAIQ.
        * Are the proper structures excluded from MOSAIQ export and invisible?

    Assumptions
    -----------
    Current beam set is the "main" beam set that will be exported to MOSAIQ.
    Each plan optimization optimizes a single beam set.
    Initial sim plan corresponding to this plan is the plan on the planning exam whose name contains Initial Sim" or "Trial_1" (case insensitive).
    Initial sim plan, if it exists, has just one beam set, and all beams in that beam set share the same isocenter.

    Bulleted lists are formatted in the following way:
    "Some message:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(some_list))
    (Message, then each bullet point on own line w/ 4 spaces before it and 2 spaces between it and its text)

    It is best practice to run the plan check before anything is approved.
    Iteratively make changed according to PlanChecks' errors/warnings and run PlanCheck again.
    """

    global case, plan, exam, struct_set

    # Get current variables
    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort script.", "No Patient Loaded")
        sys.exit(1)  # Exit script with an error
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)  # Exit script with an error
    try:
        plan = get_current("Plan")
    except:
        MessageBox.Show("There is no plan loaded. Click OK to abort script.", "No Plan Loaded")
        sys.exit(1)
    try:
        get_current("BeamSet")
    except:
        MessageBox.Show("There are no beam sets in the current plan. Click OK to abort script.", "No Beam Sets")
        sys.exit(1)  # Exit script with an error
    struct_set = plan.GetTotalDoseStructureSet()  # Structure set on planning exam
    exam = struct_set.OnExamination  # Planning exam

    # Need to determine plan types now so that we know if ANY plan types are ____
    # Example: Minimum dose grid voxel size depends on whether ANY beam set is SBRT
    # OrderedDict to retain original order of beam sets, since this dict is used later to iterate over beam sets
    # Dict elements are beam set : plan type ("SRS", "SBRT", "VMAT", "IMRT", or "3D")
    plan_types = OrderedDict()  
    for beam_set in plan.BeamSets:
        if beam_set.Modality == "Photons":  # Ignore beam sets that are not photons
            fx = beam_set.FractionationPattern
            if fx is not None:
                fx = fx.NumberOfFractions
            if beam_set.PlanGenerationTechnique == "Imrt":
                if beam_set.DeliveryTechnique == "DynamicArc":
                    if fx in [1, 3]:
                        plan_types[beam_set] = "SRS"
                    elif fx == 5:
                        plan_types[beam_set] = "SBRT"
                    else:
                        plan_types[beam_set] = "VMAT"
                else:
                    plan_types[beam_set] = "IMRT"
            else:
                plan_types[beam_set] = "3D"
    
    # Are there any photon plans?
    if not plan_types:
        MessageBox.Show("This is not a photon plan. Click OK to abort script.", "Plan Check")
        sys.exit(1)

    ## Prepare output file

    # Format patient name
    pt_name = patient.Name.split("^")  # e.g., ["Jones", "Bill", "P"]
    pt_name = "{}, {}".format(pt_name[0], pt_name[1])  # e.g., "Jones, Bill"

    # Get couch names
    template_name = "Elekta Couch" if "Supine" in beam_set.PatientPosition else "Elekta Prone Couch"
    template = get_current("PatientDB").LoadTemplatePatientModel(templateName=template_name)
    couch_names = outer_couch_name, inner_couch_name = [roi.Name for roi in template.PatientModel.RegionsOfInterest]  # [outer couch name, inner couch name]

    # Create PDF file
    filename = r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Output Files\PlanCheck\{} {}.pdf".format(pt_name, plan.Name)  # e.g., "Jones, Bill Prostate.pdf"
    pdf = SimpleDocTemplate(filename, pagesize=letter, bottomMargin=0.2 * inch, leftMargin=0.25 * inch, rightMargin=0.2 * inch, topMargin=0.2 * inch)  # 8.5 x 11" w/ 0.25" left & right margins, & 0.2" top & bottom margins

    # Headings
    hdg = Paragraph(pt_name, style=h1)  # e.g., "Jones, Bill"
    mrn = Paragraph("MR#: {}".format(patient.PatientID), style=h2)  # e.g., "MR#: 000123456"
    plan_chk = Paragraph("Plan Check: {}".format(plan.Name), style=h2)  # e.g., "Plan Check: Prostate"
    elems = [hdg, spcr_sm, mrn, spcr_sm, plan_chk, spcr_lg]  # List of elements to add to PDF later. Start w/ headings only, separated by spacing.

    is_vmat_hn = set(plan_types.values()) == {"VMAT"} and case.BodySite == "Head and Neck"  # Only VMAT H&N plans may lack couch
    
    # Lists of messages
    red_msgs = []  # Errors
    yellow_msgs = []  # Warnings
    green_msgs = OrderedDict()  # No problems. Dict instead of list so that categories can be used ("Case", "Plan", and a section for each beam set)
    # Manual checks come at the end of the report (after green messages)

    ## "Case:" section
    green_msgs_section = []

    # External ROI exists
    ext = [roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"]
    if not ext:
        msg = "There is no external ROI."
        red_msgs.append(msg)
    else:
        ext = ext[0]  # There will never be more than one external ROI
        msg = "External ROI exists."
        green_msgs_section.append(msg)

        # External is named "External", and no ROI of any other type is named "External"
        if ext.Name != "External":
            msg = "External ROI is not named 'External'."
            red_msgs.append(msg)
        else:
            msg = "External ROI is named 'External'."
            green_msgs_section.append(msg)

    # ROI named "External" is of type "External"
    external_type = [roi.Type for roi in case.PatientModel.RegionsOfInterest if roi.Name == "External" and roi.Type != "External"]
    if external_type:
        msg = "The ROI named 'External' is of type {}.".format(external_type[0])
        red_msgs.append(msg)
    else:
        msg = "The ROI named 'External' is of type 'External'."
        green_msgs_section.append(msg)

    # Case information is filled in
    case_attrs = {"Body site": case.BodySite, "Diagnosis": case.Diagnosis, "Physician name": case.Physician.Name}
    attrs = [attr for attr, case_attr in sorted(case_attrs.items()) if case_attr == ""]
    if attrs:
        msg = "Case information is missing:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(attrs))
        yellow_msgs.append(msg)
    else:
        case_attrs = ["{}: {}" for attr, case_attr in sorted(case_attrs.items())]  # e.g., "Body site: Thorax"
        msg = "All case information is filled in:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(case_attrs))
        green_msgs_section.append(msg)

    # MD name includes "MD" suffix
    if case.Physician.Name is not None:
        if not case.Physician.Name.endswith("^MD"):
            msg = "Physician name is missing 'MD' suffix: {}.".format(case.Physician.Name)
            red_msgs.append(msg)
        else:
            msg = "Physician name includes 'MD' suffix: {}.".format(case.Physician.Name)
            green_msgs_section.append(msg)

    # Exams: Imaging system name is HOST-7307
    wrong_img_sys = [e.Name for e in case.Examinations if e.EquipmentInfo.ImagingSystemReference is None or e.EquipmentInfo.ImagingSystemReference.ImagingSystemName != "HOST-7307"]
    if wrong_img_sys:
        msg = "The imaging system is incorrect for the following exams:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(wrong_img_sys))
        red_msgs.append(msg)
    else:
        msg = "All exams have imaging system HOST-7307."
        green_msgs_section.append(msg) 

    # Exam names include date
    date_regex_1 = "\d{1,2}[/\-. ]\d{1,2}[/\-. ](\d{4}|\d{2})"  # e.g., "1-24-2020"
    date_regex_2 = "\d{1,2}[/\-. ](Jan(uary)?|Feb(uary)?|Mar(ch)?|Apr(il)?|May|June?|July?|Aug(ust)?|Sep(t(ember)?)?|Oct(ober)?|Nov(ember)?|Dec(ember)?)[/\-. ](\d{4}|\d{2})"  # e.g., "24 Jan 2020"
    missing_date = [e.Name for e in case.Examinations if not re.match("\d{9} IMAGE FOR TEMPLATES", e.Name) and not re.search("({})|({})".format(date_regex_1, date_regex_2), e.Name)]  # Very crude date regex: does not validate month, day, or year numbers. Also, ignore "IMAGE FOR TEMPLATES" exams
    if missing_date:
        msg = "The following exam names are missing a date:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(missing_date))
        yellow_msgs.append(msg)
    else:
        msg = "All exam names include a date."
        green_msgs_section.append(msg)

    # Couch ROIs exist
    roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
    missing_couch_rois = [couch_name for couch_name in couch_names if couch_name not in roi_names]
    if not is_vmat_hn:
        if missing_couch_rois:
            msg = "Couch ROI(s) are missing:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(missing_couch_rois))
            if set(plan_types.values()) == {"3D"}:  # All plans are 3D
                yellow_msgs.append(msg)
            else:  # There are IMRT plans
                red_msgs.append(msg)
        else:
            msg = "Couch ROIs exist:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(couch_names))
            green_msgs_section.append(msg)

    # "Initial sim" plan exists
    ini_sim_plan = None
    if not re.search("(initial sim)|(trial_1)", plan.Name, re.IGNORECASE):
        ini_sim_plan = [p for p in case.TreatmentPlans if re.search("(initial sim)|(trial_1)", p.Name, re.IGNORECASE) and ("SBRT" in plan_types.values() or p.GetTotalDoseStructureSet().OnExamination.Name == exam.Name)]
        ini_sim_plan = ini_sim_plan[0] if ini_sim_plan else None
        if ini_sim_plan is None:
            msg = "There is no 'Initial Sim' plan."
            yellow_msgs.append(msg)
        else:
            msg = "'Initial Sim' plan is present."
            green_msgs_section.append(msg)

    # If there are any green case messages, add them to `green_msgs` dict
    if green_msgs_section:
        green_msgs["Case:"] = green_msgs_section

    ## "Plan:" section
    green_msgs_section = []

    dose_dist = plan.TreatmentCourse.TotalDose

    # Plan information is filled in
    if plan.PlannedBy == "":
        msg = "Planner is not specified."
        red_msgs.append(msg)
    else:
        msg = "Planner is specified: {}.".format(plan.PlannedBy)
        green_msgs_section.append(msg)

    if ext:
        # External has geometry on planning exam
        if not struct_set.RoiGeometries[ext.Name].HasContours():
            msg = "There is no external geometry on the planning exam."
            red_msgs.append(msg)
        else:
            msg = "External is contoured on planning exam."
            green_msgs_section.append(msg)

    # For prostate plans, ensure certain ROIs exist
    # We know it's a prostate plan if any of certain prostate-related keywords is in certain case/plan/beam set info fields
    chk_for_body_site = [case.BodySite, case.CaseName, case.Comments, case.Diagnosis, plan.Comments, plan.Name] + [beam_set.DicomPlanLabel for beam_set in plan.BeamSets]  # Fields to check for prostate keywords
    if any(site in attr for attr in chk_for_body_site for site in ["pros", "pb", "bed", "fossa"]):  # It's a prostate plan
        pros_rois = ["Bladder", "Rectum", "Colon_Sigmoid", "Bag_Bowel"]  # ROIs that must be present if this is a prostate plan
        roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
        missing_pros_rois = [pros_roi for pros_roi in pros_rois if pros_roi not in roi_names]
        if missing_pros_rois:
            msg = "Important prostate plan ROI(s) are missing:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(missing_pros_rois))
            red_msgs.append(msg)
        else:
            msg = "Important prostate plan ROIs exist:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(pros_rois))
            green_msgs_section.append(msg)

    # Empty geometries on planning exam
    empty_geom_names = [geom.OfRoi.Name for geom in struct_set.RoiGeometries if not geom.HasContours() and geom.OfRoi.Name != ext.Name]  # No external should be error (taken care of above), not warning
    if empty_geom_names:
        msg = "The following ROIs are empty on the planning exam:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(empty_geom_names))
        yellow_msgs.append(msg)
    else:
        msg = "There are no empty geometries on the planning exam."
        green_msgs_section.append(msg)

    # ROIs that have been updated since last voxel volume computation: have contours but no volume in dose grid
    dose_stats_missing = [geom.OfRoi.Name for geom in struct_set.RoiGeometries if geom.OfRoi.Name not in couch_names and geom.HasContours() and dose_dist.GetDoseGridRoi(RoiName=geom.OfRoi.Name).RoiVolumeDistribution is None]
    if dose_stats_missing:
        msg = "Dose statistics need updating."
        yellow_msgs.append(msg)
    else:
        msg = "All dose statistics are up to date."
        green_msgs_section.append(msg)

    # There are no contours "in the air" (outside external)
    # A contour extends outside external if any of its min coords are less than external min coordinates, or any of its max coordinates are greater than external max coordinates
    # Ignore coordinates outside the planning exam
    dg = plan.GetTotalDoseGrid()
    if ext and dose_dist.DoseValues is not None and len(set(dg.VoxelSize.values())) == 1:  # External exists, dose grid is defined, and dose grid voxel sizes are uniform
        ext_geom = struct_set.RoiGeometries[ext.Name]
        if ext_geom.HasContours():
            # Min and max coordinates of image
            img_stack = exam.Series[0].ImageStack
            exam_min, exam_max = img_stack.GetBoundingBox()

            # Min and max coordinates in dose grid, defined by a box geometry
            dg_min = dg.Corner
            dg_max = {dim: coord + dg.NrVoxels[dim] * dg.VoxelSize[dim] for dim, coord in dg_min.items()}
            box_name = name_item("DoseGrid", [roi.Name for roi in case.PatientModel.RegionsOfInterest], 16)
            box = case.PatientModel.CreateRoi(Name=box_name, Type="Control")
            box_min = {dim: max(coord, exam_min[dim]) for dim, coord in dg_min.items()}
            box_max = {dim: min(coord, exam_max[dim]) for dim, coord in dg_max.items()}
            box_ctr = {dim: (coord + box_max[dim]) / 2.0 for dim, coord in box_min.items()}
            box_sz = {dim: box_max[dim] - coord for dim, coord in box_min.items()}
            box.CreateBoxGeometry(Size=box_sz, Examination=exam, Center=box_ctr, VoxelSize=plan.GetTotalDoseGrid().VoxelSize.x)

            # Voxel indices of planning exam, and external w/ 3 mm margin
            ext_prv03_name = name_item("External_PRV03", [roi.Name for roi in case.PatientModel.RegionsOfInterest], 16)
            ext_prv03 = case.PatientModel.CreateRoi(Name=ext_prv03_name, Type="Control")
            ext_prv03.SetMarginExpression(SourceRoiName=ext.Name, MarginSettings={ 'Type': "Expand", 'Superior': 0.3, 'Inferior': 0.3, 'Anterior': 0.3, 'Posterior': 0.3, 'Right': 0.3, 'Left': 0.3 })
            ext_prv03.UpdateDerivedGeometry(Examination=exam)
            dose_dist.UpdateDoseGridStructures()
            box_vi = set(dose_dist.GetDoseGridRoi(RoiName=box_name).RoiVolumeDistribution.VoxelIndices)  # Voxel indices of box geometry ("voxel indices" of image)
            ext_vi = set(dose_dist.GetDoseGridRoi(RoiName=ext_prv03_name).RoiVolumeDistribution.VoxelIndices).intersection(box_vi)  # Voxel indices of external ROI that are inside the image

            # Delete unnecessary ROIs
            box.DeleteRoi()  # Box ROI no longer needed
            ext_prv03.DeleteRoi()

            stray_contours = []  # Geometries that extend outside external
            for geom in struct_set.RoiGeometries:
                if geom.HasContours() and geom.OfRoi.Type not in ["Bolus", "Control", "External", "FieldOfView", "Fixation", "Support"] and geom.OfRoi.RoiMaterial is None:  # Ignore the external contour, any other external contours, FOV or support (e.g., couch) contours, and ROIs with a material defined
                    geom_vi = set(dose_dist.GetDoseGridRoi(RoiName=geom.OfRoi.Name).RoiVolumeDistribution.VoxelIndices).intersection(box_vi)  # Voxel indices of the geometry that are inside the image
                    if not geom_vi.issubset(ext_vi):
                        stray_contours.append(geom.OfRoi.Name)
            
            if stray_contours:
                msg = "The following contours extend outside the external:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(stray_contours))
                red_msgs.append(msg)
            else:
                msg = "All contours are contained inside the external."
                green_msgs_section.append(msg)

        # External extends to couch w/o gap or overlap
        if "SBRT" in plan_types.values() and not missing_couch_rois and struct_set.RoiGeometries[outer_couch_name].HasContours() and struct_set.RoiGeometries[inner_couch_name].HasContours():
            ext_bottom = struct_set.RoiGeometries[ext.Name].GetBoundingBox()[1].y  # Bottom of external
            couch_top = struct_set.RoiGeometries[outer_couch_name].GetBoundingBox()[0].y
            diff = round(ext_bottom - couch_top, 2)
            if diff < -0.3:
                msg = "External and couch overlap by {} cm.".format(-diff)
                red_msgs.append(msg)
            elif diff > 0.3:
                msg = "There is a {}-cm gap between external and couch.".format(diff)
                red_msgs.append(msg)
            else:
                msg = "There is no overlap or gap between external and couch."
                green_msgs_section.append(msg)

    # No gap between adjacent boli
    boli = [geom for geom in struct_set.RoiGeometries if geom.HasContours() and geom.OfRoi.Type == "Bolus" and geom.OfRoi.DerivedRoiExpression is None]  # Non-derived bolus geometries
    if boli:  # Plan has bolus
        ext_coords = get_contour_coords(struct_set.RoiGeometries[ext.Name])  # All coordinates in External geometry
        gap_btwn_boli = []  # List of tuples of adjacent boli w/ a gap; e.g., [("Bolus 1", "Bolus 2"), ("Bolus 3", "Bolus 4")]
        
        # Iterate over each bolus, finding the adjacent (closest) bolus, as determined by smallest distance between closest two points that are also part of External
        for bolus in boli:
            bolus_coords = [coords for coords in get_contour_coords(bolus) if coords in ext_coords]  # All coordinates shared by bolus and External geometries

            # All other boli, from which to find adjacent bolus
            other_boli = boli[:]
            other_boli.remove(bolus)  # Bolus can't be adjacent to itself

            # Find adjacent bolus
            min_dist = float("Inf")  # Start min at largest possible so we'll be sure to find an adjacent bolus
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
            if gap and (adj_bolus, bolus) not in gap_btwn_boli:  # Prevent duplicates in gap list (e.g., ("Bolus 1", "Bolus 2") and ("Bolus 2", "Bolus 1"))
                gap_btwn_boli.append((bolus, adj_bolus))
        
        if gap_btwn_boli:
            gap_btwn_boli = ["{} and {}".format(pair[0], pair[1]) for pair in gap_btwn_boli]  # e.g., "Bolus 1 and Bolus 2"
            msg = "There is a gap between each of the follwing pair(s) of adjacent boli:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(gap_btwn_boli))
            red_msgs.append(msg)
        else:
            msg = "There are no gaps in the bolus."
            green_msgs_section.append(msg)

    # Dose grid doesn't extend outside image
    dg = plan.GetTotalDoseGrid()
    dg_high = {coord: dg.Corner[coord] + dg.NrVoxels[coord] * sz for coord, sz in dg.VoxelSize.items()}  # Max coordinates of dose grid (dose grid corner is min coordinates)
    if "SBRT" in plan_types.values() or "SRS" in plan_types.values():
        img_bounds = exam.Series[0].ImageStack.GetBoundingBox()
        if any(img_bounds[0][coord] < val for coord, val in dg.Corner.items()) or any(img_bounds[1][coord] > val for coord, val in dg_high.items()):
            msg = "Dose grid extends outside planning exam."
            yellow_msgs.append(msg)
        else:
            msg = "Planning exam contains all of dose grid."
            green_msgs_section.append(msg)

    # Dose grid includes all contours (except perhaps FOV)
    # A contour extends outside dose grid if any of its min coords are less than dose grid min coordinates, or any of its max coordinates are greater than dose grid max coordinates
    outside_dg = []  # Geometries that extend outside dose grid
    for geom in struct_set.RoiGeometries:  # Ignore empty geometries
        if geom.HasContours() and geom.OfRoi.Type != "FieldOfView" and not (geom.OfRoi.Type in ["Bolus", "Fixation", "Support"] and geom.OfRoi.RoiMaterial is not None):  # Ignore FOV, and bolus/fixation/support with material override
            bounds = geom.GetBoundingBox()  # [{"x": min x-coord, "y": min y-coord, "z": min z-coord}, {"x": max x-coord, "y": max y-coord, "z": max z-coord}]
            if any(bounds[0][coord] < val for coord, val in dg.Corner.items()) or any(bounds[1][coord] > val for coord, val in dg_high.items()):
                outside_dg.append(geom.OfRoi.Name)
    if outside_dg:
        msg = "Dose grid does not include all of the following geometries:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}.\nPlease review slices.".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(outside_dg))
        yellow_msgs.append(msg)
    else:
        msg = "All necessary contours are contained inside the dose grid."
        green_msgs_section.append(msg)

    # Uniform dose grid
    voxel_szs = ["{} = {:.0f} mm".format(coord, sz * 10) for coord, sz in sorted(dg.VoxelSize.items())]
    if not dg.VoxelSize.x == dg.VoxelSize.y == dg.VoxelSize.z:
        msg = "Dose grid voxel sizes are not uniform:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(voxel_szs))
        red_msgs.append(msg)
    else:
        msg = "Dose grid voxel sizes are uniform: x = y = z = {:.0f} mm.".format(dg.VoxelSize.x * 10)
        green_msgs_section.append(msg)

    # Dose grid voxel sizes are small enough
    max_sz = 2 if "SBRT" in plan_types.values() or "SRS" in plan_types.values() else 3  # 3 mm dose grid for non-SBRT, 2 mm for SBRT (incl. SRS)
    lg_voxels = ["{} = {:.0f} mm".format(coord, sz * 10) for coord, sz in dg.VoxelSize.items() if sz > max_sz]  # Coordinates whose voxel sizes are too large. Convert from cm to mm and display as integer, not float
    if lg_voxels:
        msg = "The following voxel sizes >{} mm:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(max_sz, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(lg_voxels))
        red_msgs.append(msg)
    else:
        voxel_szs = ["{} = {:.0f} mm".format(coord, sz * 10) for coord, sz in sorted(dg.VoxelSize.items())]
        msg = "All voxel sizes &leq;{} mm:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(max_sz, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(voxel_szs))
        green_msgs_section.append(msg)

    # Plan has Clinical Goals
    if plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count == 0:
        msg = "Plan has no Clinical Goals."
        yellow_msgs.append(msg)
    else:
        msg = "Plan has Clinical Goals."
        green_msgs_section.append(msg)

    # Lists of clinical goals that fail due to empty geometries, geometries updated since last voxel volume computation (need to run UpdateDoseGridStructures or click "Dose statistics missing" in GUI)
    # Display format uses similar logic to what RS presumably uses to display Clinical Goal text
    # E.g., "External: At most 6250 cGy dose at 0.03 cm^3 volume (6652 cGy)"
    if dose_dist.DoseValues is None:  # Plan has no dose
        msg = "Plan has no dose."
        red_msgs.append(msg)
    else:
        failing_goals = []
        for func in plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions:
            roi = func.ForRegionOfInterest
            if struct_set.RoiGeometries[roi.Name].HasContours() and roi.Name not in dose_stats_missing and not func.EvaluateClinicalGoal():
                goal_criteria = "At least" if func.PlanningGoal.GoalCriteria == "AtLeast" else "At most"
                goal_val = func.GetClinicalGoalValue()  # nan if empty or out-of-date geometry

                # Format text based on goal type
                goal_type = func.PlanningGoal.Type
                if goal_type == "DoseAtAbsoluteVolume":
                    accept_lvl = "{} cGy dose".format(format_num(func.PlanningGoal.AcceptanceLevel))
                    param_val = " at {} cm<sup>3</sup> volume".format(format_num(func.PlanningGoal.ParameterValue))
                    goal_val = "{} cGy".format(format_num(func.GetClinicalGoalValue()))
                elif goal_type == "DoseAtVolume":
                    accept_lvl = "{} cGy dose".format(format_num(func.PlanningGoal.AcceptanceLevel))
                    param_val = " at {}% volume".format(format_num(func.PlanningGoal.ParameterValue * 100))
                    goal_val = "{} cGy".format(format_num(func.GetClinicalGoalValue()))
                elif goal_type == "AbsoluteVolumeAtDose":
                    accept_lvl = "{} cm<sup>3</sup> volume".format(format_num(func.PlanningGoal.AcceptanceLevel))
                    param_val = " at {} cGy dose".format(format_num(func.PlanningGoal.ParameterValue))
                    goal_val = "{} cm<sup>3</sup>".format(format_num(func.GetClinicalGoalValue()))
                elif goal_type == "VolumeAtDose":
                    accept_lvl = "{}% volume".format(format_num(func.PlanningGoal.AcceptanceLevel * 100))
                    param_val = " at {} cGy dose".format(format_num(func.PlanningGoal.ParameterValue))
                    goal_val = "{}%".format(format_num(func.GetClinicalGoalValue() * 100)) 
                elif goal_type == "AverageDose":
                    accept_lvl = "{} cGy average dose".format(format_num(func.PlanningGoal.AcceptanceLevel))
                    param_val = ""
                    goal_val = "{} cGy".format(format_num(func.GetClinicalGoalValue()))
                elif goal_type == "ConformityIndex":
                    accept_lvl = "a conformity index of {}".format(format_num(func.PlanningGoal.AcceptanceLevel))
                    param_val = " at {} cGy dose".format(format_num(func.PlanningGoal.ParameterValue))
                    goal_val = format_num(func.GetClinicalGoalValue())
                elif goal_type == "HomogeneityIndex":  # HI
                    accept_lvl = "a homogeneity index of {}".format(format_num(func.PlanningGoal.AcceptanceLevel))
                    param_val = " at {}% volume".format(format_num(func.PlanningGoal.ParameterValue * 100))
                    goal_val = format_num(func.GetClinicalGoalValue())
                else:  # DoseAtPoint
                    accept_lvl = "{} cGy dose at point".format(format_num(func.PlanningGoal.AcceptanceLevel))
                    param_val = ""
                    goal_val = "{} cGy".format(format_num(func.GetClinicalGoalValue()))

                goal = "{}: {} {}{}".format(roi.Name, goal_criteria, accept_lvl, param_val)
                failing_goals.append("{} ({})".format(goal, goal_val))

        if failing_goals:
            msg = "The following Clinical Goals fail:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(failing_goals))
            red_msgs.append(msg)
        else:
            msg = "All evaluable Clinical Goals pass."
            green_msgs_section.append(msg)

    # Gantry does not collide with couch or patient
    # Max distance between plan isocenter and couch/Skin should be <40 cm, at worst <41.5 cm
    # Create cylinder geometry w/ radius 40 and 41.5, and height the diameter of the collimator
    # Min and max couch/skin coordinates must not be outside cylinder min and max coordinates

    # Plan isocenter coordinates
    iso = struct_set.LocalizationPoiGeometry
    if iso is None or iso.Point is None or any(abs(coord) > 1000 for coord in iso.Point.values()):
        msg = "Plan has no localization geometry."
        red_msgs.append(msg)
    else:
        iso = iso.Point
        msg = "Localization geometry is defined."    
        green_msgs_section.append(msg)

        # Bounds of couch and external
        if not missing_couch_rois and struct_set.RoiGeometries[outer_couch_name].HasContours():
            couch_bounds = struct_set.RoiGeometries[outer_couch_name].GetBoundingBox()
        else:
            couch_bounds = None
        if ext and struct_set.RoiGeometries[ext.Name].HasContours():
            ext_bounds = struct_set.RoiGeometries[ext.Name].GetBoundingBox()
        else:
            ext_bounds = None

        msg = will_gantry_collide(40, couch_bounds, ext_bounds, iso, "very likely")  # 40 cm cylinder
        if msg is not None:
            red_msgs.append(msg)
        else:
            msg = will_gantry_collide(41.5, couch_bounds, ext_bounds, iso, "likely")  # 41.5 cm cylinder
            if msg is not None:
                yellow_msgs.append(msg)
            else:
               green_msgs_section.append("Distance from gantry to patient/couch &GreaterEqual;41.5 cm. Collision is unlikely.") 

    # If there are any green case messages, add them to `green_msgs` dict
    if green_msgs_section:
        green_msgs["Plan:"] = green_msgs_section

    ## Beam set sections

    # List of all beam names (including setup beams), for preventing duplicate beam names
    beam_names = []
    for c in patient.Cases:
        for p in c.TreatmentPlans:
            if p.Name.lower() not in ["initial sim", "trial_1"]:
                for bs in p.BeamSets:
                    beam_names.extend([b.Name for b in bs.Beams])
                    beam_names.extend([sb.Name for sb in bs.PatientSetup.SetupBeams])

    # Iterate over beam sets in `plan_types` dict instead of plan.BeamSets b/c we only want photon beam sets
    for i, beam_set in enumerate(plan_types):
        green_msgs_section = []

        beam_set_name = beam_set.DicomPlanLabel
        plan_type = plan_types[beam_set]
        bs_rx = beam_set.Prescription.PrimaryPrescriptionDoseReference

        # Beam name = beam number
        bad_names = ["{} (#{})".format(beam.Name, beam.Number) for beam in beam_set.Beams if beam.Name != str(beam.Number)]  # e.g., "CCW (#2)"
        if bad_names:
            msg = "The following beam names in beam set '{}' are not the same as their numbers:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(bad_names))
            yellow_msgs.append(msg)
        else:
            msg = "Beam names are the same as their numbers."
            green_msgs_section.append(msg)

        # Duplicate beam names
        dup_names = [beam.Name for beam in beam_set.Beams if beam_names.count(beam.Name) > 1]
        if dup_names:
            msg = "The following beam names in beam set '{}' exist in other cases or plans:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(dup_names))
            red_msgs.append(msg)
        else:
            msg = "No beam name already exists in a previous case or plan."
            green_msgs_section.append(msg)

        # Duplicate setup beam names
        sbs = beam_set.PatientSetup.SetupBeams
        dup_names = [sb.Name for sb in sbs if beam_names.count(sb.Name) > 1]
        if dup_names:
            msg = "The following setup beam names in beam set '{}' exist in other cases or plans:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(dup_names))
            red_msgs.append(msg)
        else:
            msg = "No setup beam name already exists in another case or plan."
            green_msgs_section.append(msg)

        # Setup beams present: either AP + Lat kV, or CBCT (preferred)
        gantry_angles = [sb.GantryAngle for sb in sbs]
        if sbs.Count == 0:
            msg = "There are no setup beams for beam set '{}'. There should be either a CBCT, or an AP/PA and a lat.".format(beam_set_name)
            yellow_msgs.append(msg)
        elif any(re.search("C[BT]", sb.Name, re.IGNORECASE) or re.search("C[BT]", sb.Description, re.IGNORECASE) for sb in sbs):
            msg = "CBCT setup beam exists."
            green_msgs_section.append(msg)
        elif (0 in gantry_angles or 180 in gantry_angles) and (90 in gantry_angles or 270 in gantry_angles):
            msg = "AP/PA and lat setup beams exist."
            green_msgs_section.append(msg)
        else:
            msg = "Beam set '{}' contains neither a CBCT setup beam, nor both an AP/PA and a lat setup beam.".format(beam_set_name)
            yellow_msgs.append(msg) 

        # Machine is ELEKTA or SBRT 6MV
        machine = "SBRT 6MV" if plan_type in ["SRS", "SBRT"] else "ELEKTA"
        if beam_set.MachineReference.MachineName != machine:
            msg = "Machine for beam set '{}' should be '{}', not '{}'.".format(beam_set_name, machine, beam_set.MachineReference.MachineName)
            red_msgs.append(msg)
        else:
            msg = "Machine is '{}'.".format(machine)
            green_msgs_section.append(msg)
           
        # z-coordinate of beam isos between -100 and 100
        lg_z = ["{} ({} cm)".format(beam.Name, format_num(beam.Isocenter.Position.z)) for beam in beam_set.Beams if abs(beam.Isocenter.Position.z) > 100]  # e.g., "2 (105 cm)"
        if lg_z:
            msg = "The following beams in beam set '{}' have isocenter z-coordinate > 100 cm:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(lg_z))
            red_msgs.append(msg)
        else:
            z = ["{} ({} cm)".format(beam.Name, format_num(beam.Isocenter.Position.z)) for beam in beam_set.Beams]
            msg = "All beams have isocenter z-coordinate &leq; 100 cm:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(z))
            green_msgs_section.append(msg)        

        # Beam set has Rx
        if bs_rx is None:
            msg = "There is no prescription for beam set '{}'.".format(beam_set_name)
            red_msgs.append(msg)
        else:
            msg = "Beam set has a prescription."
            green_msgs_section.append(msg)

        # Beam set has dose
        if beam_set.FractionDose.DoseValues is None:
            msg = "Beam set '{}' has no dose.".format(beam_set_name)
            red_msgs.append(msg)
        else:
            msg = "Beam set has dose."
            green_msgs_section.append(msg)

            # DSPs are near target
            # For Rx to volume, DSP is within PTV
            # For Rx to point, DSP is within 80% isodose line (create a geometry from dose to determine this)
            if bs_rx is not None:  # There is an Rx, and it is to volume or to point
                if bs_rx.PrescriptionType == "DoseAtVolume":  # Rx to volume
                    roi = bs_rx.OnStructure  # PTV
                    dsp_roi = "PTV"
                else:  # Rx to point
                    roi_name = name_item("IDL_80%", [r.Name for r in case.PatientModel.RegionsOfInterest])  # We'll create an ROI w/ this name
                    roi = case.PatientModel.CreateRoi(Name=roi_name, Type="Control")  # Create ROI
                    roi.CreateRoiGeometryFromDose(DoseDistribution=dose_dist, ThresholdLevel=0.8 * bs_rx.DoseValue)  # Set geometry to isodose line for 80% of the Rx
                    dsp_roi = "80% isodose line"

                roi_bounds = struct_set.RoiGeometries[roi.Name].GetBoundingBox()
                bad_dsps = []
                for dsp in beam_set.DoseSpecificationPoints:
                    if any(dsp.Coordinates[coord] < val for coord, val in roi_bounds[0].items()) or any(dsp.Coordinates[coord] > val for coord, val in roi_bounds[1].items()):
                        bad_dsps.append("{}: {}".format(dsp.Name, format_coords(dsp.Coordinates)))

                # Delete IDL ROI
                if roi.Type == "Control":  # Delete the IDL ROI if it exists
                    roi.DeleteRoi() 

                if bad_dsps:
                    msg = "The following DSPs for beam set '{}' are outside the {}:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, dsp_roi, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(bad_dsps))
                    red_msgs.append(msg)
                else:
                    msg = "All DSPs are inside the {}.".format(dsp_roi)
                    green_msgs_section.append(msg)         

        # Iso has not been changed from initial sim
        if ini_sim_plan is not None and ini_sim_plan.BeamSets[0].Beams.Count > 0:
            ini_sim_iso = ini_sim_plan.BeamSets[0].Beams[0].Isocenter
            loc_pt = beam_set.GetStructureSet().LocalizationPoiGeometry
            if loc_pt is None or loc_pt.Point is None or any(abs(coord) > 1000 for coord in loc_pt.Point.values()):
                iso_chged = ["{}: {} {}".format(b.Name, b.Isocenter.Annotation.Name, format_coords(b.Isocenter.Position)) for b in beam_set.Beams if format_coords(b.Isocenter.Position) != format_coords(ini_sim_iso.Position)]
                if iso_chged:
                    msg = "Isocenter coordinates for the following beams in beam set '{}' were changed from initial sim {}:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, format_coords(ini_sim_iso.Position), "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(iso_chged))
                    yellow_msgs.append(msg)
                else:
                    msg = "No isocenter coordinates were changed from initial sim {}.".format(format_coords(ini_sim_iso.Position))
                    green_msgs_section.append(msg)
            else:
                print('beam set:', beam_set)
                print('ini sim plan:', ini_sim_plan)
                print('loc pt:', loc_pt, loc_pt.Point)
                print('ini sim iso:', ini_sim_iso, ini_sim_iso.Position)
                loc_pt = loc_pt.Point
                if format_coords(loc_pt) != format_coords(ini_sim_iso.Position):
                    msg = "Loc point coordinates for beam set '{}' {} were changed from initial sim iso coordinates {}".format(beam_set_name, format_coords(loc_pt), format_coords(ini_sim_iso.Position))
                    yellow_msgs.append(msg)
                else:
                    msg = "Loc point coordinates and initial sim iso coordinates {} are the same.".format(format_coords(ini_sim_iso.Position))
                    green_msgs_section.append(msg)
          
        # Dose algorithm is Collapsed Cone (CCDose)
        if beam_set.AccurateDoseAlgorithm.DoseAlgorithm != "CCDose":
            msg = "Dose algorithm for beam set '{}' is '{}'. It should be 'CCDose'.".format(beam_set_name, beam_set.AccurateDoseAlgorithm.DoseAlgorithm)
            red_msgs.append(msg)
        else:
            msg = "Dose algorithm is 'CCDose'."
            green_msgs_section.append(msg)

        # Plan optimization settings (there is a different plan optimization object for each beam set)
        opt = plan.PlanOptimizations[i]
        if not opt.AutoScaleToPrescription:
            msg = "Autoscale to prescription is disabled for beam set '{}'.".format(beam_set_name)
            red_msgs.append(msg)
        else:
            msg = "Autoscale to prescription is enabled."
            green_msgs_section.append(msg)

        # The following checks are for VMAT (incl. SRS, SBRT) only
        if plan_type in ["VMAT", "SRS", "SBRT"]:
            # Beam energy = 6 MV
            bad_energy = ["{} ({} MV)".format(beam.Name, beam.BeamQualityId) for beam in beam_set.Beams if beam.BeamQualityId != '6']  # e.g., "CCW (18 MV)"
            if bad_energy:
                msg = "The following beam energies in beam set '{}' should be 6 MV:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(bad_energy))
                red_msgs.append(msg)
            else:
                msg = "All beam energies are 6 MV."
                green_msgs_section.append(msg)    

            # Optimization tolerance <=10^-5
            if opt.OptimizationParameters.Algorithm.OptimalityTolerance > 0.0001:
                msg = "Optimization tolerance for beam set '{}' = {} > 10<sup>-5</sup>.".format(format_num(opt.OptimizationParameters.Algorithm.OptimalityTolerance), beam_set_name)
                red_msgs.append(msg)
            else:
                msg = "Optimization tolerance = {} &leq; 10<sup>-5</sup>.".format(format_num(opt.OptimizationParameters.Algorithm.OptimalityTolerance))
                green_msgs_section.append(msg)

            # ComputeIntermediateDose is checked
            if "lung" in [beam_set_name, plan.Name, case.CaseName] and not opt.OptimizationParameters.DoseCalculation.ComputeIntermediateDose:
                msg = "'Compute intermediate dose' is unchecked for beam set '{}'.".format(beam_set_name)
                yellow_msgs.append(msg)
            else:
                msg = "'Compute intermediate dose' is checked."
                green_msgs_section.append(msg)

            # Gantry spacing <=3 cm
            tss = opt.OptimizationParameters.TreatmentSetupSettings[0]
            bad_gantry_spacing = ["{} ({}&deg;)".format(beam.Name, format_num(tss.BeamSettings[j].ArcConversionPropertiesPerBeam.FinalArcGantrySpacing)) for j, beam in enumerate(beam_set.Beams) if tss.BeamSettings[j].ArcConversionPropertiesPerBeam.FinalArcGantrySpacing > 3]
            if bad_gantry_spacing:
                msg = "The following beams in beam set '{}' have gantry spacing >3&deg;:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(bad_gantry_spacing))
                red_msgs.append(msg)
            else:
                gantry_spacing = ["{} ({}&deg;)".format(beam.Name, format_num(tss.BeamSettings[j].ArcConversionPropertiesPerBeam.FinalArcGantrySpacing)) for j, beam in enumerate(beam_set.Beams)]
                msg = "All beams have gantry spacing &leq; 3&deg;:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(gantry_spacing))
                green_msgs_section.append(msg)

            # Max delivery time <=120 or 180 s
            max_del_time = 180 if plan_type in ["SRS", "SBRT"] else 120
            bad_max_del = ["{} ({} s)".format(beam.Name, format_num(tss.BeamSettings[j].ArcConversionPropertiesPerBeam.MaxArcDeliveryTime)) for j, beam in enumerate(beam_set.Beams) if tss.BeamSettings[j].ArcConversionPropertiesPerBeam.MaxArcDeliveryTime > max_del_time]
            if bad_max_del:
                msg = "The following beams in beam set '{}' have max delivery time >{} s:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, max_del_time, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(bad_max_del))
                red_msgs.append(msg)
            else:
                max_del = ["{} ({} s)".format(beam.Name, format_num(tss.BeamSettings[j].ArcConversionPropertiesPerBeam.MaxArcDeliveryTime)) for j, beam in enumerate(beam_set.Beams)]
                msg = "All beams have max delivery time &leq;{} s:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(max_del_time, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(max_del))
                green_msgs_section.append(msg)

            # Actual delivery time w/in 10% of what max should be
            bad_del_time, ok_del_time = [], []
            for b in beam_set.Beams:
                del_time = int(round(60 * b.BeamMU * sum(s.RelativeWeight / s.DoseRate for s in b.Segments if s.DoseRate != 0))) 
                if del_time > max_del_time * 1.1:
                    bad_del_time.append("{} ({} s)".format(b.Name, format_num(del_time)))
                else:
                    ok_del_time.append("{} ({} s)".format(b.Name, format_num(del_time)))
            if bad_del_time:
                msg = "The following beams in beam set '{}' have delivery time >{} s + 10%:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set_name, max_del_time, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(bad_del_time))
                red_msgs.append(msg)
            else:
                msg = "All beams have delivery time &leq;{} s + 10%:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(max_del_time, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(ok_del_time))
                green_msgs_section.append(msg)

            # Constraint on max leaf distance per degree is enabled
            if not tss.SegmentConversion.ArcConversionProperties.UseMaxLeafTravelDistancePerDegree:
                msg = "Constraint on leaf motion per degree is disabled for beam set '{}'.".format(beam_set_name)
                red_msgs.append(msg)
            else:
                msg = "Constraint on leaf motion per degree is enabled."
                green_msgs_section.append(msg)

                # Max distance per degree <=0.5 cm (only applies if this constraint is enabled)
                if tss.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree > 0.5:
                    msg = "Max leaf motion per degree = {} > 0.5 cm/deg for beam set '{}'.".format(beam_set_name, format_num(tss.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree))
                    red_msgs.append(msg)
                else:
                    msg = "Max leaf motion per degree = {} &leq; 0.5 cm/deg.".format(format_num(tss.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree))
                    green_msgs_section.append(msg)

        if bs_rx is not None and beam_set.FractionationPattern is not None and beam_set.FractionDose.DoseValues is not None:
            # Max dose
            dose_per_fx = float(bs_rx.DoseValue) / beam_set.FractionationPattern.NumberOfFractions
            max_dose = int(round(beam_set.FractionDose.GetDoseStatistic(RoiName=ext.Name, DoseType="Max") / dose_per_fx * 100))
            if plan_type in ["SRS", "SBRT"]:
                if max_dose > 125:
                    if max_dose > 140:
                        msg = "Max dose for beam set '{}' = {}% > 140% Rx".format(beam_set_name, max_dose)
                        red_msgs.append(msg)
                    else:
                        msg = "Max dose for beam set '{}' = {}% > 125% Rx. This may be okay since it is below 140.".format(beam_set_name, max_dose)
                        yellow_msgs.append(msg)
                else:
                    msg = "Max dose = {}% &LessEqual; 125% Rx.".format(max_dose)
                    green_msgs_section.append(msg)
            elif plan_type == "VMAT":
                if max_dose > 108:
                    if max_dose > 110:
                        msg = "Max dose for beam set '{}' = {}% > 110% Rx.".format(beam_set_name, max_dose)
                        red_msgs.append(msg)
                    else:
                        msg = "Max dose for beam set '{0}' = {1}% Rx. Ideal is 107&ndash;108%, but {1}% may be okay since it &leq; 110.".format(beam_set_name, max_dose)
                        yellow_msgs.append(msg)
                else:
                    msg = "Max dose = {}% &leq; 108% Rx.".format(max_dose)
                    green_msgs_section.append(msg)
            else:
                if max_dose > 110:
                    if max_dose > 118:
                        msg = "Max dose for beam set '{}' = {}% > 118% Rx.".format(beam_set_name, max_dose)
                        red_msgs.append(msg)
                    else:
                        msg = "Max dose for beam set '{}' = {}% > 110% Rx. This may be okay since it &leq; 118.".format(beam_set_name, max_dose)
                        yellow_msgs.append(msg)
                else:
                    msg = "Max dose = {}% &leq; 110% Rx.".format(max_dose)
                    green_msgs_section.append(msg)
            if plan_type in ["VMAT", "SRS", "SBRT"]:
                # Beam MU >= 110% beam dose
                too_modulated, modulation_ok = [], []
                for i, beam in enumerate(beam_set.Beams):
                    mu = beam.BeamMU
                    dose = beam_set.FractionDose.BeamDoses[i].DoseAtPoint.DoseValue
                    msg = "{} ({:.0f} MU, {:.0f} cGy dose)".format(beam.Name, mu, dose)
                    if mu < 1.1 * dose:
                        too_modulated.append(msg)
                    elif not too_modulated:
                        modulation_ok.append(msg)
                if too_modulated:
                    msg = "The following beams in beam set '{}' may be too modulated:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format(beam_set.DicomPlanLabel, "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(too_modulated))
                    yellow_msgs.append(msg)
                else:
                    msg = "Modulation is appropriate for all beams:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;{}".format("<br/>&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;".join(modulation_ok))
                    green_msgs_section.append(msg)

        if green_msgs_section:
            green_msgs["Beam Set '{}':".format(beam_set.DicomPlanLabel)] = green_msgs_section

    # Add red messages
    if red_msgs:
        # Section header
        hdr = Paragraph("Errors:", style=h3)
        elems.extend([hdr, spcr_lg])

        # Red messages
        for msg in red_msgs:
            elems.extend([Paragraph(msg, style=red), spcr_lg])

    # Add yellow messages
    if yellow_msgs:
        # Section header
        hdr = Paragraph("Warnings:", style=h3)
        elems.extend([hdr, spcr_lg])

        # Yellow messages
        for msg in yellow_msgs:
            elems.extend([Paragraph(msg, style=yellow), spcr_lg])

    # Add green messages
    if green_msgs:
        # Section header
        hdr = Paragraph("Passing:", style=h3)
        elems.extend([hdr, spcr_lg])
        for heading, msgs in green_msgs.items():
            hdr = Paragraph(heading, style=h4)
            elems.extend([hdr, spcr_lg])

            # Green messages
            for msg in msgs:
                elems.extend([Paragraph(msg, style=green), spcr_lg])

    ## "Manual Checks:" section

    # Section header
    manual_hdr = Paragraph("Manual Checks:", style=h3)
    elems.extend([manual_hdr, spcr_lg])
    
    # Messages to display in blue
    msgs = ["Is admission date filled in in MOSAIQ? If not, ask Amber Spicer (x2041) to enter it.", "Is the Rx in MOSAIQ titled '{}' to match the RS plan name?".format(plan.Name), "Did the MD request any dose sums? Are they present in RS?", "Are structures excluded from MOSAIQ export, and invisible? You may run script ExcludeFromMOSAIQExport."]
    
    # If any VMAT (incl. SRS, SBRT) plans, view MLC movie
    if any(plan_type in plan_types.values() for plan_type in ["VMAT", "SRS", "SBRT"]):  
        msgs.append("View MLC movie. There should be no weird/unexpected MLC positions, and MLC should approximately conform to PTV size.")
    
    # If Rx isodose is not 100%, remind user to double check MOSAIQ for this change
    rx = get_current("BeamSet").Prescription.PrimaryPrescriptionDoseReference
    if rx is not None and rx.PrescriptionType == "DoseAtVolume" and rx.DoseVolume != 100:
        msgs.append("The current beam set's Rx is to {}% volume, not 100%. Does this match in D and I in MOSAIQ?".format(format_num(rx.DoseVolume)))
    
    # Create blue check for each message
    for msg in msgs:
        elems.append(Paragraph(msg, style=blue))
        elems.append(spcr_lg)

    # Build PDF. If a file with this name is already open, alert user and exit script with an error.
    try:
        pdf.build([KeepTogether(elem) for elem in elems])
    except:
        MessageBox.Show("A plan check for this plan is open. Click OK to abort the script.")
        sys.exit(1)
    
    # Open report
    reader_paths = [r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe", r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"]  # Paths to Adobe Reader on RS servers
    for reader_path in reader_paths:
        try:
            os.system(r'START /B "{}" "{}"'.format(reader_path, filename))
            #os.system('start "{}" "{}"'.format(reader_path, filename))
            break
        except:
            continue
