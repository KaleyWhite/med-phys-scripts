import clr
import os
import re
import sys

from connect import *  # Interact with RayStation

from PyPDF2 import PdfFileMerger, PdfFileReader

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # To display errors

sys.path.append(os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-scripts', 'RayStation'))
from sbrt_lung_analysis import sbrt_lung_analysis


class PrintRptConstants(object):
    """Class that defines several useful constants for this remainder of the script"""

    # ----------------- Clinic-specific constants. CHANGE THESE! ----------------- #
    
    # Name of the RayStation beam set report template to use
    PRINT_REPORT_TEMPLATE_NAME = 'ReportTemplateV11A_10132021'

    # Absolute path to the directory in which to create the report
    # Does not have to already exist
    PRINT_REPORT_OUTPUT_DIR = os.path.join('Z:', os.sep, 'TreatmentPlans')

    # --------------------------- No changes necessary --------------------------- #

    # Paths to Adobe Reader on RS servers
    ADOBE_READER_PATHS = [os.path.join('C:', os.sep, 'Program Files (x86)', 'Adobe', 'Reader 11.0', 'Reader', 'AcroRd32.exe'), os.path.join('C:', os.sep, 'Program Files (x86)', 'Adobe', 'Acrobat Reader DC', 'Reader', 'AcroRd32.exe')]


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


def print_report() -> None:
    """Prints and opens a report for the current beam set, appending an SBRT Lung Analysis report if applicable"""
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
        patient = get_current('Plan')
    except:
        MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
        sys.exit()
    try:
        beam_set = get_current('BeamSet')
    except:
        MessageBox.Show('There are no beam sets in the current plan. Click OK to abort the script.', 'No Beam Sets')
        sys.exit()
    clinic_db = get_current('ClinicDB')

    # Exit with an error if beam set has no dose
    if beam_set.FractionDose.DoseValues is None:
        MessageBox.Show('Beam set has no dose. Click OK to abort the script.', 'No Dose')
        sys.exit()

    # Exit with an error if dose statistics need updating
    if any(geom.HasContours() and plan.TreatmentCourse.TotalDose.GetDoseGridRoi(RoiName=geom.OfRoi.Name).RoiVolumeDistribution is None for geom in plan.GetTotalDoseStructureSet().RoiGeometries):
        MessageBox.Show('Dose statistics missing. Click OK to abort the script.', 'Missing Dose Statistics')
        sys.exit()

    # Report template
    try:
        template = next(t for t in clinic_db.GetSiteSettings().ReportTemplates if t.Name == PrintRptConstants.PRINT_REPORT_TEMPLATE_NAME)
    except StopIteration:
        # Display message if report template does not exist
        MessageBox.Show('The report template "' + PRINT_REPORT_TEMPLATE_NAME + '" does not exist. Click OK to abort the script.', 'Template Does Not Exist')
        sys.exit()

    patient.Save()  # Must save patient before report creation

    # Report filepath
    if not os.path.isdir(PrintRptConstants.PRINT_REPORT_OUTPUT_DIR):
        os.makedirs(PrintRptConstants.PRINT_REPORT_OUTPUT_DIR)
    pt_name = format_name(patient.Name)
    filename = pt_name + ' ' + beam_set.BeamSetIdentifier() + ' ' + datetime.now().strftime('%Y-%m-%d %H_%M_%S') + '.pdf'
    filepath = os.path.join(PrintRptConstants.PRINT_REPORT_OUTPUT_DIR, re.sub(r'[<>:"/\\\|\?\*]', '_', filename))
    
    # Report
    try:
        beam_set.CreateReport(templateName=PrintRptConstants.PRINT_REPORT_TEMPLATE_NAME, filename=filepath, ignoreWarnings=False)
    except Exception as e:
        res = MessageBox.Show('The script generated the following warnings:\n\n{}\nContinue?'.format(str(e).split('at ScriptClient')[0]), 'Warnings', MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit()
        beam_set.CreateReport(templateName=PrintRptConstants.PRINT_REPORT_TEMPLATE_NAME, filename=filepath, ignoreWarnings=True)

    # Append SBRT Lung Analysis report if this is SBRT lung plan
    if beam_set.Modality == 'Photons' and beam_set.PlanGenerationTechnique == 'Imrt' and beam_set.DeliveryTechnique == 'DynamicArc':
        rx = beam_set.Prescription
        if rx is not None:
            rx = sum(rx.DoseValue for rx in rx.PrescriptionDoseReferences)
            num_fx = beam_set.FractionationPattern.NumberOfFractions
            if num_fx == 5 and rx / num_fx >= 600 and case.BodySite in ['Thorax', '']:
                chk_for_body_site = [case.BodySite, case.CaseName, case.Comments, case.Diagnosis, plan.Comments, plan.Name] + [beam_set.DicomPlanLabel for beam_set in plan.BeamSets]
                if any(re.search(r'(^|[^A-Z])lung($|[^A-Z])', name, re.IGNORECASE) is not None for name in chk_for_body_site):
                    sbrt_filepath = sbrt_lung_analysis()
                    # Merge beam set report and SBRT report
                    merger = PdfFileMerger()
                    merger.append(PdfFileReader(open(filepath, 'rb')))
                    merger.append(PdfFileReader(open(sbrt_filepath, 'rb')))
                    merger.write(filepath)
                    merger.close()

    # Open report
    for reader_path in PrintRptConstants.ADOBE_READER_PATHS:
        try:
            os.system(fr'START /B "{reader_path}" "{filepath}"')
            break
        except:
            continue
