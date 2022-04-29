import os
import shutil
import sys
from typing import List

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons


# Name of the colorwash template to apply
COLORWASH_TEMPLATE = 'CRMC Standard Dose Colorwash'

# Absolute path to the directory to export DICOM files to
EXPORT_PATH = os.path.join('T:', os.sep, 'Physics', 'QA & Procedures', 'Delta4', 'DQA Plans')


def unique_name(desired_name: str, existing_names: List[str]) -> str:
    '''Makes the desired name unique among all names in the list

    Name is made unique with a copy number in parentheses
    New name is truncated to be at most 16 charcaters long

    Arguments
    ---------
    desired_name: The new name to make unique
    existing_names: List of names among which the new name must be unique

    Returns
    -------
    The new name, made unique

    Example
    -------
    unique_name('1234567890abcdef', ['hello']) -> '1234567890abcdef'
    unique_name('1234567890abcdef', ['1234567890ab (1)']) -> '1234567890ab (2)'
    '''
    new_name = desired_name[:16]  # Truncate to at most 16 characters
    copy_num = 0  # Assume no copies
    # Increment the copy number until it makes the name unique
    while new_name in existing_names:
        copy_num += 1
        copy_str = ' (' + str(copy_num) + ')'  # Suffix to add the name to make it unique
        name_len = 16 - len(copy_str)  # Number of characters allowed before the suffix
        new_name = desired_name[:name_len] + copy_str
    return new_name


def format_pt_name(pt: PyScriptObject) -> str:
    """Converts the patient's Name attribute into a better format for display

    Arguments
    ---------
    pt: The patient whose name to format

    Returns
    -------
    The formatted patient name

    Example
    -------
    Given some_pt with Name "^Jones^Bill^^M":
    format_pt_name(some_pt) -> "Jones, Bill M"
    """
    parts = [part for part in re.split(r'\^+', pt.Name) if part != '']
    name = parts[0]
    if len(parts) > 0:
        name += ', ' + ' '.join(parts[1:])
    return name


def create_qa_plan() -> None:
    """Creates and exports a DQA plan for the current beam set
    
    If export folder with the computed name already exists, delete and recreate it

    Detailed example of the code that names the new QA plan:
    Existing QA plans are 'Test' and 'Rectal Boost DQA'
    New QA plan for plan 'Rectal Boost' would be 'Rectal Boost DQA (1)', but name length must <=16 characters, so truncate the plan name to get QA plan name 'Rectal B DQA (1)'
    If we create another QA plan for 'Rectal Boost', it is named 'Rectal B (2)'
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
    except:
        MessageBox.Show('There are no plans in the current case. Click OK to abort the script.', 'No Plans')
        sys.exit()
    try:
        beam_set = get_current('BeamSet')
    except:
        MessageBox.Show('There are no beam sets in the current plan. Click OK to abort the script.', 'No Beam Sets')
        sys.exit()
    patient_db = get_current('PatientDB')

    # Get unique DQA plan name while limiting to 16 characters
    qa_plan_names = [qa_plan.BeamSet.DicomPlanLabel for qa_plan in plan.VerificationPlans]
    qa_plan_name = plan.Name + ' DQA'
    qa_plan_name = unique_name(qa_plan_name, qa_plan_names)

    # Create QA plan
    iso = {'x': 21.93, 'y': 21.93, 'z': 0}  # DICOM coordinates of 'center' POI of Delta4 phantom. Wish we could get this directly from RayPhysics
    dg = plan.GetTotalDoseGrid().VoxelSize
    beam_set.CreateQAPlan(PhantomName='Delta4 Phantom', PhantomId='Delta4_2mm', QAPlanName=qa_plan_name, IsoCenter=iso, DoseGrid=dg, CouchRotationAngle=0, ComputeDoseWhenPlanIsCreated=True)

    # Set colorwash to CRMC standard (adapted for smaller dose)
    try:
        case.CaseSettings.DoseColorMap.ColorTable = patient_db.LoadTemplateColorMap(templateName=COLORWASH_TEMPLATE).ColorMap.ColorTable
    except:
        pass  # Not a big deal if we can't apply the color wash

    # Ensure phantom electronics are not excessively irradiated
    qa_plan = list(plan.VerificationPlans)[-1]  # Latest QA plan
    point = {'x': 21.93, 'y': 21.93, 'z': 20}  # InterpolateDoseInPoint takes DICOM coordinates
    dose_at_electronics = qa_plan.BeamSet.FractionDose.InterpolateDoseInPoint(Point=point, PointFrameOfReference=qa_plan.BeamSet.FrameOfReference)
    if dose_at_electronics > 20:
        res = MessageBox.Show(f'Phantom electronics may be irradiated at {dose_at_electronics:.1f} > 20 cGy. Export anyway?', 'Create QA Plan', MessageBoxButtons.YesNo)
        # Exit script if user does not want to continue
        if res == DialogResult.No:
            sys.exit()

    # Create export folder
    patient.Save()  # Must save before any DICOM export
    pt_name = format_pt_name(patient)  # e.g., 'Jones, Bill'
    qa_folder_name = pt_name + ' ' + qa_plan_name  # e.g., 'Jones, Bill Prostate DQA'
    qa_folder_name = os.path.join(EXPORT_PATH, qa_folder_name)  # Absolute path to export folder
    # Remove folder if it exists 
    if os.path.isdir(qa_folder_name):
        shutil.rmtree(qa_folder_name)
    os.makedirs(qa_folder_name)  # Create folder
    
    # Export QA plan
    export_args = {'ExportFolderPath': qa_folder_name, 'QaPlanIdentity': 'Patient', 'ExportBeamSet': True, 'ExportBeamSetDose': True, 'ExportBeamSetBeamDose': True, 'IgnorePreConditionWarnings': False}
    try:
        qa_plan.ScriptableQADicomExport(**export_args)
    except SystemError as e:
        res = MessageBox.Show('{}\nProceed?'.format(e), 'Create QA Plan', MessageBoxButtons.YesNo)
        if res == DialogResult.Yes:
            export_args['IgnorePreConditionWarnings'] = True
            qa_plan.ScriptableQADicomExport(**export_args)