import os
import shutil
import re
import sys
from typing import List

import pydicom

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons


# Name of the colorwash template to apply
COLORWASH_TEMPLATE = 'CRMC Standard Dose Colorwash'

# Absolute path to the directory to export DICOM files to
EXPORT_PATH = os.path.join('T:', os.sep, 'Physics', 'QA & Procedures', 'Delta4', 'DQA Plans')

# List of machine names to export DQA plans for
MACHINES = ['E1', 'E2', 'ELEKTA']


def unique_qa_plan_name(plan: PyScriptObject) -> str:
    """Creates a DQA plan name unique in the plan

    Name is based on the plan name with " DQA" appended
    Name is made unique with a copy number in parentheses
    New name is truncated to be at most 16 characters long

    Arguments
    ---------
    eplan: The plan to create the DQA plan name for

    Returns
    -------
    The new DQA plan name, made unique

    Example
    -------
    Assuming some_plan named "SBRT Lung_L" has DQA plans "SBRT Lung_L DQA" and "SBRT Lun DQA (1)":
    unique_qa_plan_name(some_plan) -> "SBRT Lun DQA (1)"
    """
    new_name = plan.Name[:12] + ' DQA'  # Truncate to at most 16 characters, and " DQA" suffix is 4 characters
    existing_names = [qa_plan.BeamSet.DicomPlanLabel for qa_plan in plan.VerificationPlans]
    copy_num = 0  # Assume no copies
    # Increment the copy number until it makes the name unique
    while new_name in existing_names:
        copy_num += 1
        copy_str = ' (' + str(copy_num) + ')'  # Copy suffix to add the name to make it unique
        name_len = 12 - len(copy_str)  # Number of characters allowed before the " DQA" and copy suffixes
        new_name = plan.Name[:name_len] + ' DQA' + copy_str
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


def export_qa_plan(qa_plan: PyScriptObject, qa_folder_name: str, machine: str) -> None:
    """Exports the QA plan's RTDOSE and RTPLAN files to a subfolder of `qa_folder_name` called `machine`

    Arguments
    ---------
    qa_plan: The QA plan to export
    qa_folder_name: The path of the directory to export to, relative to `EXPORT_PATH`
                    Must already exist
    machine: The name of the subdirectory of `qa_folder_name` to export to
             Must not already exist
    """
    # Create subfolder to export to
    machine_folder_name = os.path.join(qa_folder_name, machine)
    os.mkdir(machine_folder_name)
    # Export
    export_args = {'ExportFolderPath': machine_folder_name, 'QaPlanIdentity': 'Patient', 'ExportBeamSet': True, 'ExportBeamSetDose': True, 'ExportBeamSetBeamDose': True, 'IgnorePreConditionWarnings': False}
    try:
        qa_plan.ScriptableQADicomExport(**export_args)
    except SystemError as e:
        res = MessageBox.Show(str(e) + '\nProceed?', 'Create QA Plan', MessageBoxButtons.YesNo)
        if res == DialogResult.Yes:
            export_args['IgnorePreConditionWarnings'] = True
            qa_plan.ScriptableQADicomExport(**export_args)


def copy_for_new_machine(qa_folder_name: str, old_machine: str, new_machine: str) -> None:
    """Copies the QA DICOM files from one folder to another and changes the machine name in the copied RTPLAN file

    Arguments
    ---------
    qa_folder_name: The path to the directory containing the `old_machine` directory, relative to `EXPORT_PATH`
    old_machine: The subdirectory to copy, relative to `qa_folder_name`
                 Must already exist
    new_machine: The machine to change the new DICOM files to use. Also the name of the new subdirectory
                 Must not already exist
    """
    src = os.path.join(EXPORT_PATH, qa_folder_name, old_machine)
    dst = os.path.join(EXPORT_PATH, qa_folder_name, new_machine)
    shutil.copytree(src, dst)

    change_machine(qa_folder_name, new_machine)


def change_machine(qa_folder_name: str, new_machine: str) -> None:
    """Changes the machine of the RTPLAN DICOM file in `qa_folder_name`

    Acoording to RayStation's export conventions, assumes the RTPLAN file starts with "RP" while all other files (RTDOSE) in the machine directory start with "RD"

    The following DICOM attributes are changed:
    - For each beam in the BeamSequence (300A, 00B0), TreatmentMachineName is set to the new machine name
    - The RayStation-specific Private Data element (4001, 1012) value is set to the binary representation of the new machine name

    Arguments
    ---------
    qa_folder_name: The path of the QA DICOM files, relative to `EXPORT_PATH`
    new_machine: The subdirectory of `qa_folder_name` that contains the RTPLAN file to change
                 Must already exist
    """
    path = os.path.join(EXPORT_PATH, qa_folder_name, new_machine)
    rp_filename = sorted(os.listdir(path))[-1]
    rp_filepath = os.path.join(path, rp_filename)
    dcm = pydicom.dcmread(rp_filepath)
    # Set machine name for each beam
    for bs in dcm.BeamSequence:
        bs.TreatmentMachineName = new_machine
        bs[0x4001, 0x1012].value = new_machine.encode('ascii')  # Convert to byte string
    dcm.save_as(rp_filepath, write_like_original=False)  # Overwrite original DICOM file


def create_qa_plan() -> None:
    """Creates and exports DQA plan for the current photon beam set
    
    If export folder with the computed name already exists, delete and recreate it
    
    Export folder contains subfolder for each machine name in `MACHINES`
    The DICOM attributes TreatmentMachineName and a RayStation-specific private field in each machine folder are set to the machine name

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

    # Ensure beam set modality is photons
    if beam_set.Modality != 'Photons':
        MessageBox.Show('This is not a photon beam set. Click OK to abort the script.', 'Unsupported Modality')
        sys.exit()

    # Get unique DQA plan name while limiting to 16 characters
    qa_plan_name = unique_qa_plan_name(plan)

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

    # Export once and copy for all machines
    beam_set_machine = beam_set.MachineReference.MachineName 
    if beam_set_machine in MACHINES:
        export_qa_plan(qa_plan, qa_folder_name, beam_set_machine)
        MACHINES.remove(beam_set_machine)
        for machine in MACHINES:
            copy_for_new_machine(qa_folder_name, beam_set_machine, machine)
    else:
        export_qa_plan(qa_plan, qa_folder_name, MACHINES[0])
        change_machine(qa_folder_name, MACHINES[0])
        for machine in MACHINES[1:]:
            copy_for_new_machine(qa_folder_name, MACHINES[0], machine)
