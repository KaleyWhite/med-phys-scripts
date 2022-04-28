import clr
from datetime import datetime
import os
import re
import shutil
import sys

from connect import *
from connect.connect_cpython import PyScriptObject

import pydicom

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # For displaying error messages


# Absolute path to the directory where the temporary directory to hold the exported DICOm files will be created
EXPORT_PATH_PARENT = os.path.join('T:', os.sep, 'Physics', 'Temp', 'copy_exam')


def timestamp() -> str:
    """Creates a timestamp of the current time, in the format MM-DD-YYYY HH_MM_SS

    The format uses underscores instead of colons so that the timestamp can be part of a valid Windows filename

    Returns
    -------
    The timestamp string

    Example
    -------
    timestamp() -> "04-28-2022 12_11_00"
    """
    dt = datetime.now()
    return dt.strftime('%M-%D-%YYYY %H_%M_%S')


def compute_new_id(case: PyScriptObject, study_or_series: str) -> str:
    """Creates a unique DICOM StudyInstanceUID or SeriesInstanceUID for an examination
    
    The new ID is based on the "maximum" ID across all exams in the case
    This "maximum" is the largest integer value of an ID with the periods removed
    The new ID is the same as the "maximum" except with the integer value of the last section (after the last period) incremented by 1

    Arguments
    ---------
    case: The case containing the exam to generate a new ID for
    study_or_series: The type of ID to generate
                     Must be either 'Study' or 'Series'

    Returns
    -------
    The new, unique ID

    Example
    -------
    Assuming that some_case has two exams with study IDs "1.4.2.9", "1.5.3.1", respectively:
    "Maximum" ID = max(1429, 1531) = 1531
    New ID = "1.5.3." + (1 + 1) = "1.5.3.2"

    compute_new_id(some_case, 'Study') -> "1.2.3.4"
    """
    # IDs of all exams in the case
    all_ids = [e.GetAcquisitionDataFromDicom()[study_or_series + 'Module'][study_or_series + 'InstanceUID'] for e in case.Examinations]
    all_ids.sort(key=lambda id_: int(id_.split('.')[-1]))
    
    # ID to base the new ID on
    last_id = all_ids[-1]

    # Extract last portion of ID, and everything else
    dot_idx = last_id.rfind('.')
    after_dot = last_id[(dot_idx + 1):]
    before_dot = last_id[-1][:dot_idx]
    
    # Construct new ID
    # First part is same. Last part is incremented by 1
    new_id = before_dot + '.' + str(int(after_dot) + 1)
    
    return new_id


def name_exam(case, exam, exam_name):
    """Generates an exam name unique among all exams in the case

    Name is made unique with a copy number in parentheses
    Name comparison is case insensitive

    Arguments
    ---------
    case: The case containing the exam to generate a name for
    exam: The exam to generate a name for
    exam_name: The desired name of the new exam

    Returns
    -------
    The desired new exam name, made unique

    Example
    -------
    Assuming some_case contains exams "Exam A", "Exam A (2)", and "An exam", to name a some_exam "exam A":
    name_exam(some_case, some_exam, "Exam A") -> "Exam A (3)"
    """
    regex = re.escape(exam_name) + r' \(([1-9]\d*)\)'  # Match the name plus a copy number
    copy_num = 0  # Assume no copies
    # Check all exam names in the case
    for e in case.Examinations:
        # Ignore the exam itself
        if e.Equals(exam):
            continue
        # Exact name match, without a copy number
        if e.Name.lower() == exam_name.lower():
            copy_num = 1
        else:
            m = re.match(regex, e.Name, re.IGNORECASE)  # Case insensitive
            if m is not None:  # Names match, with a copy number
                m_copy_num = int(m.group(1)) + 1  # The copy number in the matched name, plus 1
                copy_num = max(copy_num, m_copy_num)
    # No matches found -> no copy number necessary to make the name unique
    if copy_num == 0:
        return exam_name
    # Add copy number
    return exam_name + ' (' + str(copy_num) + ')'


def copy_exam():
    """Copies the current exam, including all structure sets on that exam

    Note that new exam "Used for" is "Evaluation" regardless of copied exam "Used for".

    New exam name is old exam name plus " - Copy", made unique with a copy number (e.g., "Breast 1/1/21 - Copy (1)")
    New exam has a new study UID and series ID

    Also copies level/window presets
    """
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('No patient is open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('No case is open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        old_exam = get_current('Examination')
    except:
        MessageBox.Show('There are no exams in the open case. Click OK to abort the script.', 'No Exams')
        sys.exit()
    patient_db = get_current('PatientDB')

    old_struct_set = case.PatientModel.StructureSets[old_exam.Name]

    # Create export directory
    export_path = os.path.join(EXPORT_PATH_PARENT, timestamp())
    os.makedirs(export_path)

    # Export exam and beam sets
    patient.Save()  # Error if you attempt to export when there are unsaved modifications
    export_args = {'ExportFolderPath': export_path, 'Examinations': [old_exam.Name], 'IgnorePreConditionWarnings': False}
    try:
        case.ScriptableDicomExport(**export_args)
    except:
        export_args['IgnorePreConditionWarnings'] = True
        case.ScriptableDicomExport(**export_args)  # Retry the export, ignoring warnings

    # Compute new study and series IDs
    new_study_id = compute_new_id(case, 'Study')
    new_series_id = compute_new_id(case, 'Series')

    # Change study and series IDs in all exported DICOM files
    for f in os.listdir(export_path):
        f = os.path.join(export_path, f)  # Absolute path to file
        dcm = pydicom.dcmread(f)  # Read DICOM data
        dcm.StudyInstanceUID = new_study_id
        dcm.SeriesInstanceUID = new_series_id
        dcm.save_as(f, write_like_original=False)  # Overwrite original DICOM file

    # Find and import the edited DICOM files
    try:
        study = patient_db.QueryStudiesFromPath(Path=export_path, SearchCriterias={'PatientID': patient.PatientID})[0]  # There is only one study in the directory
    except IndexError:
        MessageBox.Show('Could not find an exported exam with the correct StudyInstanceUID. Examine DICOM in "' + export_path + '" to troubleshoot. The new ID should be ' + new_study_id + '. Click OK to abort the script.', 'Series Not Found')
        sys.exit()
    series = patient_db.QuerySeriesFromPath(Path=export_path, SearchCriterias=study)  # Series belonging to the study
    patient.ImportDataFromPath(Path=export_path, SeriesOrInstances=series, CaseName=case.CaseName)  # Import into current case

    # The exam that was just imported
    try:
        new_exam = next(e for e in case.Examinations if e.Series[0].ImportedDicomUID == new_series_id)
    except StopIteration:
        MessageBox.Show('Could not find an exported exam with the correct SeriesInstanceUID. Examine DICOM in "' + export_path + '" to troubleshoot. The new ID should be ' + new_series_id + '. Click OK to abort the script.', 'Series Not Found')
        sys.exit()
    
    # Rename new exam
    new_exam_name = old_exam.Name + ' - Copy'
    new_exam.Name = name_exam(case, new_exam, new_exam_name)
    
    new_struct_set = case.PatientModel.StructureSets[new_exam.Name]

    # Set new exam imaging system
    old_img_sys = old_exam.EquipmentInfo.ImagingSystemReference.ImagingSystemName
    if old_img_sys is not None:
        new_exam.EquipmentInfo.SetImagingSystemReference(ImagingSystemName=old_img_sys)

    # Copy (non-empty) ROI geometries from old exam to new exam
    geom_names = [geom.OfRoi.Name for geom in old_struct_set.RoiGeometries if geom.HasContours()]
    if geom_names:
        case.PatientModel.CopyRoiGeometries(SourceExamination=old_exam, TargetExaminationNames=[new_exam.Name], RoiNames=geom_names)

    # Update derived geometries
    # This shouldn't change any geometries since the new exam is the same as the old
    # Ignore:
    # - Empty geometries
    # - Underived geometries (obviously)
    # - Dirty derived ROI geometries - these are not updated on old exam
    for geom in old_struct_set.RoiGeometries:
        roi = case.PatientModel.RegionsOfInterest[geom.OfRoi.Name]
        if geom.HasContours() and geom.OfRoi.DerivedRoiExpression is not None and geom.PrimaryShape.DerivedRoiStatus is not None and not geom.PrimaryShape.DerivedRoiStatus.IsShapeDirty:
            roi.UpdateDerivedGeometry(Examination=new_exam)

    # Copy POI geometries from old exam to new exam
    # Ignore empty geometries
    for i, poi in enumerate(old_struct_set.PoiGeometries):
        if poi.Point is not None and abs(poi.Point.x) < 1000:
            new_struct_set.PoiGeometries[i].Point = poi.Point

    # Copy level/window presets
    new_exam.Series[0].LevelWindow = old_exam.Series[0].LevelWindow

    # Delete the temporary directory and all its contents
    shutil.rmtree(export_path)
