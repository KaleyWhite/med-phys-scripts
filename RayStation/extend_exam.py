import clr
from datetime import datetime
import json
import os
from math import ceil
import re
import shutil
import sys
from typing import Dict, List, Optional, Union

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

from pydicom import dcmread

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System import EventArgs  # For type hints
from System.Drawing import *
from System.Windows.Forms import *


EXPORT_DIR = os.path.join('T:', os.sep, 'Physics', 'Temp', 'Extend Exam')  # Parent directory of the new directory that the exam will be exported to
TG263_PATH = os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-spreadsheets', 'TG-263 Nomenclature with CRMC Colors.xlsm')  # Absolute path to the "TG-263 Nomenclature with CRMC Colors" spreadsheet


# ---------------------------- DICOM ---------------------------- #

def compute_new_ids(case: PyScriptObject) -> List[str]:
    """Creates a unique DICOM StudyInstanceUID and SeriesInstanceUID for an examination
    
    The new ID is based on the "maximum" ID across all exams in the case
    This "maximum" is the largest integer value of an ID with the periods removed
    The new ID is the same as the "maximum" except with the integer value of the last section (after the last period) incremented by 1

    Arguments
    ---------
    case: The case containing the exam to generate a new ID for

    Returns
    -------
    The new, unique ID

    Example
    -------
    Assuming that some_case has two exams with study IDs "1.4.2.9", "1.5.3.1", respectively, and series IDs "7.18.9.1" and "9.0.1.13", respectively:
    
    "Maximum" study ID = max(1429, 1531) = 1531
    New study ID = '1.5.3.' + (1 + 1) = '1.5.3.2'

    "Maximum" series ID = max(71891, 90113) = 90113
    New series ID = '9.0.1.' + (13 + 1) = '9.0.1.14'

    compute_new_id(some_case) -> ['1.5.3.2', '9.0.1.14']
    """
    ids = []
    for id_type in ['Study', 'Series']:
        # IDs of all exams in the case
        all_ids = [e.GetAcquisitionDataFromDicom()[id_type + 'Module'][id_type + 'InstanceUID'] for e in case.Examinations]
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

        ids.append(new_id)
    
    return ids


def copy_dicom_files(export_path: str, exam: PyScriptObject, min_dist: float, curr_dist: float, sup: Optional[bool] = True) -> None:
    """Copies the inferior or superior slice's DICOM file at `export_path` according to the expansion (`dist`) needed
    
    Arguments
    ---------
    export_path: Absolute path to the exported DICOM files for this run of the script
    exam: The examination to be extended
    min_dist: The desired minimum distance from the target to the inf/sup edge
    curr_dist: The current distance from the target to the inf/sup edge
    sup: True if exam should be extended in superior direction, False for inferior
    
    Sets SOP Instance UID (also use for filename), Slice Location, z-coordinate of Image Position (Patient), and Instance Number in the copied slice DICOM files
    """
    # Info for computing slice UIDs for new filenames
    # In RS, slices are ordered inf to sup, smallest ID to largest
    # Slice ID example: '1.2.840.113704.1.111.2528.1583439123.410'
    if sup:
        slice_id = list(exam.Series[0].ImageStack.ImportedDicomSliceUIDs)[-1]  # Last slice in list
    else:  # Inf
        slice_id = list(exam.Series[0].ImageStack.ImportedDicomSliceUIDs)[0]  # First slice in list
    
    # Split slice ID into first part (used for all new slice IDs) and second part (incremented/decremented for each new slice ID)
    dot_idx = slice_id.rfind('.')
    part_1, part_2 = slice_id[:dot_idx], int(slice_id[(dot_idx + 1):])

    # Get DICOM data for top (for sup) or bottom (for inf) slice
    dcm_filepath = os.path.join(export_path, 'CT' + slice_id + '.dcm')
    dcm = dcmread(dcm_filepath)

    # Slice data
    slice_thickness = dcm.SliceThickness
    slice_loc = dcm.SliceLocation
    num_copies = ceil((min_dist - curr_dist) * 10 / slice_thickness)  # x 10 to convert to mm

    # `copy_instance_num` = InstanceNumber for next slice
    if sup:
        copy_instance_num = len(os.listdir(export_path))  # We will increment the largest slice ID (see above)
    else:
        copy_instance_num = num_copies + 1  # We will decrement the smallest slice ID (see above)

    # Make `num_copies` copies of top/bottom slice
    for _ in range(num_copies):
        if sup:
            part_2 += 1  # Last part of slice UID (filename) is 1 more than previous slice (filename)
            copy_instance_num += 1
            slice_loc += slice_thickness
        else:
            part_2 -= 1  # Last part of slice UID (filename) is 1 less than previous slice (filename)
            copy_instance_num -= 1
            slice_loc -= slice_thickness

        # Copy DICOM file to appropriate new filename
        instance_id = part_1 + '.' + str(part_2)
        copy_dcm_filepath = os.path.join(export_path, 'CT' + instance_id + '.dcm')
        shutil.copy2(dcm_filepath, copy_dcm_filepath)

        # Change instance data in new slice
        copy_dcm = dcmread(copy_dcm_filepath)
        copy_dcm.SOPInstanceUID = instance_id
        copy_dcm.SliceLocation = slice_loc
        copy_dcm.ImagePositionPatient[2] = slice_loc
        copy_dcm.InstanceNumber = copy_instance_num

        # Overwrite copied DICOM file
        copy_dcm.save_as(copy_dcm_filepath)

    # Renumber the old instances if we added slices to the beginning
    if not sup:
        copy_instance_num = num_copies + 1
        for f in os.listdir(export_path)[num_copies:]:  # Ignore the copies, as they already have the correct instance number
            f = os.path.join(export_path, f)
            dcm = dcmread(f)
            dcm.InstanceNumber = copy_instance_num
            dcm.save_as(f)
            copy_instance_num += 1


def exam_by_ids(case: PyScriptObject, study_id: str, series_id: str) -> Optional[PyScriptObject]:
    """Gets the exam in the case that has the given DICOM StudyInstanceUID and SeriesInstanceUID

    Arguments
    ---------
    case: The case to search the exams in
    study_id: The StudyInstanceUID DICOM attribute to search exams for
    series_id: The SeriesInstanceUID DICOM attribute to search exams for

    Returns
    -------
    The exam with the given StudyInstanceUID and SeriesInstanceUID, or None if no such exam exists
    """
    for e in case.Examinations:
        dcm = e.GetAcquisitionDataFromDicom()
        if dcm['StudyModule']['StudyInstanceUID'] == study_id and dcm['SeriesModule']['SeriesInstanceUID'] == series_id:
            return e

# ----------------------------- Exam name ---------------------------- #

def name_exam(case: PyScriptObject, exam: PyScriptObject, exam_name: str) -> str:
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


# --------------------------- Color of External ROI -------------------------- #

def crmc_color(roi_name: str) -> str:
    """Gets the CRMC conventional color for the provided ROI, according to the "TG-263 Nomenclature with CRMC Colors" spreadsheet

    If the ROI name is not present in the spreadsheet, returns purple

    Arguments
    ---------
    roi_name: The name of the ROI whose conventional color to return

    Returns
    -------
    The ROI color, in the format "A, R, G, B"
    """
    tg263 = pd.read_excel(TG263_PATH, sheet_name='Names & Colors', usecols=['TG-263 Primary Name', 'Color'])
    tg263.set_index('TG-263 Primary Name', drop=True, inplace=True)
    try:
        return tg263.loc[roi_name, 'Color'][1:-1]  # Remove parens
    except KeyError:
        return Color.Purple


# -------------------------- Display import/export results -------------------------- #

def handle_export_warnings(result: Dict[str, Union[str, List[str]]]) -> None:
    # Currently only handles successful export - "Export finished."
    warnings = json.loads(str(result))
    if warnings['Comment'] == 'Export finished.' and not warnings['Warnings'] and not warnings['Notifications']:
        return


def handle_import_warnings(result: str) -> None:
    """Handles errors and warnings generated by a RayStation Import.. function (e.g., ImportDataFromPath)

    Arguments
    ---------
    result: The return value of the import function
    """
    # Exit function if no warnings
    if result == '':
        return
    # Compose and display warning message
    msg = ''
    warnings = warnings.split('\n')
    for w in warnings:
        if 'failed ValidateVR validation' in w:
            filename = w.split(' ')[1]
            msg = 'The DICOM file "CT' + filename + '.dcm" has an invalid UID value.\nOpen the file in a DICOM reader and examine values of attributes whose VR is UI.\nPossible culprits include Study Instance UID and Series Instance UID.'
    MessageBox.Show(msg, 'Import Warnings')

# ------------------------------ # decimal places ------------------------------ #

def num_dec_places(num: float) -> int:
    """Gets the number of digits after the decimal place, ignoring trailing zeroes

    Arguments
    ---------
    num: The number whose decimal places to count

    Returns
    -------
    The number of decimal places in the number
    """
    str_num = str(num).rstrip('0')
    if '.' not in str_num:
        return 0
    dec = str_num.split('.')[1]
    return len(dec)


# ---------------------------- GUI for user input ---------------------------- #

class ExtendExamForm(Form):
    """Class that gets a number from the user

    The number will be used as the minimum distance from the target to the sup or inf edge of the exam
    """
    DEFAULT_MIN_DIST = 5

    def __init__(self):
        """Initializes an ExtendExamForm object"""
        self._y = 15  # Vertical coordinate of next control

        self.min_dist = self.DEFAULT_MIN_DIST

        self._min_dist_tb = TextBox()
        self._ok_btn = Button()

        self._set_up_form()
        self._set_up_min_dist_tb()
        self._set_up_ok_btn()

    # ------------------------------ Event handlers ------------------------------ #

    def _min_dist_tb_TextChanged(self, sender: TextBox, event: EventArgs) -> None:
        """Handles the TextChanged event for the textbox

        If the value is valid (numeric, 0 <= value <= 10, and <=1 decimal places), set self.min_dist to the value, turn textbox background white, and enable "OK" button
        If invalid, set self_min_dist to None; turn textbox background red if nonempty, white otherwise; and disable "OK" button
        """
        try:
            min_dist = float(sender.Text)
            if min_dist < 0 or min_dist > 10 or num_dec_places(min_dist) > 1:  # Number out of range
                raise ValueError
            self.min_dist = min_dist
            sender.BackColor = Color.White
            self._ok_btn.Enabled = True
        except ValueError:  # Invalid value
            self.min_dist = None
            sender.BackColor = Color.White if sender.Text == '' else Color.Red
            self._ok_btn.Enabled = False

    def _ok_btn_Click(self, sender, event):
        """Handles the Click event for the "OK" button"""
        self.DialogResult = DialogResult.OK

    # ------------------------------- Add controls ------------------------------- #

    def _add_lbl(self, txt):
        """Adds a Label to the Form

        Arguments
        ---------
        txt: The text for the Label
        """
        l = Label()
        l.AutoSize = True
        l.AutoSizeMode = AutoSizeMode.GrowAndShrink
        l.Text = txt
        self.Controls.Add(l)
        return l

    # ------------------------------- Setup/layout ------------------------------- #

    def _set_up_form(self):
        """Styles the Form"""
        self.Text = 'Extend Exam'  # Form title

        # Adapt form size to controls
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

    def _set_up_min_dist_tb(self):
        """Styles and adds the textbox to the Form"""
        # Instructions for input
        instrs_lbl = self._add_lbl('Enter the minimum distance, in cm, from the target to the inferior or superior edge of the exam.\nValid values are between zero and 10, inclusive, with up to 1 decimal place.')
        instrs_lbl.Location = Point(15, 15)
        self._y += instrs_lbl.Height + 5
        x = instrs_lbl.Width // 2

        # Textbox to hold min_dist input
        self._min_dist_tb.TextChanged += self._min_dist_tb_TextChanged
        self._min_dist_tb.Location = Point(x, self._y)
        self._min_dist_tb.Text = str(self.min_dist)
        self._min_dist_tb.Width = TextRenderer.MeasureText('____', self._min_dist_tb.Font).Width  # Room for four characters: DD.D
        self.Controls.Add(self._min_dist_tb)

        # Unit
        unit_lbl = self._add_lbl('cm')
        unit_lbl.Font = Font(unit_lbl.Font.Name, unit_lbl.Font.Size + 2, FontStyle.Bold, unit_lbl.Font.Unit)
        unit_lbl.Location = Point(x + self._min_dist_tb.Width + 5, self._y)

        self._y += max(self._min_dist_tb.Height, unit_lbl.Height) + 15

    def _set_up_ok_btn(self):
        """Styles and adds the "OK" button to the Form"""
        self.AcceptButton = self._ok_btn

        # Adapt size to contents
        self._ok_btn.AutoSize = True
        self._ok_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink

        self._ok_btn.Click += self._ok_btn_Click
        self._ok_btn.Location = Point(self.ClientSize.Width // 2, self._y)
        self._ok_btn.Text = 'OK'
        self.Controls.Add(self._ok_btn)


# ------------------------------- Main function ------------------------------ #

def extend_exam() -> None:
    """Extends the current exam so that the target is at least the user-specified distance from the superior and inferior edges of the exam

    The user provides the minimum distance in a GUI
    If a beam set is loaded and its primary Rx is to a non-empty target, uses that target. If not, uses the first non-empty PTV, GTV, or CTV (checked in that order) found on the exam
    Exports exam, copies the top and/or bottom slice(s) so that target is far enough from the edges, and reimports as a new exam
    New exam name is old exam name plus ' - Expanded' (possibly with a copy number - e.g., 'SBRT Lung_R - Extended (1)')
    Copies ROI and POI geometries from old exam to new exam
    Does not copy any plans to new exam
    Deletes the temporary export directory
    """
    # Get current objects
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
        exam = get_current('Examination')
    except:
        MessageBox.Show('There are no exams in the current case. Click OK to abort the script.', 'No Exams')
        sys.exit()
    patient_db = get_current('PatientDB')
    
    struct_set = case.PatientModel.StructureSets[exam.Name]

    # Find the target
    target = None  # Assume no targets on exam
    try:
        rx_struct = get_current('BeamSet').PrimaryDosePrescriptionReference.OnStructure
        if rx_struct.OrganData.OrganType == 'Target' and struct_set.RoiGeometries[rx_struct.Name].HasContours():
            target = struct_set.RoiGeometries[rx_struct.Name]
    except:
        # Find first PTV, GTV, or CTV (in that order) contoured on the exam
        for target_type in ['PTV', 'GTV', 'CTV']:
            try:
                target = next(geom for geom in struct_set.RoiGeometries if geom.OfRoi.Type.upper() == target_type and geom.HasContours())
            except StopIteration:
                continue
    
    # If no targets contoured on exam, alert user and exit script with an error
    if target is None:
        MessageBox.Show('There are no target geometries on the current exam. Click OK to abort the script.', 'No Target Geometries')
        sys.exit()

    # Get min distance from user
    form = ExtendExamForm()
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
    min_dist = form.min_dist

    # Exam bounds
    img_bounds = exam.Series[0].ImageStack.GetBoundingBox()
    img_inf, img_sup = img_bounds[0].z, img_bounds[1].z

    # Target bounds
    target_bounds = target.GetBoundingBox()
    target_inf, target_sup = target_bounds[0].z, target_bounds[1].z

    # Distance from target to inf and sup exam edges
    inf_dist = abs(img_inf - target_inf)
    sup_dist = abs(img_sup - target_sup)

    # Exit script if target is at least `min_dist` cm from top or bottom of exam
    if inf_dist >= min_dist and sup_dist >= min_dist:
        MessageBox.Show(f'The target ({target.OfRoi.Name}) is {inf_dist:.1f} cm from the inferior edge, and {sup_dist:.1f} from the superior edge of the exam. These are both within your tolerance of {min_dist} cm, so no action is necessary.', 'Exam OK')
        sys.exit()

    # Create export directory
    export_path = os.path.join(EXPORT_DIR, datetime.now().strftime('%m-%d-%Y %H_%M_%S'))
    os.makedirs(export_path)
    
    # Export exam
    # Note that we could also export the structure set, 
    # but there is no way to access a structure set's UID from RS, 
    # and exporting every structure set so we could get the UID from the DICOM would unnecessary slow down the script.
    # Instead, after importing the new exam (later), we simply copy all ROI and POI geometries from the old exam to the new exam
    patient.Save()  # Error if you attempt to export when there are unsaved modifications
    export_args = {'ExportFolderPath': export_path, 'Examinations': [exam.Name], 'IgnorePreConditionWarnings': False}
    try:
        res = case.ScriptableDicomExport(**export_args)
        print(res)
        handle_export_warnings(res)
    except System.InvalidOperationException as error:
        print(error)
        export_args['IgnorePreConditionWarnings'] = True
        res = case.ScriptableDicomExport(**export_args)
        print(res)

    # Compute new study and series IDs so RS doesn't think the new exam is the same as the old
    study_id, series_id = compute_new_ids(case)

    # Add slices to top, if necessary
    if sup_dist < min_dist:
        copy_dicom_files(export_path, exam, min_dist, sup_dist)
    
    # Add slices to top, if necessary
    if inf_dist < min_dist:
        copy_dicom_files(export_path, exam, min_dist, inf_dist, False)

    # Change study and series UIDs in all files so RS doesn't think the new exam is the same as the old
    for f in os.listdir(export_path):
        f = os.path.join(export_path, f)  # Absolute path
        dcm = dcmread(f)
        dcm.StudyInstanceUID = study_id
        dcm.SeriesInstanceUID = series_id
        dcm.save_as(f)

    # Import new exam
    study = patient_db.QueryStudiesFromPath(Path=export_path, SearchCriterias={'PatientID': patient.PatientID})[0]  # There is only one study in the directory
    series = patient_db.QuerySeriesFromPath(Path=export_path, SearchCriterias=study)  # Series belonging to the study
    res = patient.ImportDataFromPath(Path=export_path, CaseName=case.CaseName, SeriesOrInstances=series)
    print(res)
    
    # Select new exam
    new_exam = exam_by_ids(case, study_id, series_id)
    new_exam.Name = name_exam(case, new_exam, exam.Name + ' - Extended')

    # Set new exam imaging system
    if exam.EquipmentInfo.ImagingSystemReference is not None:
        new_exam.EquipmentInfo.SetImagingSystemReference(ImagingSystemName=exam.EquipmentInfo.ImagingSystemReference.ImagingSystemName)

    # Add external geometry to new exam
    try:
        ext = next(roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == 'External')
    except StopIteration:  # No external in the case
        ext_name = case.PatientModel.GetUniqueRoiName(DesiredName='External')
        color = crmc_color('External')
        ext = case.PatientModel.CreateRoi(Name=ext_name, Color=color, Type='External')
    ext.CreateExternalGeometry(Examination=new_exam)

    # Copy ROI geometries from old exam to new exam
    geom_names = [geom.OfRoi.Name for geom in struct_set.RoiGeometries if geom.HasContours() and geom.OfRoi.Type != 'External']
    if geom_names:
        case.PatientModel.CopyRoiGeometries(SourceExamination=exam, TargetExaminationNames=[new_exam.Name], RoiNames=geom_names)

    # Update derived geometries (this shouldn't change any geometries since the new exam is effectively the same as the old)
    for geom in case.PatientModel.StructureSets[exam.Name].RoiGeometries:
        roi = case.PatientModel.RegionsOfInterest[geom.OfRoi.Name]
        if geom.OfRoi.DerivedRoiExpression is not None and geom.PrimaryShape.DerivedRoiStatus is not None and not geom.PrimaryShape.DerivedRoiStatus.IsShapeDirty:
            roi.UpdateDerivedGeometry(Examination=new_exam)

    # Copy POI geometries from old exam to new exam
    for i, poi in enumerate(struct_set.PoiGeometries):
        if poi.Point is not None and abs(poi.Point.x) != float('inf'):  # Empty POI geometry if infinite coordinates
            case.PatientModel.StructureSets[new_exam.Name].PoiGeometries[i].Point = poi.Point

    # Copy level/window presets
    new_exam.Series[0].LevelWindow = exam.Series[0].LevelWindow

    # Delete the temporary directory and all its contents
    shutil.rmtree(export_path)
