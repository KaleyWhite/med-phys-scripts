import clr
from collections import OrderedDict
import sys
from typing import List, Optional

from connect import *

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


class MyLabel(Label):
    """Class that defines a WinForms Label with certain common settings"""

    def __init__(self, parent: Control, txt: Optional[str] = '') -> None:
        """Initializes a MyLabel object.
        
        Arguments
        ---------
        parent: The object to add the Label to
        txt: The text of the Label
        """
        super().__init__()

        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.Text = txt
        self.Location = Point(15, parent._y)
        parent._y += self.Height
        parent.Controls.Add(self)


class CompareContoursForm(Form):
    """Form that allows the user to choose settings for copying contours between exams

    User chooses:
    - Exam to copy to
    - Exams to copy from, and suffix to use for each
    """
    def __init__(self, to_exam_name: Optional[str], from_exam_names: List[str]) -> None:
        """Initializes a CompareContoursForm object
        
        Arguments
        ---------
        to_exam_name: Name of the default selected exam to copy contours to. 
                      If None, no exam is selected by default.
        from_exam_names: List of exam names that the user may copy contours from.
        """

        # Name of exam to copy to
        self.copy_to = to_exam_name

        # {exam name : suffix} dict for exam names that can be copied from
        self.copy_from = OrderedDict([(from_exam_name, '') for from_exam_name in from_exam_names])

        # Vertical coordinate of next control
        self._y = 15

        self._style_form()
        
        # Control to select exam to copy to
        self._to_exam_cb = ComboBox()
        self._set_up_to_exam_cb()

        # Control to select exam(s) to copy from, and suffix for each
        self._from_exams_dgv = DataGridView()
        self._set_up_from_exams_dgv()

        # "OK" button to submit input
        self._ok_btn = Button()
        self._set_up_ok_btn()

    def _input_chged(self, sender, event):
        # Event handler for DGV CellValueChanged, DGV SelectionChanged, and ComboBox SelectedValueChanged
        # Sets self.copy_to and self.copy_from
        # Enables "OK" button only if self.copy_to is provided, self.copy_from is not empty, and a unique suffix is provided for each selected row in the DGV
        
        self.copy_to = self._to_exam_cb.SelectedItem

        # Recreate self.copy_from
        self.copy_from = OrderedDict()
        # Add exam name and suffix from each selected row, to self.copy_from
        for row in self._from_exams_dgv.SelectedRows:
            row_idx = row.Index
            exam_name = self._from_exams_dgv[self._exam_col.Index, row_idx].Value
            suffix = self._from_exams_dgv[self._suffix_col.Index, row_idx].Value
            self.copy_from[exam_name] = suffix
        
        # Convert all suffixes to lowercase for case-insensitive uniqueness check
        lower_suffixes = [suffix.lower() for suffix in self.copy_from.values()]

        # Enable or disable "OK" button
        self._ok_btn.Enabled = self.copy_to is not None and self.copy_from and '' not in lower_suffixes and len(set(lower_suffixes)) == len(self.copy_from)

    def _ok_btn_Click(self, sender, event):
        #Event handler for click of the "OK" button
        self.DialogResult = DialogResult.OK
        
    def _set_up_to_exam_cb(self):
        # Styles and adds the ComboBox for the exam to copy to
        
        MyLabel(self, 'Exam to copy to:')

        self.Controls.Add(self._to_exam_cb)

        # Autosize to contents
        self._to_exam_cb.AutoSize = True
        self._to_exam_cb.AutoSizeMode = AutoSizeMode.GrowAndShrink

        self._to_exam_cb.DropDownStyle = ComboBoxStyle.DropDownList  # User cannot type custom text
        self._to_exam_cb.Items.AddRange(self.copy_from)  # Populate
        self._to_exam_cb.Location = Point(15, self._y)

        # If self.copy_to is provided, select it
        if self.copy_to is not None:
            self._to_exam_cb.SelectedIndex = list(self.copy_from.keys()).index(self.copy_to)

        self._y += self._to_exam_cb.Height + 15

    def _set_up_from_exams_dgv(self):
        # Styles and adds the DataGridView for exams to copy from and their suffixes

        MyLabel(self, 'Exam(s) to copy from:')

        # Basic settings
        self.Controls.Add(self._from_exams_dgv)
        self._from_exams_dgv.Location = Point(15, self._y)
        self._from_exams_dgv.ScrollBars = 0  # Do not use scrollbars
        self._from_exams_dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self._from_exams_dgv.GridColor = Color.Black

        # Row settings
        self._from_exams_dgv.RowHeadersVisible = False
        self._from_exams_dgv.AllowUserToAddRows = False
        self._from_exams_dgv.AllowUserToDeleteRows = False
        self._from_exams_dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells

        # Column settings
        self._from_exams_dgv.AllowUserToOrderColumns = False
        self._from_exams_dgv.AllowUserToResizeColumns = False
        self._from_exams_dgv.AutoGenerateColumns = False
        self._from_exams_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        self._from_exams_dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        self._from_exams_dgv.ColumnHeadersDefaultCellStyle.Font = Font(self._from_exams_dgv.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)  # Bold column headers
        self._from_exams_dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self._from_exams_dgv.EnableHeadersVisualStyles = False

        # Event handlers
        self._from_exams_dgv.SelectionChanged += self._input_chged
        self._from_exams_dgv.CellValueChanged += self._input_chged
        self._to_exam_cb.SelectedValueChanged += self._input_chged

        self._add_dgv_cols()
        self._populate_dgv()

    def _set_up_ok_btn(self):
        # Styles and adds the "OK" button

        self.Controls.Add(self._ok_btn)

        # Autosize to text
        self._ok_btn.AutoSize = True
        self._ok_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink

        self._ok_btn.Click += self._ok_btn_Click
        self._ok_btn.Enabled = False  # By default no exams are selected to copy from, so "OK" button is disabled
        self._ok_btn.Text = 'OK'
        self._ok_btn.Location = Point(15, self._y)

    def _add_dgv_cols(self):
        # Adds "Exam" and "ROI Name Suffix" columns to the DGV

        # "Exam" column
        self._exam_col = DataGridViewTextBoxColumn()
        self._exam_col.HeaderText = 'Exam'
        self._exam_col.ReadOnly = True  # User cannot change values
        self._from_exams_dgv.Columns.Add(self._exam_col)

        # "ROI Name Suffix" column
        self._suffix_col = DataGridViewTextBoxColumn()
        self._suffix_col.HeaderText = 'ROI Name Suffix'
        self._from_exams_dgv.Columns.Add(self._suffix_col)

    def _populate_dgv(self):
        # Adds rows to the DGV

        # Add a row with an empty suffix for each exam
        for exam_name, suffix in self.copy_from.items():
            self._from_exams_dgv.Rows.Add([exam_name, suffix])

        # By default, no exams are selected to copy from
        self._from_exams_dgv.ClearSelection()
        
        # Resize table to contents
        self._from_exams_dgv.Width = sum(col.Width for col in self._from_exams_dgv.Columns) + 2
        self._from_exams_dgv.Height = self._from_exams_dgv.Rows[0].Height * (self._from_exams_dgv.Rows.Count + 1)
        
        self._y += self._from_exams_dgv.Height + 15

    def _style_form(self):
        # Styles the Form

        self.Text = 'Compare Contours'  # Form title
        
        # Autosize to contents
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)  # At least as wide as form title plus some room for "X" button, etc.

        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen


def compare_contours():
    """Copies contours to enable comparison of contours between exams.

    The user selects the exam to copy contours to and the exam(s) to copy from. Only exams with contours are available for selection.
    A new ROI is created for each nonempty contour on the exams to copy from. Each new ROI has a geometry only on the exam to copy to. The ROI name has the provided suffix appended to the original name according to TG-263.
    Ignores ROIs of type External, Support, or Fixation.
    """

    # Get current variables
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        to_exam = get_current('Examination')
    except:
        MessageBox.Show('There are no exams in the current case. Click OK to abort the script.', 'No Exams')
        sys.exit()

    # Create dictionary of {exam name : [contour names]} for all exams in current case
    from_exam_names: OrderedDict[str, List[str]] = OrderedDict()
    for from_exam in case.Examinations:
        from_exam_roi_names = [geom.OfRoi.Name for geom in case.PatientModel.StructureSets[from_exam.Name].RoiGeometries if geom.HasContours() and geom.OfRoi.Type not in ['External', 'Fixation', 'Support']]  # ROIs with contours that are not External, Support, or Fixation
        if from_exam_roi_names:
            from_exam_names[from_exam.Name] = from_exam_roi_names
    # If no exams with non-External, -Support, or -Fixation geometries, alert user and exit script
    if not from_exam_names:
        MessageBox.Show('There are no exams with contours that are not External, Support, or Fixation. Click OK to abort the script.', 'No Contours')
        sys.exit()

    # Default exam to copy to is current exam if it has geometries
    to_exam_name = to_exam.Name if to_exam.Name in from_exam_names else None

    # Get exam to copy to, exam(s) to copy from, and suffix(es) from user
    form = CompareContoursForm(to_exam_name, from_exam_names)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
    to_exam_name = form.copy_to
    from_exam_suffixes = form.copy_from

    # Copy geometries from each selected exam
    for from_exam_name, suffix in from_exam_suffixes.items():
        from_exam = case.Examinations[from_exam_name]  # Exam to copy from
        from_exam_roi_names = from_exam_names[from_exam_name]  # Names of geometries to copy from that exam
        to_exam_roi_names = []  # Names of geometries to delete from exam after they are copied
        with CompositeAction('Copy from "' + from_exam_name + '"'):
            for from_exam_roi_name in from_exam_roi_names:
                from_roi = case.PatientModel.RegionsOfInterest[from_exam_roi_name]
                new_roi_name = case.PatientModel.GetUniqueRoiName(DesiredName=from_exam_roi_name + '^' + suffix)  # Add suffix to ROI name (e.g., "Lung_L" with suffix "SBRT" -> "Lung_L^SBRT")
                new_roi = case.PatientModel.CreateRoi(Name=new_roi_name, Color=from_roi.Color, Type=from_roi.Type, TissueName=from_roi.OrganData.ResponseFunctionTissueName, RbeCellTypeName=from_roi.OrganData.RbeCellTypeName, RoiMaterial=from_roi.RoiMaterial)  # Create new ROI with same color, type, etc. as the ROi we are copying
                new_roi.CreateMarginGeometry(Examination=from_exam, SourceRoiName=from_exam_roi_name, MarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })  # Copy old geometry into new ROI
                to_exam_roi_names.append(new_roi_name)  # We will delete the geometry after copying it to the other exam
            # If we aren't copying to the same exam, copy geometries to new exam and delete geometries from the exam we copied from
            if from_exam_name != to_exam_name:
                case.PatientModel.CopyRoiGeometries(SourceExamination=from_exam, TargetExaminationNames=[to_exam_name], RoiNames=to_exam_roi_names)
                for to_exam_roi_name in to_exam_roi_names:
                    case.PatientModel.StructureSets[from_exam_name].RoiGeometries[to_exam_roi_name].DeleteGeometry()
