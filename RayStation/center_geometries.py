import clr
from collections import OrderedDict
import sys
from typing import Dict, List

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System import EventArgs  # For type hints
from System.Drawing import *
from System.Windows.Forms import *


class CenterGeometriesForm(Form):
    def __init__(self, roi_names: List[str]) -> None:
        """Initializes a CenterGeometriesForm object

        Arguments
        ---------
        roi_names: List of ROI names that the user can choose to center
        """
        self.center_info = OrderedDict([(roi_name, [False] * 3) for roi_name in roi_names])

        self._y = 15

        self._dgv = DataGridView()
        self._ok_btn = Button()

        self._set_up_form()
        self._set_up_dgv()
        self._populate_dgv()
        self._set_up_ok_btn()

    def _dgv_CellContentClick(self, sender: DataGridView, event: DataGridViewCellEventArgs) -> None:
        # Event handler for clicking a checkbox
        # This allows the cell value to actually change (the CellValueChanged event to be fired)
        sender.CommitEdit(DataGridViewDataErrorContexts.Commit)

    def _dgv_CellValueChanged(self, sender: DataGridView, event: DataGridViewCellEventArgs) -> None:
        # Event handler for cell value change in the DGV

        # Only look at the CheckBoxColumns
        if event.ColumnIndex > 0:
            roi_name = sender[0, event.RowIndex].Value
            val = sender[event.ColumnIndex, event.RowIndex].Value
            self.center_info[roi_name][event.ColumnIndex - 1] = val
            self._ok_btn.Enabled = self._ok_btn.Enabled or val
            event.Handled = True

    def _ok_btn_Click(self, sender: Button, event: EventArgs) -> None:
        self.DialogResult = DialogResult.OK

    def _set_up_form(self) -> None:
        # Styles the Form

        self.Text = 'Center Geometries'  # Form title

        # Adapt form size to controls
        # Make form at least as wide as title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink 
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)

        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

    def _set_up_dgv(self) -> None:
        # Style and add a DGV to the Form

        # Basic settings
        self.Controls.Add(self._dgv)
        self._dgv.Location = Point(15, self._y)
        self._dgv.ScrollBars = 0
        self._dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self._dgv.GridColor = Color.Black

        # Row settings
        self._dgv.RowHeadersVisible = False
        self._dgv.AllowUserToAddRows = False
        self._dgv.AllowUserToDeleteRows = False

        # Column settings
        self._dgv.AllowUserToOrderColumns = False
        self._dgv.AllowUserToResizeColumns = False
        self._dgv.AutoGenerateColumns = False
        self._dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        self._dgv.ColumnHeadersDefaultCellStyle.Font = Font(self._dgv.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
        self._dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = self._dgv.ColumnHeadersDefaultCellStyle.BackColor
        self._dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = self._dgv.ColumnHeadersDefaultCellStyle.ForeColor
        self._dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self._dgv.EnableHeadersVisualStyles = False
    
        # Cell settings
        self._dgv.DefaultCellStyle.SelectionBackColor = self._dgv.DefaultCellStyle.BackColor
        self._dgv.DefaultCellStyle.SelectionForeColor = self._dgv.DefaultCellStyle.ForeColor
        
        # Event handlers
        self._dgv.CellContentClick += self._dgv_CellContentClick
        self._dgv.CellValueChanged += self._dgv_CellValueChanged

    def _populate_dgv(self) -> None:
        # Adds the four columns to the table: "ROI Geometry", "R-L", "A-P", and "I-S", and populates rows

        # "ROI Geometry" column
        col = DataGridViewTextBoxColumn()
        col.HeaderText = 'ROI Geometry'
        col.ReadOnly = True  # User cannot change values
        col.Height = 25
        col.Width = 150
        self._dgv.Columns.Add(col)

        # CheckBox columns
        for dim in ['R-L', 'A-P', 'I-S']:
            col = DataGridViewCheckBoxColumn()
            col.HeaderText = dim
            col.Height = col.Width = 30
            self._dgv.Columns.Add(col)

        # Add a row for each ROI
        # By default, no checkboaxes are checked
        for roi_name, vals in self.center_info.items():
            self._dgv.Rows.Add([roi_name] + vals)

        # Resize table
        self._dgv.Width = 240
        self._dgv.Height = 25 * len(self.center_info)
        
        self._y += self._dgv.Height + 15

    def _set_up_ok_btn(self) -> None:
        # Styles and adds the "OK" button to the form

        # Autosize to contents
        self._ok_btn.AutoSize = True
        self._ok_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink

        self._ok_btn.Click += self._ok_btn_Click
        self._ok_btn.Enabled = False  # By default, no checkboxes are checked, so "OK" button should not be clickable
        self._ok_btn.Location = Point(15, self._y)
        self._ok_btn.Text = 'OK'

        self.Controls.Add(self._ok_btn)


def exam_ctr(exam: PyScriptObject) -> Dict[str, float]:
    exam_min, exam_max = exam.Series[0].ImageStack.GetBoundingBox()
    ctr = {}
    for dim, coord in exam_min.items():
        ctr[dim] = (coord + exam_max[dim]) / 2
    return ctr


def center_geometries() -> None:
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('No case is open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        exam = get_current('Examination')
    except:
        MessageBox.Show('There are no exams in the open case. Click OK to abort the script.', 'No Exams')
        sys.exit()

    # Ensure there are unapproved, non-empty ROI geometries on the current exam
    struct_set = case.PatientModel.StructureSets[exam.Name]
    approved_roi_names = list(set(geom.OfRoi.Name for approved_ss in struct_set.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures))
    roi_names = sorted(geom.OfRoi.Name for geom in struct_set.RoiGeometries if geom.OfRoi.Name not in approved_roi_names and geom.HasContours())
    if not roi_names:
        MessageBox.Show('There are no non-empty, unapproved ROI geometries on the current exam. Click OK to abort the script.', 'No Non-Empty, Unapproved Geometries')
        sys.exit()

    # Get user input from form
    form = CenterGeometriesForm(roi_names)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
    center_info = form.center_info

    # Exam center
    img_ctr = exam_ctr(exam)

    # Center the ROIs
    with CompositeAction('Center ROI Geometries'):
        for roi_name, info in center_info.items():
            geom = struct_set.RoiGeometries[roi_name]  # That ROI's geometry on the exam

            # Current position
            geom_ctr = geom.GetCenterOfRoi()

            # Transformation matrix
            # Each x, y, and z transform is the difference between the correct and current coordinates
            m14 = img_ctr['x'] - geom_ctr.x if info[0] else 0
            m24 = img_ctr['y'] - geom_ctr.y if info[1] else 0
            m34 = img_ctr['z'] - geom_ctr.z if info[2] else 0
            mat = {'M11': 1, 'M12': 0, 'M13': 0, 'M14': m14,
                   'M21': 0, 'M22': 1, 'M23': 0, 'M24': m24,
                   'M31': 0, 'M32': 0, 'M33': 1, 'M34': m34, 
                   'M41': 0, 'M42': 0, 'M43': 0, 'M44': 1}

            # Reposition geometry
            geom.OfRoi.TransformROI3D(Examination=exam, TransformationMatrix=mat)
