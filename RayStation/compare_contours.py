import clr
import sys
from typing import Dict

from connect import *

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


class MyLabel(Label):
    def __init__(self, parent, txt):
        super().__init__()

        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.Text = txt
        self.Location = Point(15, parent._y)
        parent._y += self.Height
        parent.Controls.Add(self)


class CompareContoursForm(Form):
    def __init__(self, case):
        self.copy_to: str = None
        self.copy_from: Dict[str, str] = {}

        self._case = case
        self._exam_names = self._exam_names_list()
        self._y = 15  # Vertical coordinate of next control

        self._style_form()
        
        self._to_exam_cb = ComboBox()
        self._set_up_to_exam_cb()
        print('cb set up')

        self._from_exams_dgv = DataGridView()
        self._set_up_from_exams_dgv()
        print('dgv set up')

        self._ok_btn = Button()
        self._set_up_ok_btn()
        print('btn set up')

    def _input_chged(self, sender, event):
        self.copy_to = self._to_exam_cb.SelectedItem
        self.copy_from = {}
        for row in self._from_exams_dgv.SelectedRows:
            row_idx = row.Index
            exam_name = self._from_exams_dgv[self._exam_col.Index, row_idx].Value
            suffix = self._from_exams_dgv[self._suffix_col.Index, row_idx].Value
            print(exam_name, suffix)
            self.copy_from[exam_name] = suffix
        lower_suffixes = [suffix.lower() for suffix in self.copy_from.values()]
        print(self.copy_to, self.copy_from)
        self._ok_btn.Enabled = self.copy_to is not None and self.copy_from and '' not in lower_suffixes and len(set(lower_suffixes)) == len(self.copy_from)

    def _ok_btn_Click(self, sender, event):
        self.DialogResult = DialogResult.OK

    def _exam_names_list(self):
        msg = ''
        if self._case.Examinations.Count == 0:
            msg = 'There are no exams in the current case.'
        elif self._case.Examinations.Count == 1:
            msg = 'There is only one exam in the current case, so there are no exams to copy from.'
        if msg:
            MessageBox.Show(msg, 'Not Enough Exams')
            sys.exit()

        return [exam.Name for exam in self._case.Examinations]
        
    def _set_up_to_exam_cb(self):
        MyLabel(self, 'Exam to copy to:')

        self.Controls.Add(self._to_exam_cb)
        self._to_exam_cb.AutoSize = True
        self._to_exam_cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self._to_exam_cb.DropDownStyle = ComboBoxStyle.DropDownList
        self._to_exam_cb.Items.AddRange(self._exam_names)
        self._to_exam_cb.Location = Point(15, self._y)
        self._to_exam_cb.SelectedValueChanged += self._input_chged
        self._y += self._to_exam_cb.Height + 15

    def _set_up_from_exams_dgv(self):
        MyLabel(self, 'Exam(s) to copy from:')

        # Basic settings
        self.Controls.Add(self._from_exams_dgv)
        self._from_exams_dgv.Location = Point(15, self._y)
        self._from_exams_dgv.ScrollBars = 0
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
        self._from_exams_dgv.ColumnHeadersDefaultCellStyle.Font = Font(self._from_exams_dgv.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
        self._from_exams_dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self._from_exams_dgv.EnableHeadersVisualStyles = False

        # Event handlers
        self._from_exams_dgv.SelectionChanged += self._input_chged
        self._from_exams_dgv.CellValueChanged += self._input_chged

        self._add_dgv_cols()
        self._populate_dgv()

    def _set_up_ok_btn(self):
        self.Controls.Add(self._ok_btn)
        self._ok_btn.AutoSize = True
        self._ok_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self._ok_btn.Click += self._ok_btn_Click
        self._ok_btn.Enabled = False
        self._ok_btn.Text = 'OK'
        self._ok_btn.Location = Point(15, self._y)

    def _add_dgv_cols(self):
        self._exam_col = DataGridViewTextBoxColumn()
        self._exam_col.HeaderText = 'Exam'
        self._exam_col.ReadOnly = True  # User cannot change values
        self._from_exams_dgv.Columns.Add(self._exam_col)

        self._suffix_col = DataGridViewTextBoxColumn()
        self._suffix_col.HeaderText = 'ROI Name Suffix'
        self._from_exams_dgv.Columns.Add(self._suffix_col)

    def _populate_dgv(self):
        for exam_name in self._exam_names:
            self._from_exams_dgv.Rows.Add([exam_name, ''])
        
        # Resize table to contents
        self._from_exams_dgv.Width = sum(col.Width for col in self._from_exams_dgv.Columns) + 2
        self._from_exams_dgv.Height = self._from_exams_dgv.Rows[0].Height * self._from_exams_dgv.Rows.Count
        self._y += self._from_exams_dgv.Height + 15

    def _style_form(self):
        self.Text = 'Compare Contours'  # Form title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen


def compare_contours():
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    form = CompareContoursForm(case)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
    to_exam_name = form.copy_to

    for from_exam_name, suffix in form.copy_from.items():
        from_exam = case.Examinations[from_exam_name]
        from_exam_roi_names = [geom.OfRoi.Name for geom in case.PatientModel.StructureSets[from_exam_name].RoiGeometries if geom.HasContours() and geom.OfRoi.Type not in ['External', 'Fixation', 'Support']]
        to_exam_roi_names = []
        if from_exam_roi_names:
            with CompositeAction('Copy from "' + from_exam_name + '"'):
                for from_exam_roi_name in from_exam_roi_names:
                    from_roi = case.PatientModel.RegionsOfInterest[from_exam_roi_name]
                    new_roi_name = case.PatientModel.GetUniqueRoiName(DesiredName=from_exam_roi_name + '^' + suffix)
                    new_roi = case.PatientModel.CreateRoi(Name=new_roi_name, Color=from_roi.Color, Type=from_roi.Type, TissueName=from_roi.OrganData.ResponseFunctionTissueName, RbeCellTypeName=from_roi.OrganData.RbeCellTypeName, RoiMaterial=from_roi.RoiMaterial)
                    new_roi.CreateMarginGeometry(Examination=from_exam, SourceRoiName=from_exam_roi_name, MarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                    to_exam_roi_names.append(new_roi_name)
                if from_exam_name != to_exam_name:
                    case.PatientModel.CopyRoiGeometries(SourceExamination=from_exam, TargetExaminationNames=[to_exam_name], RoiNames=to_exam_roi_names)
                    for to_exam_roi_name in to_exam_roi_names:
                        case.PatientModel.StructureSets[from_exam_name].RoiGeometries[to_exam_roi_name].DeleteGeometry()
