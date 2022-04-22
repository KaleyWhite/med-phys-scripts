import clr
import sys

from connect import *

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


class MyButton(Button):
    def __init__(self, parent, txt=''):
        super().__init__()

        parent.Controls.Add(self)
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.Text = txt


class CleanUpGeometriesForm(Form):
    def __init__(self, exam_names):
        self._exam_names = exam_names
        self._set_up_form()

        y = 15  # Vertical coordinate of next control

        # If multiple exams in current case, add tabs and buttons to propagate settings among exams
        if len(self._exam_names) > 1:
            self._tab_ctrl = TabControl()
            self._set_up_tab_ctrl()

            
        #else:
            #self.add_grid(self)

    def _set_up_form(self):
        self.Text = 'Clean Up Geometries'  # Form title
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

    def _set_up_tab_ctrl(self):
        self.Controls.Add(self._tab_ctrl)
        self._tab_ctrl.SizeMode = TabSizeMode.Fixed

        tab_width = max(TextRenderer.MeasureText(exam_name, self._tab_ctrl.Font).Width for exam_name in self._exam_names)
        print(tab_width)
        tab_ht = TextRenderer.MeasureText(self._exam_names[0], self._tab_ctrl.Font).Height
        for exam_name in self._exam_names:
            tab_pg = TabPage(exam_name)
            self._tab_ctrl.Controls.Add(tab_pg)

            # Buttons for copying info among exams
            self._add_btns(tab_pg)

            #self.add_grid(tab_pg)

        self._tab_ctrl.ItemSize = Size(tab_width, tab_ht + 5)
        self._tab_ctrl.Width = tab_width * len(self._exam_names) + 16
        self.Width = self._tab_ctrl.Width + 15

    def _add_btns(self, tab_pg):
        other_exam_names = [exam_name for exam_name in self._exam_names if not exam_name != tab_pg.Text]

        # "Copy to all" button
        copy_to_btn = MyButton(tab_pg)

        if len(self._exam_names) == 2:
            copy_from_btn_txt = 'Copy settings from ' + other_exam_names[0]
            copy_from_btn = MyButton(tab_pg, copy_from_btn_txt)
            copy_from_btn.Location = Point(8, 8)
            
            copy_to_btn_x = copy_from_btn.Width
            copy_to_btn.Text = 'Copy settings to ' + other_exam_names[0]
        else:
            # "Copy from ___" GroupBox that behaves as a button
            copy_from_gb = GroupBox()
            copy_from_gb.Location = Point(8, 8)
            copy_from_gb.BackColor = copy_to_btn.BackColor
            tab_pg.Controls.Add(copy_from_gb)
            copy_from_gb.Height = 50
            lbl = Label()
            lbl.BackColor = copy_from_btn.BackColor
            lbl.Location = Point(8, 8)
            lbl.AutoSize = True
            lbl.AutoSizeMode = AutoSizeMode.GrowAndShrink
            copy_from_gb.Controls.Add(lbl)

            lbl.Text = 'Copy settings from '
            copy_from_cb = ComboBox()
            copy_from_cb.BackColor = copy_to_btn.BackColor
            copy_from_cb.Items.AddRange(other_exam_names)
            copy_from_cb.Location = Point(lbl.Width + 5, 8)
            copy_from_gb.Controls.Add(copy_from_cb)
            copy_from_cb.Width = max(TextRenderer.MeasureText(exam_name, copy_from_cb.Font).Width for exam_name in other_exam_names)
            copy_from_gb.Width = lbl.Width + copy_from_cb.Width + 16
            
            copy_to_btn_x = copy_from_gb.Width
            copy_to_btn.Text = 'Copy settings to all exams'

        copy_to_btn.Location = Point(copy_to_btn_x + 16, 8)
        print(copy_to_btn.Width)
        tab_pg.Width = copy_to_btn_x + copy_to_btn.Width + 300

    def add_grid(self, parent):
        data_grid_view = DataGridView()


def clean_up_geometries():
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        exam = get_current('Examination')
    except:
        MessageBox.Show('The current case has no exams. Click OK to abort the script.', 'No Exams')
        sys.exit()

    exam_names = [exam.Name] + [other_exam.Name for other_exam in case.Examinations if not other_exam.Equals(exam)]

    form = CleanUpGeometriesForm(exam_names)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
