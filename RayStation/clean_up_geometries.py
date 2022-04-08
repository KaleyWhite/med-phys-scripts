import clr
import sys

from connect import *

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


class CleanUpGeometriesForm(Form):
    def __init__(self, case, exam):
        self.Text = 'Clean Up Geometries'  # Form title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

        y = 15  # Vertical coordinate of next control

        if case.Examinations.Count > 1:
            self.tab_ctrl = TabControl()
            self.tab_ctrl.SizeMode = TabSizeMode.Fixed
            self.Controls.Add(self.tab_ctrl)

            for exam in case.Examinations:
                tab_pg = TabPage(exam.Name)
                self.tab_ctrl.Controls.Add(tab_pg)

                # Buttons for copying info among exams
                cp_btn_panel = Panel()
                tab_pg.Controls.Add(cp_btn_panel)
                cp_btn_panel.Dock = DockStyle.Top

                panel_x = 8

                btn = Button()
                btn.AutoSize = True
                btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
                cp_btn_panel.Controls.Add(btn)
                
                gb = GroupBox()
                gb.Location = Point(panel_x, 8)
                gb.BackColor = btn.BackColor
                cp_btn_panel.Controls.Add(gb)
                gb.Height = 50
                lbl = Label()
                lbl.BackColor = btn.BackColor
                lbl.Location = Point(8, 8)
                lbl.AutoSize = True
                lbl.AutoSizeMode = AutoSizeMode.GrowAndShrink
                gb.Controls.Add(lbl)
                
                if case.Examinations.Count == 2:
                    lbl.Text = 'Copy settings from ' + [e.Name for e in case.Examinations if not e.Equals(exam)][0]
                    gb.Width = lbl.Width + 16

                    btn.Text = 'Copy settings to ' + [e.Name for e in case.Examinations if not e.Equals(exam)][0]
                else:
                    lbl.Text = 'Copy settings from '
                    cb = ComboBox()
                    cb.BackColor = btn.BackColor
                    cb.Items.AddRange([e.Name for e in case.Examinations if not e.Equals(exam)])
                    cb.Location = Point(lbl.Width + 5, 8)
                    gb.Controls.Add(cb)
                    cb.Width = max(TextRenderer.MeasureText(e.Name, cb.Font).Width for e in case.Examinations if not e.Equals(exam))
                    gb.Width = lbl.Width + cb.Width + 16

                    btn.Text = 'Copy settings to all exams'

            btn.Location = Point(gb.Width + 8, 8)
            self.add_grid(tab_pg)

            width = max(TextRenderer.MeasureText(e.Name, self.tab_ctrl.Font).Width for e in case.Examinations)
            ht = TextRenderer.MeasureText(exam.Name, self.tab_ctrl.Font).Height
            self.tab_ctrl.ItemSize = Size(width, ht + 5)
            self.tab_ctrl.Size = Size((width + 1) * case.Examinations.Count, self.Height)
        else:
            self.add_grid(self)

    def add_grid(self, parent):
        data_grid_view = DataGridView()


def clean_up_geometries():
    try:
        case = get_current('Case')
        try:
            exam = get_current('Examination')
        except:
            MessageBox.Show('The current case has no exams. Click OK to abort the script.', 'No Exams')
            sys.exit()
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    form = CleanUpGeometriesForm(case, exam)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()