import clr
import sys

from connect import *

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')


class MyLabel(Label):
    def __init__(self, parent, txt):
        super().__init__()

        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.Location = Point(15, parent._y)
        parent.Controls.Add(self)


class SmoothContoursForm(Form):
    def __init__(self, roi_names):
        self.margin = None
        self.selected_roi_names = []

        self._y = 15

        self._set_up_form()

        self._margin_tb = TextBox()
        self._roi_lb = ListBox()
        self._ok_btn = Button()

        self._set_up_margin_tb()
        self._set_up_roi_lb(roi_names)
        self._set_up_ok_btn()

    def _margin_tb_TextChanged(self, sender, event):
        try:
            self.margin = float(sender.Text)
            sender.BackColor = Color.White
            self._ok_btn.Enabled = self._ok_btn.Enabled and True
        except ValueError:
            self.margin = None
            sender.BackColor = Color.Red
            self._ok_btn.Enabled = False

    def _roi_lb_SelectionChanged(self, sender, event):
        self.selected_roi_names = sender.SelectedItems
        self._ok_btn.Enabled = self._ok_btn.Enabled and bool(self.selected_roi_names)

    def _ok_btn_Click(self, sender, event):
        self.DialogResult = DialogResult.OK

    def _set_up_form(self):
        # Styles the Form
    
        self.Text = 'Smooth Contours'  # Form title

        # Adapt form size to controls
        # Make form at least as wide as title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink 
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

    def _set_up_margin_tb(self):
        l = MyLabel(self, 'Margin (cm):')
        self._margin_tb.TextChanged += self._margin_tb_TextChanged
        self._margin_tb.Location = Point(15 + l.Width, self._y)
        self._margin_tb.Text = '3'
        self._margin_tb.Width = 25
        self.Controls.Add(self._margin_tb)

    def _set_up_roi_lb(self, roi_names):
        l = MyLabel('Choose the ROI(s) to smooth.\nNeither approved nor empty ROIs are shown.')
        self._y += l.Height

        self._roi_lb.Items.AddRange(roi_names)
        self._roi_lb.Height = self._roi_lb.PreferredHeight
        self._roi_lb.Location = Point(15, self._y)
        self._roi_lb.Width = max(TextRenderer.MeasureText(roi_name, self._roi_lb.Font).Width for roi_name in roi_names)
        self._roi_lb.ClearSelected()
        self._roi_lb.SelectionChanged += self._roi_lb_SelectionChanged
        self.Controls.Add(self._roi_lb)
        self._y += self._roi_lb.Height

    def _set_up_ok_btn(self):
        self._ok_btn.AutoSize = True
        self._ok_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self._ok_btn.Click += self._ok_btn_Click
        self._ok_btn.Location = Point(15, self._y)
        self._ok_btn.Text = 'OK'
        self.Controls.Add(self._ok_btn)


def smooth_contours():
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('No case is open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        exam = get_current('Exam')
    except:
        MessageBox.Show('There are no exams in the open case. Click OK to abort the script.', 'No Exams')
        sys.exit()

    struct_set = case.PatientModel.StructureSets[exam.Name]
    approved_roi_names = list(set(geom.OfRoi.Name for approved_ss in struct_set.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures))
    roi_names = [geom.OfRoi.Name for geom in struct_set.RoiGeometries if geom.OfRoi.Name not in approved_roi_names and geom.HasContours()]
    if not roi_names:
        MessageBox.Show('There are no non-empty, unapproved ROI geometries on the current exam. Click OK to abort the script.', 'No Non-Empty, Unapproved Geometries')
        sys.exit()

    form = SmoothContoursForm(sorted(roi_names))
    if form.DialogResult != DialogResult.OK:
        sys.exit()
    margin = form.margin
    selected_roi_names = form.selected_roi_names

    with CompositeAction('Smooth Contours'):
        for selected_roi_name in selected_roi_names:
