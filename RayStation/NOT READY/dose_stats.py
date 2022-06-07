# For GUI
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import re  # For regex
import sys  # For sys.exit
from os import system

from connect import *  # Interact w/ RS

# ReportLab is used for PDF report generation
from reportlab.lib.colors import color2bw, obj_R_G_B, toColor, black, grey, white
from reportlab.lib.enums import TA_CENTER  # Paragraph alignment
from reportlab.lib.pagesizes import letter  # 8.5 x 11"
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle  # Custom ParagraphStyle objects are based on sample stylesheet
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth  # For determining table column width
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, KeepTogether
from reportlab.platypus.tables import Table, TableStyle

# For GUI
from System.Drawing import *
from System.Windows.Forms import *


case = None


## ReportLab Paragraph styles
styles = getSampleStyleSheet()  # Base styles (e.g., "Heading1", "Normal")
hdg_style = ParagraphStyle(name="hdg_style", fontName="Helvetica-Bold", fontSize=16, alignment=TA_CENTER)  # For patient name
subhdg_style = ParagraphStyle(name="subhdg_style", parent=styles["Normal"], fontName="Helvetica", fontSize=16, alignment=TA_CENTER)  # For MR# and "Dose Statistics"
tbl_hdg_style = ParagraphStyle(name="tbl_hdg_style", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, alignment=TA_CENTER)  # For center-aligned table headings
tbl_data_style = ParagraphStyle(name="tbl_data_style", parent=styles["Normal"], fontName="Helvetica", fontSize=10, leading=15, alignment=TA_CENTER)  # For center-aligned table data; leading necessary due to subscripts


# ReportLab Spacer objects to reuse for nice formatting
width, _ = letter  # Need the width (8.5") for Spacer width
spcr_sm = Spacer(width, 0.1 * inch)  # Small
spcr_md = Spacer(width, 0.2 * inch)  # Medium
spcr_lg = Spacer(width, 0.3 * inch)  # Large


def format_num(num, num_dec_places=0):
    # Helper function that returns a string version of the number rounded to the given number of decimal places
    # If `num_dec_places` is zero, return an int
    # Return "<<cutoff>" if the number rounds to zero
    # Assume both `num` and `num_dec_places` are non negative

    if num == 0:
        return "0"

    num = round(float(num), num_dec_places)  # Round to given number of places (convert to float just in case `num` is a str)
    
    # If number < cutoff, it rounds to zero
    if num == 0:  # Number rounds to 0
        return "<{}".format(10 ** -num_dec_places / 2.0)  # e.g., If rounding to 2 decimal places, cutoff = 0.0005
    
    # Rounding to zero places -> return an int
    if num_dec_places == 0:
        return int(num)

    # Remove leading zeroes
    num = str(num).lstrip("0") 
    
    # Remove trailing zeroes
    if "." in num:
        num = num.rstrip("0")
    
    # Prepend with zero
    if num.startswith("."):
        num = "0{}".format(num)

    # Remove extraneous decimal point
    if num.endswith("."):
        num = num[:-1]

    return num


def get_text_color(bk_color):
    # Helper function that returns the appropriate text color based on the background color
    
    return black if color2bw(bk_color) == white else white


def html_txt(txt):
    # Helper function that adds <sub> tags to the appropriate dose or volume text
    # Add a space between number and unit, if necessary

    regex = "([DV])(.+)(%|cc|Gy)"
    match = re.match(regex, txt)
    if not match:
        return txt
    if match.group(3) == "%":
        return re.subn(regex, "\g<1><sub>\g<2>\g<3></sub>", txt)[0]  # No space between number and "%"
    else:
        return re.subn(regex, "\g<1><sub>\g<2> \g<3></sub>", txt)[0]  # Space between number and "cc" or "Gy"


def sorting_key(item):
    # Helper function that returns the sorting key for a given valid dose or volume specification (e.g., "D0.035cc, "V20Gy")
    # Absolutes come before relatives. Within each group, sort the values in descending order

    val = item[1:-1] if item.endswith("%") else item[1:-2]
    return (-int(item.endswith("%")), -float(val))


def string_width(paragraph):
    # Helper function that measures the width of the text of a Paragraph, accounting for the smaller font used in subscripts
    # Estimate that subscript font is 3/4 the base font size

    font_name = paragraph.style.fontName
    base_font_sz = font_sz = paragraph.style.fontSize
    str_width = 0
    for txt in paragraph.text.split("<br/>"):
        width = 0
        i = 0
        while i < len(txt):
            if txt[i:(i + 5)] == "<sub>":
                font_sz = 0.75 * base_font_sz
                i += 5
            elif txt[i:(i + 6)] == "</sub>":
                font_sz = base_font_sz
                i += 6
            else:
                width += stringWidth(txt[i], fontName=font_name, fontSize=font_sz)
                i += 1
        str_width = max(width, str_width)
    return str_width


class DoseStatisticsForm(Form):
    """Form that allows user to select settings for dose statistics report:
        - ROIs
        - Plans
        - Doses on additional set
        - Doses at volume
        - Volumes at dose
    """

    def __init__(self):
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = "Dose Statistics"
        self.y = 15  # Vertical coordinate of next control

        # ROIs
        # Default: all ROIs are checked
        self.rois_gb = GroupBox()
        self.rois_gb.AutoSize = True
        self.rois_gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.rois_gb.Location = Point(15, self.y)
        self.rois_gb.Text = "ROI(s):"
        cb_y = 20
        cb = CheckBox()
        cb.AutoSize = True
        cb.IsThreeState = True
        cb.CheckState = cb.Tag = CheckState.Checked
        cb.Click += self.select_all_clicked
        cb.Location = Point(15, cb_y)
        cb.Text = "Select all"
        self.rois_gb.Controls.Add(cb)
        cb_y += cb.Height

        cb_x = 30
        base_cb_y = cb_y
        roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
        num_rows = int(round(len(roi_names) / 2.0))

        for i, roi_name in enumerate(roi_names):
            if i == num_rows:
                cb_x = 85 + max(TextRenderer.MeasureText(name, cb.Font).Width for name in roi_names[:num_rows])
                cb_y = base_cb_y
            cb = CheckBox()
            cb.AutoSize = True
            cb.Checked = True
            cb.Click += self.checkbox_clicked
            cb.Location = Point(cb_x, cb_y)
            cb.Text = re.sub("&", "&", roi_name)  # Escape "&" as "&&" so that it is displayed
            self.rois_gb.Controls.Add(cb)
            cb_y += cb.Height
        self.rois_gb.Height = (len(roi_names) + 1) * 20  # Account for select all
        self.Controls.Add(self.rois_gb)
        self.y += self.rois_gb.Height + 15

        # Plans
        # Default: current plan is checked
        curr_plan = get_current("Plan")
        plan_names = [plan.Name for plan in case.TreatmentPlans if plan.TreatmentCourse is not None and plan.TreatmentCourse.TotalDose is not None and plan.TreatmentCourse.TotalDose.DoseValues is not None]  # All plan names with dose
        if not plan_names:
            MessageBox.Show("There are no plans with dose in the current case. Click OK to abort script.", "No Plans")
            sys.exit(1)
        self.plans_gb = GroupBox()
        self.plans_gb.AutoSize = True
        self.plans_gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.plans_gb.Location = Point(15, self.y)
        self.plans_gb.Text = "Plan(s):"
        cb_y = 20
        cb = CheckBox()
        cb.AutoSize = True
        cb.ThreeState = True
        if len(plan_names) == 1:
            cb.CheckState = cb.Tag = CheckState.Unchecked if curr_plan is None else CheckState.Checked
        else:
            cb.CheckState = cb.Tag = CheckState.Unchecked if curr_plan is None else CheckState.Indeterminate
        cb.Click += self.select_all_clicked
        cb.Location = Point(15, cb_y)
        cb.Text = "Select all"
        self.plans_gb.Controls.Add(cb)
        cb_y += cb.Height
        for plan_name in plan_names:
            cb = CheckBox()
            cb.AutoSize = True
            curr_plan = get_current("Plan")
            cb.Checked = curr_plan is not None and plan_name == curr_plan.Name
            cb.Click += self.checkbox_clicked
            cb.Location = Point(30, cb_y)
            cb.Text = re.sub("&", "&", plan_name)  # Escape "&" as "&&" so that it is displayed
            self.plans_gb.Controls.Add(cb)
            cb_y += cb.Height
        self.plans_gb.Height = (len(plan_names) + 1) * 20  # Account for select all
        self.Controls.Add(self.plans_gb)
        self.y += self.plans_gb.Height + 15

        # Eval doses (doses on additional set)
        self.eval_doses = {}  # dict of eval dose name: eval dose
        fes = case.TreatmentDelivery.FractionEvaluations
        for fe in fes:
            for doe in fe.DoseOnExaminations:
                exam_name = doe.OnExamination.Name
                for de in doe.DoseEvaluations:
                    if de.Name == "" and hasattr(de, "ForBeamSet"):  # Ensure it's a dose on additional set
                        de_name = "{} on {}".format(de.ForBeamSet.DicomPlanLabel, exam_name)
                        self.eval_doses[de_name] = de

        if self.eval_doses:  # There is dose computed on additional set
            # Add eval dose checkboxes
            self.eval_doses_gb = GroupBox()
            self.eval_doses_gb.AutoSize = True
            self.eval_doses_gb.Location = Point(15, self.y)
            self.eval_doses_gb.Text = "Dose(s) on additional set:"
            cb_y = 20
            cb = CheckBox()
            cb.AutoSize = True
            cb.Text = "Select all"
            cb.Location = Point(15, cb_y)
            cb.ThreeState = True
            cb.CheckState = cb.Tag = CheckState.Unchecked  # Default: no eval doses checked
            cb.Click += self.select_all_clicked
            self.eval_doses_gb.Controls.Add(cb)
            cb_y += cb.Height
            for name in self.eval_doses:
                cb = CheckBox()
                cb.AutoSize = True
                cb.Location = Point(30, cb_y)
                cb.Checked = False  # Default: no eval doses checked
                cb.Click += self.checkbox_clicked
                cb.Text = name
                self.eval_doses_gb.Controls.Add(cb)
                cb_y += cb.Height
            self.eval_doses_gb.Height = (len(self.eval_doses) + 1) * 20  # Account for select all
            self.Controls.Add(self.eval_doses_gb)
            self.y += self.eval_doses_gb.Height + 15

        # Add dose checkboxes
        doses = ["D0.035cc", "D100%", "D99%", "D98%", "D95%", "D50%", "D2%", "D1%"]
        self.doses_gb = GroupBox()
        self.doses_gb.AutoSize = True
        self.doses_gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.doses_gb.Location = Point(15, self.y)
        self.doses_gb.Text = "Dose(s):"
        cb_y = 20
        cb = CheckBox()
        cb.AutoSize = True
        cb.IsThreeState = True
        cb.CheckState = cb.Tag = CheckState.Checked
        cb.Click += self.select_all_clicked
        cb.Location = Point(15, cb_y)
        cb.Text = "Select all"
        self.doses_gb.Controls.Add(cb)
        cb_y += cb.Height
        for dose in doses:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Checked = True
            cb.Click += self.checkbox_clicked
            cb.Location = Point(30, cb_y)
            cb.Text = dose
            self.doses_gb.Controls.Add(cb)
            cb_y += cb.Height
        self.doses_gb.Height = (len(doses) + 1) * 20  # Account for select all
        self.Controls.Add(self.doses_gb)
        # Don't increase y just yet: volume groupbox is at same y as doses groupbox


        # Add option to add custom dose checkboxes
        ht = self.doses_gb.Height - 15
        self.dose_tb = TextBox()
        self.dose_tb.Location = Point(30, ht)
        self.dose_tb.Width = 40
        self.doses_gb.Controls.Add(self.dose_tb)

        self.dose_pct_rb = RadioButton()
        self.dose_pct_rb.AutoSize = True
        self.dose_pct_rb.Checked = True
        self.dose_pct_rb.Click += self.dose_pct_rb_clicked
        self.dose_pct_rb.Location = Point(85, ht)
        self.dose_pct_rb.Text = "%"
        self.dose_pct_rb.Width = 35
        self.doses_gb.Controls.Add(self.dose_pct_rb)

        self.cc_rb = RadioButton()
        self.cc_rb.AutoSize = True
        self.cc_rb.Click += self.cc_rb_clicked
        self.cc_rb.Location = Point(120, ht)
        self.cc_rb.Text = "cc"
        self.cc_rb.Width = 35
        self.doses_gb.Controls.Add(self.cc_rb)

        self.dose_add_btn = Button()
        self.dose_add_btn.AutoSize = True
        self.dose_add_btn.Click += self.dose_add_btn_clicked
        self.dose_add_btn.Location = Point(170, ht)
        self.dose_add_btn.Text = "Add Dose"
        self.dose_add_btn.Width = 30
        self.doses_gb.Controls.Add(self.dose_add_btn)

        # Add volume checkboxes
        x = 30 + self.doses_gb.Width
        vols = ["V100%", "V99%", "V98%", "V95%", "V50%", "V2%", "V1%"]
        self.vols_gb = GroupBox()
        self.vols_gb.AutoSize = True
        self.vols_gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.vols_gb.Location = Point(x, self.y)
        self.vols_gb.Text = "Volume(s):"
        cb_y = 20
        cb = CheckBox()
        cb.AutoSize = True
        cb.IsThreeState = True
        cb.CheckState = cb.Tag = CheckState.Checked
        cb.Click += self.select_all_clicked
        cb.Location = Point(15, cb_y)
        cb.Text = "Select all"
        self.vols_gb.Controls.Add(cb)
        cb_y += cb.Height
        for vol in vols:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Checked = True
            cb.Click += self.checkbox_clicked
            cb.Location = Point(30, cb_y)
            cb.Text = vol
            self.vols_gb.Controls.Add(cb)
            cb_y += cb.Height
        self.vols_gb.Height = (len(vols) + 1) * 20  # Account for select all
        self.Controls.Add(self.vols_gb)
        self.y += max(self.doses_gb.Height, self.vols_gb.Height) + 15

        # Add option to add custom volume checkboxes
        ht = self.vols_gb.Height - 15
        self.vol_tb = TextBox()
        self.vol_tb.Location = Point(30, ht)
        self.vol_tb.Width = 40
        self.vols_gb.Controls.Add(self.vol_tb)

        self.vol_pct_rb = RadioButton()
        self.vol_pct_rb.AutoSize = True
        self.vol_pct_rb.Checked = True
        self.vol_pct_rb.Click += self.vol_pct_rb_clicked
        self.vol_pct_rb.Location = Point(85, ht)
        self.vol_pct_rb.Text = "%"
        self.vol_pct_rb.Width = 35
        self.vols_gb.Controls.Add(self.vol_pct_rb)

        self.gy_rb = RadioButton()
        self.gy_rb.AutoSize = True
        self.gy_rb.Click += self.gy_rb_clicked
        self.gy_rb.Location = Point(120, ht)
        self.gy_rb.Text = "Gy"
        self.gy_rb.Width = 35
        self.vols_gb.Controls.Add(self.gy_rb)

        self.vol_add_btn = Button()
        self.vol_add_btn.AutoSize = True
        self.vol_add_btn.Click += self.vol_add_btn_clicked
        self.vol_add_btn.Location = Point(170, ht)
        self.vol_add_btn.Text = "Add Volume"
        self.vol_add_btn.Width = 30
        self.vols_gb.Controls.Add(self.vol_add_btn)

        self.y += self.dose_tb.Height + 15

        # Set uniform groupbox width
        widths = [self.rois_gb.Width, self.plans_gb.Width, self.doses_gb.Width + self.vols_gb.Width + 15]
        to_resize = [self.rois_gb, self.plans_gb]
        if hasattr(self, "eval_doses_gb"):
            widths.append(self.eval_doses_gb.Width)
            to_resize.append(self.eval_doses_gb)
        for gb in to_resize:
            gb.MinimumSize = Size(max(widths), 0)

        # Add "Compute" button
        self.compute_btn = Button()
        self.compute_btn.AutoSize = True
        self.compute_btn.Click += self.compute_btn_clicked
        x = self.ClientSize.Width - 15 - self.compute_btn.Width
        self.compute_btn.Location = Point(x, self.y)
        self.compute_btn.Text = "Compute"
        self.Controls.Add(self.compute_btn)

        self.ShowDialog()  # Launch window

    def add_dose_checkbox(self, text):
        # Helper method that adds a dose checkbox to the doses groupbox
        # If dose already exists, check that checkbox
        # Assume `text` is a valid dose (e.g., "D1cc")

        match = [ctrl for ctrl in self.doses_gb.Controls if isinstance(ctrl, CheckBox) and ctrl.Text == text]
        if match:
            match[0].Checked = True
            self.dose_tb.Text = ""
        else:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Checked = True
            cb.Click += self.checkbox_clicked
            cb.Location = Point(30, self.dose_tb.Location.Y)
            cb.Text = text
            self.doses_gb.Controls.Add(cb)

            # Move custom input down a row
            for ctrl in [self.dose_tb, self.dose_pct_rb, self.cc_rb, self.dose_add_btn]:
                ctrl.Top += cb.Height
            
            # Clear custom input
            self.dose_tb.Text = ""
            self.dose_pct_rb.Checked = True

            # Move compute button down, if necessary
            if self.doses_gb.Height > self.vols_gb.Height:
                self.compute_btn.Top += cb.Height
                self.y += cb.Height

        # Change "select all" check state, if necessary
        select_all = list(self.doses_gb.Controls)[0]
        cbs_cked = [ctrl.Checked for ctrl in list(self.doses_gb.Controls)[1:] if isinstance(ctrl, CheckBox)]
        if all(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Checked
        elif any(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Indeterminate
        else:
            select_all.CheckState = select_all.Tag = CheckState.Unchecked
        self.set_compute_enabled()

    def add_vol_checkbox(self, text):
        # Helper method that adds a volume checkbox to the volumes groupbox
        # If volume already exists, check that checkbox
        # Assume `text` is a valid volume (e.g., "V20Gy")

        match = [ctrl for ctrl in self.vols_gb.Controls if isinstance(ctrl, CheckBox) and ctrl.Text == text]
        if match:
            match[0].Checked = True
            self.vol_tb.Text = ""
        else:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Checked = True
            cb.Click += self.checkbox_clicked
            cb.Location = Point(30, self.vol_tb.Location.Y)
            cb.Text = text
            self.vols_gb.Controls.Add(cb)

            # Move custom input down a row
            for ctrl in [self.vol_tb, self.vol_pct_rb, self.gy_rb, self.vol_add_btn]:
                ctrl.Top += cb.Height

            # Move compute button down, if necessary
            if self.vols_gb.Height > self.doses_gb.Height:
                self.compute_btn.Top += cb.Height
                self.y += cb.Height

            # Clear custom input
            self.vol_tb.Text = ""
            self.vol_pct_rb.Checked = True

        # Change "select all" check state, if necessary
        select_all = list(self.vols_gb.Controls)[0]
        cbs_cked = [ctrl.Checked for ctrl in list(self.vols_gb.Controls)[1:] if isinstance(ctrl, CheckBox)]
        if all(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Checked
        elif any(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Indeterminate
        else:
            select_all.CheckState = select_all.Tag = CheckState.Unchecked
        self.set_compute_enabled()

    def select_all_clicked(self, sender, event):
        # Helper method that checks/unchecks checkboxes when a "select all" checkbox is clicked

        cbs = [ctrl for ctrl in list(sender.Parent.Controls)[1:] if isinstance(ctrl, CheckBox)]
        if sender.Tag == CheckState.Checked:  # If checked, uncheck
            sender.CheckState = sender.Tag = CheckState.Unchecked
            for cb in cbs:
                cb.Checked = False
        else:  #If unchecked or indeterminate, check
            sender.CheckState = sender.Tag = CheckState.Checked
            for cb in cbs:
                cb.Checked = True
        self.set_compute_enabled()
        
    def checkbox_clicked(self, sender, event):
        # Helper method that sets the check state of the "select all" checkbox in the group, when another checkbox in that group is clicked

        select_all = list(sender.Parent.Controls)[0]
        cbs_cked = [ctrl.Checked for ctrl in list(sender.Parent.Controls)[1:] if isinstance(ctrl, CheckBox)]
        if all(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Checked
        elif any(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Indeterminate
        else:
            select_all.CheckState = select_all.Tag = CheckState.Unchecked
        self.set_compute_enabled()

    def set_compute_enabled(self):
        # Helper method that enabled or disables "Compute" button
        # Enable only if at least one ROI, at least one plan or eval dose, and at least one dose or volume is checked

        roi_cked = any(cb.Checked for cb in self.rois_gb.Controls)
        plan_cked = any(cb.Checked for cb in self.plans_gb.Controls)
        eval_dose_cked = any(cb.Checked for cb in self.eval_doses_gb.Controls) if hasattr(self, "eval_doses_gb") else False
        dose_cked = any(cb.Checked for cb in self.doses_gb.Controls if isinstance(cb, CheckBox))
        vol_cked = any(cb.Checked for cb in self.vols_gb.Controls if isinstance(cb, CheckBox))
        
        self.compute_btn.Enabled = roi_cked and (plan_cked or eval_dose_cked) and (dose_cked or vol_cked)

    # Helper methods for unchecking unit radiobuttons when other unit radiobutton is checked
    # These are necessary b/c the radiobuttons aren't the only controls in their groupboxes
    def dose_pct_rb_clicked(self, sender, event):
        self.cc_rb.Checked = False

    def cc_rb_clicked(self, sender, event):
        self.dose_pct_rb.Checked = False

    def vol_pct_rb_clicked(self, sender, event):
        self.gy_rb.Checked = False

    def gy_rb_clicked(self, sender, event):
        self.vol_pct_rb.Checked = False

    def dose_add_btn_clicked(self, sender, event):
        # Helper method that validates dose input and calls another method to add dose checkbox

        vol = self.dose_tb.Text.strip()  # Extract user input
        if self.dose_pct_rb.Checked:  # D__%
            # Validate user input as percentage
            try:
                vol = float(vol)
                if  vol < 0 or vol > 100:
                    msg = "Percentage must be between 0 and 100, inclusive. Try again."
                    MessageBox.Show(msg, "Invalid Input", MessageBoxButtons.OK)
                else:
                    self.add_dose_checkbox("D{}%".format(format_num(vol, 4)))  # Round to 4 decimal places
            except:
                msg = "Input must be numeric. Try again."
                MessageBox.Show(msg, "Invalid Input", MessageBoxButtons.OK)
        else:  # D__cc
            # Validate user input as non-negative
            try:
                vol = float(vol)
                if vol < 0:
                    msg = "Volume must be non negative. Try again."
                    MessageBox.Show(msg, "Invalid Input", MessageBoxButtons.OK)
                else:
                    self.add_dose_checkbox("D{}cc".format(format_num(vol, 4)))  # Round to 4 decimal places
            except:
                msg = "Input must be numeric. Try again."
                MessageBox.Show(msg, "Invalid Input", MessageBoxButtons.OK)

    def vol_add_btn_clicked(self, sender, event):
        # Helper method that validates volume input and calls another method to add volume checkbox

        dose = self.vol_tb.Text.strip()  # Extract user input
        if self.vol_pct_rb.Checked:  # V__%
            # Validate user input as percentage
            try:
                dose = float(dose)
                if dose < 0 or dose > 100:
                    msg = "Percentage must be between 0 and 100, inclusive. Try again."
                    MessageBox.Show(msg, "Invalid Input", MessageBoxButtons.OK)
                else:
                    self.add_vol_checkbox("V{}%".format(format_num(dose, 4)))  # Round to 4 decimal places
            except:
                msg = "Input must be numeric. Try again."
                MessageBox.Show(msg, "Invalid Input", MessageBoxButtons.OK)
        else:  # V__gy
            # Validate user input as non-negative
            try:
                dose = float(dose)
                if dose < 0:
                    msg = "Dose must be non negative. Try again."
                    MessageBox.Show(msg, "Invalid Input", MessageBoxButtons.OK)
                else:
                    self.add_vol_checkbox("V{}Gy".format(format_num(dose, 4)))  # Round to 4 decimal places
            except:
                msg = "Input must be numeric. Try again."
                MessageBox.Show(msg, "Invalid Input", MessageBoxButtons.OK)
                
    def compute_btn_clicked(self, sender, event):
        # When user clicks "Compute", set DialogResult

        self.DialogResult = DialogResult.OK


class TableData():
    # Helper class to store data about a table
    # Attributes:
    # row: list of Paragraph objects (cell data)
    # style: list of tuples, to be passed to TableStyle constructor in main function

    def __init__(self, row, roi_name, num_plans, plan_num):
        # Initialize a TableData object
        # row: list of cell values
        # roi_name: name of the ROI that the row is for
        # num_plans: total number of plans (including eval doses) that we are computing data for
        # plan_num: 0-based index of the plan that this row is for

        row = [Paragraph(str(text), style=tbl_data_style) for text in row]  # Create Paragraph objects from cell data (may be floats or strs, so convert to strs just in case)

        # Base color
        roi_color = case.PatientModel.RegionsOfInterest[roi_name].Color
        r, g, b, a = roi_color.R, roi_color.G, roi_color.B, roi_color.A
        roi_color = toColor("rgba({}, {}, {}, {})".format(r, g, b, a))

        # Plan color is a different hue of base color
        a = 0.5 + 0.5 / (num_plans + 2) * (plan_num + 1)
        plan_color = toColor("rgba({}, {}, {}, {})".format(r, g, b, a))

        self.row = row
        self.style = [
            ("VALIGN", (0, 0), (-1, 0), "TOP"),  # Stats at very bottom of cell are an eyesore (can happen if plan name spans 2 lines), but don't want to center align b/c there may be blank cells below the cell
            # Line before after last column (lines between other columns are added in below loop)
            ("LINEAFTER", (-1, 0), (-1, 0), 0.25, black),
            # First column color is ROI color
            ("BACKGROUND", (0, 0), (0, 0), roi_color),
            ("TEXTCOLOR", (0, 0), (0, 0), get_text_color(roi_color)),
            # Remaining column colors are plan color (based on ROI color)
            ("BACKGROUND", (1, 0), (-1, 0), plan_color),
            ("TEXTCOLOR", (1, 0), (-1, 0), get_text_color(plan_color))
        ]
        for i, val in enumerate(row):
            # Lines between columns (line before first column is in above list)
            self.style.append(("LINEBEFORE", (i, 0), (i, 0), 0.25, black))
            # Thin line below each cell with data
            # If there is no data, we want the illusion of a span, so no line
            if val.text != "":
                self.style.append(("LINEABOVE", (i, 0), (i, 0), 0.25, black))


def dose_statistics():
    """Print a PDF report of selected dose statistics (dose at volume, volume at dose) for the selected ROIs, plans, and doses on additional set

    Report is saved to "T:\Physics - T\Scripts\Output Files\DoseStatistics" and automatically opened

    User chooses ROIs, plans, doses on additional set, doses, and volumes from a GUI, which also allows (and validates) custom dose and vol input
    Table hierarchy is ROI -> Plan or eval dose -> dose or volume
    Both absolute and relative dose or volume are displayed
    Rows are color coded based on ROI color in RS, with plan colors based on ROI color
    
    Rows for an empty geometry have "[Empty]" in the cells instead of a number
    Absolute volumes (e.g., "V1000cc") that are greather than the geometry volume have "[<vol> cc > ROI vol]" in the cells instead of a number
    
    Table is implemented as many tables stacked on top of each other
    The ROI-colored lines between rows are unavoidable because "LINEABOVE" and "LINEBELOW" styles do not support alpha values other than 255
    Column widths are based on maximum cell data width in that column
    """

    global case

    # Get current variables
    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort the script.", "No Patient Loaded")
        sys.exit(1)  # Exit script with an error
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort the script.", "No Case Loaded")
        sys.exit(1)  # Exit script with an error

    form = DoseStatisticsForm()
    if form.DialogResult != DialogResult.OK:
        sys.exit()  # User exited GUI

    # Get user input from form
    # Start at 2nd control in each groupbox to account for "select all" checkbox
    roi_names = [cb.Text for cb in list(form.rois_gb.Controls)[1:] if cb.Checked]
    plan_names = [cb.Text for cb in list(form.plans_gb.Controls)[1:] if cb.Checked]
    eval_doses = {cb.Text: form.eval_doses[cb.Text] for cb in list(form.eval_doses_gb.Controls)[1:] if cb.Checked} if hasattr(form, "eval_doses_gb") else {}  # Use empty dict if eval doses were not an option in the GUI
    doses = [cb.Text for cb in list(form.doses_gb.Controls)[1:] if isinstance(cb, CheckBox) and cb.Checked]  # Ignore the controls used for custom input
    doses.sort(key=sorting_key)  # Sort by absolute vs. relative, then by volume
    vols = [cb.Text for cb in list(form.vols_gb.Controls)[1:] if isinstance(cb, CheckBox) and cb.Checked]
    vols.sort(key=sorting_key)  # Sort by absolute vs. relative, then by dose

    # Create PDF
    pt_name = "{}, {}".format(patient.Name.split("^")[0], patient.Name.split("^")[1])  # e.g., "Jones, Bill"
    filepath = r"\\vs20filesvr01\groups\CANCER\Physics - T\Scripts\Output Files\DoseStatistics\{} {} Dose Statistics.pdf".format(pt_name, case.CaseName)
    pdf = SimpleDocTemplate(filepath, pagesize=letter, bottomMargin=0.25 * inch, leftMargin=0.25 * inch, rightMargin=0.25 * inch, topMargin=0.25 * inch)  # 8.5 x 11", 1/4" margins

    num_dose_vols = max(len(doses), len(vols))  # Number of rows for this plan
    
    # List of TableData objects from which to create Table objects later
    # We can't go ahead and create TableData objects as we go, b/c we don't yet know the column widths
    tbl_datas = []  

    for roi_name in roi_names:
        for i, plan_name in enumerate(plan_names + list(eval_doses.keys())):  # Iterate over all plan and eval dose names (plan names first)
            if plan_name in plan_names:  # It's a plan
                plan = case.TreatmentPlans[plan_name]
                dose_dist = plan.TreatmentCourse.TotalDose
                rx = sum(beam_set.Prescription.PrimaryDosePrescription.DoseValue for beam_set in plan.BeamSets if beam_set.Prescription.PrimaryDosePrescription is not None)  # Total plan dose is sum of beam set doses
                geom = plan.GetStructureSet().RoiGeometries[roi_name]  # ROI geometry on planning exam
            else:  # It's an eval dose
                dose_dist = eval_doses[plan_name]  # Dose distribution is eval dose itself
                exam_name = plan_name.split(" on ")[1]  # e.g., "T1 on QACT" -> "QACT"
                plan_name = re.sub(" on ", "<br/>on ", plan_name)  # Add a newline to control column width ("\n" character does not work); e.g., "T1 on QACT" -> "T1<br/>on QACT"
                rx = dose_dist.ForBeamSet.Prescription.PrimaryDosePrescription.DoseValue
                geom = case.PatientModel.StructureSets[exam_name].RoiGeometries[roi_name]  # ROI geometry on additional set that dose is computed on
            
            roi_vol = geom.GetRoiVolume() if geom.HasContours() else 0  # Empty geometry has a volume of zero

            for j in range(num_dose_vols):
                if i == j == 0:  # First row for this ROI
                    row = [roi_name, plan_name]
                elif j == 0:  # First row for this plan, but not for this ROI
                    row = ["", plan_name]
                else:  # Not the first row of this ROI or this plan
                    row = ["", ""]
                if j < len(doses):  # There are more doses to compute for this plan
                    # Compute dose
                    dose = doses[j]
                    row.append(html_txt(dose))
                    if roi_vol:  # Nonempty geometry
                        if dose.endswith("%"):  # D__%
                            rel_vol = float(dose[1:-1]) / 100  # e.g., "D90%" -> 0.9
                            abs_vol = rel_vol * roi_vol  # e.g., 0.9 * 500.0 = 450.0
                        else:  # D__cc
                            abs_vol = float(dose[1:-2])  # e.g., "D1cc" -> 1
                            rel_vol = abs_vol / roi_vol  # e.g., 1 / 500.0 = 0.002 (don't need to worry about integer division b/c ROI vol is always a float)
                        if rel_vol > 1:  # Absolute volume > ROI volume
                            abs_dose = rel_dose = "[{} cc > ROI vol]".format(format_num(abs_vol, 4))  # e.g., "[1000 cc > ROI vol]""
                        else:  # Absolute volume <= ROI volume
                            abs_dose = dose_dist.GetDoseAtRelativeVolumes(RoiName=roi_name, RelativeVolumes=[rel_vol])[0]  # e.g., 1000.0
                            rel_dose = format_num(abs_dose / rx * 100)  # Convert to %; e.g., 1000.0 / 5000.0 * 100 = 20.0 
                            abs_dose = format_num(abs_dose)  # Wait until here to format abs_dose because we first use it to compute rel_dose
                        row.extend([abs_dose, rel_dose])
                    else:  # Empty geometry
                        row.extend(["[Empty]"] * 2)
                elif doses:  # All doses have already been computed for this plan
                    row.extend([""] * 3)
                if j < len(vols):  # There are more volumes to compute for this plan
                    # Compute vol
                    vol = vols[j]
                    row.append(html_txt(vol))
                    # Add vol
                    if roi_vol:  # Nonempty geometry
                        if vol.endswith("%"):  # V__%
                            rel_dose = float(vol[1:-1]) / 100  # e.g., "V90%" -> 0.9
                            abs_dose = rel_dose * rx  # e.g., 0.9 * 5000.0 = 4500.0
                        else:  #V__Gy
                            abs_dose = float(vol[1:-2]) * 100  # Convert to cGy; e.g., "V20Gy" -> 20.0 * 100 = 2000.0
                            rel_dose = abs_dose / rx  # e.g., 2000.0 / 5000.0 = 0.4
                        rel_vol = dose_dist.GetRelativeVolumeAtDoseValues(RoiName=roi_name, DoseValues=[abs_dose])[0]  # e.g., 0.1
                        abs_vol = rel_vol * roi_vol  # e.g, 0.1 * 500.0 = 50.0
                        rel_vol *= 100  # Convert to %; wait until this line because we first use abs_vol to compute rel_vol
                        row.extend([format_num(abs_vol), format_num(rel_vol)])
                    else:  # Empty geometry
                        row.extend(["[Empty]"] * 2)
                elif vols:  # All vols have already been computed for this plan
                    row.extend([""] * 3)

                # Add TableData object to list
                tbl_data = TableData(row, roi_name, len(plan_names) + len(eval_doses), i)
                tbl_datas.append(tbl_data)

    # Data for header table
    row_1 = ["ROI", "Plan", "Dose Statistics"]
    row_2 = [""] * 2  # B/c "ROI" and "Plan" span 1st and 2nd rows
    if doses:
        row_1.extend([""] * 2)  # 2 more columns for "Dose Statistics" to span (to make "Dose", "cGy", and "% Rx" columns) 
        row_2.extend(["Dose", "cGy", "% Rx"])
        if vols:
            row_1.extend([""] * 3)  # 3 more columns for "Dose Statistics" to span (to make ""Volume", "cc", and "% ROI vol" columns)
            row_2.extend(["Volume", "cc", "% ROI vol"]) 
    elif vols:
        row_1.extend([""] * 2)  # 2 more columns for "Dose Statistics" to span (to make "Volume", "cc", and "% ROI vol" columns)
        row_2.extend(["Volume", "cc", "% ROI vol"])

    # Make a Paragraph from each row text
    row_1 = [Paragraph(text, style=tbl_hdg_style) for text in row_1]
    row_2 = [Paragraph(text, style=tbl_hdg_style) for text in row_2]

    # Compute column widths
    # Ignore first element ("Dose Statistics") of column 2
    col_widths = []
    for i, paragraph in enumerate(row_2):
        row_2_width = string_width(paragraph)
        max_tbl_data_width = max([string_width(tbl_data.row[i]) for tbl_data in tbl_datas])
        if i == 2:
            max_width = max(row_2_width, max_tbl_data_width)
        else:
            row_1_width = string_width(row_1[i])
            max_width = max(row_1_width, row_2_width, max_tbl_data_width)
        col_widths.append(15 + max_width)  # Leave 15 left-right padding
    
    # Create tables
    # Style for header table
    hdr_tbl_style = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),  # Vertically align all contents
        ("SPAN", (0, 0), (0, 1)), # "ROI" spans 2 rows
        ("SPAN", (1, 0), (1, 1)),  # "Plan" spans 2 rows
        ("SPAN", (2, 0), (-1, 0)),  # "Dose Statistics" spans all dose stat rows
        ("BOX", (0, 0), (-1, -1), 0.25, black),  # Black outline
        ("INNERGRID", (0, 0), (-1, -1), 0.25, black),
        ("BACKGROUND", (0, 0), (-1, 1), grey)  # Header row color
    ]
    hdr_tbl = Table([row_1, row_2], style=TableStyle(hdr_tbl_style), colWidths=col_widths)

    # Heading & subheading
    heading = Paragraph(pt_name, style=hdg_style)  # e.g., "Jones, Bill"
    mrn = Paragraph("MR#: {}".format(patient.PatientID), style=subhdg_style)  # e.g., "MR#: 000123456"
    dose_stats = Paragraph("Dose Statistics", style=subhdg_style)
    
    elems = [KeepTogether([heading, spcr_sm, mrn, spcr_sm, dose_stats, spcr_lg, hdr_tbl])]  # Keep these items together on the first page

    tbl_datas[-1].style.append(("LINEBELOW", (0, 0), (-1, 0), 0.25, black))  # Black line at end of report table
    for tbl_data in tbl_datas:
        tbl = Table([tbl_data.row], style=TableStyle(tbl_data.style), colWidths=col_widths)
        elems.append(tbl)
    
    # Create PDF if possible
    # Assume that failure means report is already open
    try:
        pdf.build(elems)  # Add elements to PDF
    except:
        MessageBox.Show("A dose statistics report is already open for this case. Close it and then re-run the script.", "DoseStatistics")
        sys.exit(1)

    # Open report
    reader_paths = [r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe", r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"]  # Paths to Adobe Reader on RS servers
    for reader_path in reader_paths:
        try:
            os.system(r'START /B "{}" "{}"'.format(reader_path, filepath))
            break
        except:
            continue
