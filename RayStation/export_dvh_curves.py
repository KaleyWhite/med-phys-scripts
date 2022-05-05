import clr
from collections import OrderedDict
from datetime import datetime
import os
import re
import sys
from typing import Optional

import numpy as np

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System import EventArgs  # For type hints
from System.Drawing import *
from System.Windows.Forms import *

sys.path.append(os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-scripts', 'RayStation'))
from dose_grid_box import dose_grid_box


OUTPUT_DIR = os.path.join('T:', os.sep, 'Physics', 'Scripts', 'Output Files', 'Export DVH Curves')


class ExportDvhCurvesForm(Form):
    """Form that allows the user to choose parameters for an exported DVH curve file.

    User chooses:
    - Dose unit: cGy or Gy
    - Volume unit: % or cc
    - Dose or volume interval (bin width)
    - Dose(s): plan, beam set, beam, eval
    - Extra header info to include in the DVH files:
        * Volume unit
        * ROI volume
        * Min dose
        * Max dose
        * Mean dose
    - ROI(s)
    - Filetype(s) of exported DVH files: .DVH, .TXT
    """

    def __init__(self, case: PyScriptObject) -> None:
        """Initializes an ExportDvhCurvesForm object.

        Args:
            case: The case from which to export a DVH curve.
        """
        # Horizontal and vertical coordinates of next control
        x = y = 15

        try:
            curr_plan = get_current('Plan')
        except:
            curr_plan = None

        # -------------- Variables to hold user-inputted values on form -------------- #

        # Whether or not user-inputted dose and vol intervals are valid
        # Form defaults to a volume interval
        self._dose_interval_valid = False
        self._vol_interval_valid = True

        # User-inputted dose and volume units
        self.dose_unit = 'cGy'
        self.vol_unit = '%'

        # Pause form update so the user doesn't see a half-finished form
        self.SuspendLayout()

        self._set_up_form()
        
        # Script instructions
        lbl = Label()
        lbl.AutoSize = True
        lbl.AutoSizeMode = AutoSizeMode.GrowAndShrink
        lbl.Location = Point(15, y)
        lbl.Text = 'This script extends RayStation\'s "Export DVH curves" functionality.'
        self.Controls.Add(lbl)
        y += lbl.Height + 5

        # -------------------------------- User inputs ------------------------------- #

        # Dose unit
        gb = GroupBox()
        gb.AutoSize = True
        gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        gb.Location = Point(x, y)
        gb.Text = 'Dose unit:'
        self.Controls.Add(gb)
        
        rb_x = 15 
        self.dose_unit_rbs = []

        rb = RadioButton()
        rb.AutoSize = True
        rb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        rb.Text = 'cGy'
        rb.Checked = True
        rb.Click += self._dose_rb_Click
        rb.Location = Point(rb_x, 15)
        gb.Controls.Add(rb)
        self.dose_unit_rbs.append(rb)
        rb_x += rb.Width + 5
        
        rb = RadioButton()
        rb.AutoSize = True
        rb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        rb.Text = 'Gy'
        rb.Checked = False
        rb.Click += self._dose_rb_Click
        rb.Location = Point(rb_x, 15)
        gb.Controls.Add(rb)
        self.dose_unit_rbs.append(rb)
        rb_x += rb.Width + 15

        x += rb_x + 45

        # Volume unit
        gb = GroupBox()
        gb.AutoSize = True
        gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        gb.Location = Point(x, y)
        gb.Text = 'Volume unit:'
        self.Controls.Add(gb)

        rb_x = 15
        self.vol_unit_rbs = []

        rb = RadioButton()
        rb.AutoSize = True
        rb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        rb.Text = '%'
        rb.Checked = True
        rb.Click += self._vol_rb_Click
        rb.Location = Point(rb_x, 15)
        gb.Controls.Add(rb)
        self.vol_unit_rbs.append(rb)
        rb_x += rb.Width + 15

        rb = RadioButton()
        rb.AutoSize = True
        rb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        rb.Text = 'cc'
        rb.Checked = False
        rb.Click += self._vol_rb_Click
        rb.Location = Point(rb_x, 15)
        gb.Controls.Add(rb)
        self.vol_unit_rbs.append(rb)
        rb_x += rb.Width + 15

        x += rb_x + 45

        # Dose or volume interval
        gb = GroupBox()
        gb.AutoSize = True
        gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        gb.Location = Point(x, y)
        gb.Text = 'Interval:'
        self.Controls.Add(gb)
        
        interval_x = 15

        # Dose interval
        self.dose_interval_rb = RadioButton()
        self.dose_interval_rb.AutoSize = True
        self.dose_interval_rb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.dose_interval_rb.Checked = False
        self.dose_interval_rb.Location = Point(interval_x, 15)
        self.dose_interval_rb.Text = 'Dose:'
        self.dose_interval_rb.Click += self._interval_rb_Click
        gb.Controls.Add(self.dose_interval_rb)
        interval_x += self.dose_interval_rb.Width + 5

        self.dose_interval_tb = TextBox()
        self.dose_interval_tb.Enabled = False
        self.dose_interval_tb.Width = 30
        self.dose_interval_tb.Location = Point(interval_x, 15)
        self.dose_interval_tb.TextChanged += self._validate_dose_interval
        gb.Controls.Add(self.dose_interval_tb)
        interval_x += self.dose_interval_tb.Width + 5

        self.dose_interval_unit = Label()
        self.dose_interval_unit.AutoSize = True
        self.dose_interval_unit.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.dose_interval_unit.Location = Point(interval_x, 15)
        self.dose_interval_unit.Text = 'cGy'
        gb.Controls.Add(self.dose_interval_unit)
        interval_x += self.dose_interval_unit.Width + 15

        self.dose_interval_tb.TextChanged += self._set_export_enabled
        self.dose_interval_unit.TextChanged += self._set_export_enabled

        # Volume interval
        self.vol_interval_rb = RadioButton()
        self.vol_interval_rb.AutoSize = True
        self.vol_interval_rb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.vol_interval_rb.Checked = True
        self.vol_interval_rb.Location = Point(interval_x, 15)
        self.vol_interval_rb.Text = 'Volume:'
        self.vol_interval_rb.Click += self._interval_rb_Click
        gb.Controls.Add(self.vol_interval_rb)
        interval_x += self.vol_interval_rb.Width + 5

        self.vol_interval_tb = TextBox()
        self.vol_interval_tb.Text = '0.5'
        self.vol_interval_tb.Location = Point(interval_x, 15)
        self.vol_interval_tb.Width = 30
        self.vol_interval_tb.TextChanged += self._validate_vol_interval
        gb.Controls.Add(self.vol_interval_tb)
        interval_x += self.vol_interval_tb.Width + 5
        
        self.vol_interval_unit = Label()
        self.vol_interval_unit.AutoSize = True
        self.vol_interval_unit.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.vol_interval_unit.Text = '%'
        self.vol_interval_unit.Location = Point(interval_x, 15)
        gb.Controls.Add(self.vol_interval_unit)
        interval_x += self.vol_interval_unit.Width

        self.vol_interval_tb.TextChanged += self._set_export_enabled
        self.vol_interval_unit.TextChanged += self._set_export_enabled

        self.interval_rbs = [self.dose_interval_rb, self.vol_interval_rb]

        y += 75
        x = 15

        max_gb_ht = 0

        # Doses
        gb = GroupBox()
        gb.Location = Point(x, y)
        gb.Text = 'Dose(s):'
        self.Controls.Add(gb)
        
        cb_y = 15
        max_item_width = 0

        cb = CheckBox()
        cb.AutoSize = True
        cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        cb.IsThreeState = True
        cb.CheckState = cb.Tag = CheckState.Unchecked
        cb.Location = Point(15, cb_y)
        cb.Text = 'Select all'
        cb.Click += self._select_all_Click
        gb.Controls.Add(cb)
        cb_y += cb.Height
        max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 15)

        self.dose_cbs: OrderedDict[CheckBox, PyScriptObject] = OrderedDict()
        for plan in case.TreatmentPlans:
            dose = plan.TreatmentCourse.TotalDose
            cb = CheckBox()
            cb.AutoSize = True
            cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
            cb.Checked = False
            cb.Location = Point(30, cb_y)
            gb.Controls.Add(cb)
            cb.Click += self._checkbox_Click
            cb.Text = 'Plan - ' + plan.Name
            cb_y += cb.Height
            max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 30)
            self.dose_cbs[cb] = dose

            for beam_set in plan.BeamSets:
                beam_set_dose = beam_set.FractionDose
                cb = CheckBox()
                cb.AutoSize = True
                cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
                cb.Checked = False
                cb.Location = Point(45, cb_y)
                gb.Controls.Add(cb)
                cb.Click += self._checkbox_Click
                cb.Text = 'Beam set - ' + beam_set.DicomPlanLabel
                cb_y += cb.Height
                max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 45)
                self.dose_cbs[cb] = beam_set_dose

                for i, beam in enumerate(beam_set.Beams):
                    dose = beam_set_dose.BeamDoses[i]
                    cb = CheckBox()
                    cb.AutoSize = True
                    cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
                    cb.Checked = False
                    cb.Location = Point(60, cb_y)
                    gb.Controls.Add(cb)
                    cb.Click += self._checkbox_Click
                    cb.Text = 'Beam - ' + plan.Name
                    cb_y += cb.Height
                    max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 60)
                    self.dose_cbs[cb] = dose

        for fe in case.TreatmentDelivery.FractionEvaluations:
            for doe in fe.DoseOnExaminations:
                exam_name = doe.OnExamination.Name
                for de in doe.DoseEvaluations:
                    if de.PerturbedDoseProperties is not None:  # perturbed dose
                        rds = de.PerturbedDoseProperties.RelativeDensityShift
                        density = f'{(rds * 100):.1f}%'
                        iso = de.PerturbedDoseProperties.IsoCenterShift
                        isoctr = f'({iso.x:.1f}, {iso.z:.1f}, {(-iso.y):.1f}) cm'
                        beam_set_name = de.ForBeamSet.DicomPlanLabel
                        dose_txt = f'Perturbed dose of ' + beam_set_name + ':' + density + ', ' + isoctr
                    elif de.Name != '':  # not perturbed, but has a name (probably summed)
                        dose_txt = de.Name
                    elif hasattr(de, 'ByStructureRegistration'):  # registered (mapped) dose
                        reg_name = de.ByStructureRegistration.Name
                        name = de.OfDoseDistribution.ForBeamSet.DicomPlanLabel
                        dose_txt = 'Deformed dose of ' + name + ' by registration ' + reg_name
                    else:  # neither perturbed, summed, nor mapped
                        dose_txt = de.ForBeamSet.DicomPlanLabel
                    
                    cb = CheckBox()
                    cb.AutoSize = True
                    cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
                    cb.Checked = False
                    cb.Location = Point(30, cb_y)
                    gb.Controls.Add(cb)
                    cb.Click += self._checkbox_Click
                    cb.Text = 'Eval - ' + dose_txt
                    cb_y += cb.Height
                    max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 30)
                    self.dose_cbs[cb] = de

        gb.Size = Size(max_item_width + 45, gb.Controls.Count * cb.Height + 30)
        max_gb_ht = max(max_gb_ht, gb.Height)
        x += gb.Width + 15

        # ROIs
        gb = GroupBox()
        gb.Location = Point(x, y)
        gb.Text = 'ROI(s):'
        self.Controls.Add(gb)
        
        cb_y = 15
        max_item_width = 0

        cb = CheckBox()
        cb.AutoSize = True
        cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        cb.IsThreeState = True
        cb.CheckState = cb.Tag = CheckState.Checked
        cb.Location = Point(15, cb_y)
        cb.Text = 'Select all'
        cb.Click += self._select_all_Click
        gb.Controls.Add(cb)
        cb_y += cb.Height
        max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 15)

        self.roi_cbs = []
        for roi in case.PatientModel.RegionsOfInterest:
            cb = CheckBox()
            cb.AutoSize = True
            cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
            cb.Text = roi.Name
            cb.Checked = True
            cb.Location = Point(30, cb_y)
            gb.Controls.Add(cb)
            cb.Click += self._checkbox_Click
            self.roi_cbs.append(cb)
            cb_y += cb.Height
            max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 15)

        gb.Size = Size(max_item_width + 45, gb.Controls.Count * cb.Height + 30)
        max_gb_ht = max(max_gb_ht, gb.Height)
        x += gb.Width + 15

        # "Extras" in DVH file
        extras = ['Volume unit', 'ROI volume', 'Min dose', 'Max dose', 'Mean dose']
        
        gb = GroupBox()
        gb.Location = Point(x, y)
        gb.Text = 'Extra header(s):'
        self.Controls.Add(gb)

        cb_y = 15
        max_item_width = 0

        cb = CheckBox()
        cb.AutoSize = True
        cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        cb.IsThreeState = True
        cb.CheckState = cb.Tag = CheckState.Unchecked
        cb.Text = 'Select all'
        cb.Click += self._select_all_Click
        cb.Location = Point(15, cb_y)
        gb.Controls.Add(cb)
        cb_y += cb.Height
        max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 15)
        
        self.extras_cbs = []
        for extra in extras:
            cb = CheckBox()
            cb.AutoSize = True
            cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
            cb.Text = extra
            cb.Checked = False
            cb.Click += self._checkbox_Click
            cb.Location = Point(30, cb_y)
            gb.Controls.Add(cb)
            self.extras_cbs.append(cb)
            cb_y += cb.Height
            max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 15)
        
        gb.Size = Size(max_item_width + 45, gb.Controls.Count * cb.Height + 30)
        max_gb_ht = max(max_gb_ht, gb.Height)
        x += gb.Width + 15

        # File extension(s)
        file_types = ['.DVH', '.TXT']
        gb = GroupBox()
        gb.Location = Point(x, y)
        gb.Text = 'File type(s):'
        self.Controls.Add(gb)
    
        self.file_type_cbs = []
        cb_y = 15
        max_item_width = 0

        cb = CheckBox()
        cb.AutoSize = True
        cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        cb.IsThreeState = True
        cb.CheckState = cb.Tag = CheckState.Unchecked
        cb.Text = 'Select all'
        cb.Click += self._select_all_Click
        cb.Location = Point(15, cb_y)
        gb.Controls.Add(cb)
        cb_y += cb.Height
        max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 15)

        for file_type in file_types:
            cb = CheckBox()
            cb.AutoSize = True
            cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
            cb.Text = file_type
            cb.Checked = file_type == '.DVH'
            cb.Click += self._checkbox_Click
            cb.Location = Point(30, cb_y)
            gb.Controls.Add(cb)
            self.extras_cbs.append(cb)
            self.file_type_cbs.append(cb)
            cb_y += cb.Height
            max_item_width = max(max_item_width, TextRenderer.MeasureText(cb.Text, cb.Font).Width + 15)
        
        gb.Size = Size(max_item_width + 45, gb.Controls.Count * cb.Height + 30)
        max_gb_ht = max(max_gb_ht, gb.Height)
        y += max_gb_ht + 15

        # "Export" button
        self._export_btn = Button()
        self._export_btn.AutoSize = True
        self._export_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self._export_btn.Enabled = False
        self._export_btn.Location = Point(15, y)
        self._export_btn.Text = 'Export'
        self._export_btn.Click += self._export_btn_Click
        self.Controls.Add(self._export_btn)

        self.ResumeLayout()

    def _set_up_form(self) -> None:
        """Styles the Form"""
        self.Text = 'Export DVH Curves'  # Form title

        # Adapt form size to controls
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)  # At least as wide as title plus some room for "X" button, etc.

        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

    # ---------------------------------------------------------------------------- #
    #                                Event handlers                                #
    # ---------------------------------------------------------------------------- #

    # --------------------------- Dose and volume units -------------------------- #

    def _dose_rb_Click(self, sender: RadioButton, event: EventArgs) -> None:
        """Event handler for clicking a dose unit radio button
        
        Updates the attribute that holds the dose interval unit
        Ensures that the dose interval is still valid, since the unit changed
        """
        self.dose_unit = self.dose_interval_unit.Text = sender.Text
        self._validate_dose_interval()

    def _vol_rb_Click(self, sender: RadioButton, event: EventArgs) -> None:
        """Event handler for clicking a volume unit radio button
        
        Updates the attribute that holds the volume interval unit
        Ensures that the volume interval is still valid, since the unit changed
        """
        self.vol_unit = self.vol_interval_unit.Text = sender.Text
        self._validate_vol_interval()

    # ------------------------- Dose and volume intervals ------------------------ #

    def _interval_rb_Click(self, sender: RadioButton, event: EventArgs) -> None:
        """Event handler clicking a dose or vlume interval radio button
        
        Enables and disables the appropriate interval text boxes when the interval is changed form dose to volume or vice versa
        """
        if sender.Equals(self.dose_interval_rb):  # Dose interval was checked
            self.vol_interval_tb.Enabled = False
            self.vol_interval_tb.BackColor = Color.White
            self.dose_interval_tb.Enabled = True
        else:  # Volume interval was checked
            self.vol_interval_tb.Enabled = True
            self.dose_interval_tb.Enabled = False
            self.dose_interval_tb.BackColor = Color.White
    
    def _validate_dose_interval(self, sender: Optional[TextBox] = None, event: Optional[EventArgs] = None) -> None:
        """Event handler for dose interval textbox text change
        
        May also be called directly. Thus the optional arguments.
        
        Validates user input in dose interval text box
        Dose interval must be a positive number with no more than three decimal places
        If dose unit is cGy, dose interval must be 10,000 or less. If dose unit is Gy, dose interval must be 100 or less
        If invalid, turn textbox red
        """
        if self.dose_interval_rb.Checked:
            self.SuspendLayout()

            dose_interval = self.dose_interval_tb.Text
            dose_unit = self.dose_interval_unit.Text
            try:
                float_dose_interval = float(dose_interval)
                if ('.' in dose_interval and not dose_interval.endswith('.') and len(dose_interval.split('.')[1]) > 3) or float_dose_interval <= 0 or (dose_unit == 'cGy' and float_dose_interval >= 10000) or (dose_unit == 'Gy' and float_dose_interval >= 100):  
                    raise ValueError
                self.dose_interval_tb.BackColor = Color.White
                self._dose_interval_valid = True
            except ValueError:
                self.dose_interval_tb.BackColor = Color.Red
                self._dose_interval_valid = False

            self.ResumeLayout(False)
            self.PerformLayout()

    def _validate_vol_interval(self, sender=None, event=None):
        """Event Handler for volume interval textbox text change

        May also be called directly. Thus the optional arguments.
        
        Validate user input in volume interval text box
        Volume interval must be a positive number with no more than three decimal places
        If volume unit is set to %, volume interval must be 100 or less. If volume unit is cc, volume interval must be 10,000 or less
        If invalid, turn textbox red
        """
        if self.vol_interval_rb.Checked:
            self.SuspendLayout()

            vol_interval = self.vol_interval_tb.Text
            vol_unit = self.vol_interval_unit.Text
            try:
                float_vol_interval = float(vol_interval)
                if ('.' in vol_interval and not vol_interval.endswith('.') and len(vol_interval.split('.')[1]) > 3) or float_vol_interval <= 0 or (vol_unit == '%' and float_vol_interval > 100) or (vol_unit == 'cc' and float_vol_interval >= 10000):  
                    raise ValueError
                self.vol_interval_tb.BackColor = Color.White
                self._vol_interval_valid = True
            except ValueError:
                self.vol_interval_tb.BackColor = Color.Red
                self._vol_interval_valid = True

            self.ResumeLayout(False)
            self.PerformLayout()

    # -------------------------------- Checkboxes -------------------------------- #

    def _checkbox_Click(self, sender: CheckBox, event: EventArgs) -> None:
        """Event handler for clicking a checkbox that is not "Select all"
        
        Sets "Select all" checkbox state based on the sending checkbox's new state
        Enables "Select all" checkbox if any non-"Select all" checkboxes in the group are enabled
        Enables/disables export button
        """
        self.SuspendLayout()

        select_all = list(sender.Parent.Controls)[0]
        cbs = list(sender.Parent.Controls)[1:]  # All non-"Select all" checkboxes in this checkbox's group
        cbs_cked = [cb.Checked for cb in cbs]
        if all(cbs_cked):  # All non-"Select all" checkboxes are checked
            select_all.CheckState = select_all.Tag = CheckState.Checked
        elif any(cbs_cked):  # Some, but not all, non-"Select all" checkboxes are checked
            select_all.CheckState = select_all.Tag = CheckState.Indeterminate
        else:  # All non-"Select all" checboxes are unchecked
            select_all.CheckState = select_all.Tag = CheckState.Unchecked
        select_all.Enabled = any(cb.Enabled for cb in cbs)
        self._set_export_enabled()
        
        self.ResumeLayout(False)
        self.PerformLayout()

    def _select_all_Click(self, sender: CheckBox, event: EventArgs) -> None:
        """Event hanbdler for clicking a "Select all" checkbox
        
        Sets "Select all" checkbox state based on previous state
        Checks or unchecks non-"Select all" checkboxes based on new "Select all" state
        """
        self.SuspendLayout()

        if sender.Tag == CheckState.Checked:  # If checked, uncheck
            sender.CheckState = sender.Tag = CheckState.Unchecked
            for cb in list(sender.Parent.Controls)[1:]:  # The first checkbox is select all
                cb.Checked = False
        else:  # If unchecked or indeterminate, check
            all_enabled = True
            for cb in list(sender.Parent.Controls)[1:]:
                if cb.Enabled:
                    cb.Checked = True
                else:
                    all_enabled = False
            if all_enabled:
                sender.CheckState = sender.Tag = CheckState.Checked   
            else:
                sender.CheckState = sender.Tag = CheckState.Indeterminate
        self._set_export_enabled()

        self.ResumeLayout(False)
        self.PerformLayout()

    # ------------------------------ "Export" button ----------------------------- #

    def _set_export_enabled(self, sender: Optional[TextBox] = None, event: Optional[EventArgs] = None) -> None:
        """Event handler for changing dose or volume interval
        
        May also be called directly. Thus the optional args
        
        Enables "Export" button only if all inputs are valid
        """
        self._export_btn.Enabled = ((self.dose_interval_rb.Checked and self._dose_interval_valid) or (self.vol_interval_rb.Checked and self._vol_interval_valid)) and \
                                   all(any(ctrl.Checked for ctrl in ctrl_list) for ctrl_list in (self.dose_cbs, self.roi_cbs, self.dose_unit_rbs, self.vol_unit_rbs, self.interval_rbs, self.file_type_cbs))

    def _export_btn_Click(self, sender: Button, event: EventArgs) -> None:
        """Event handler for clicking the "Export" button
        
        Set DialogResult
        """
        self.DialogResult = DialogResult.OK


def beam_set_from_beam(case: PyScriptObject, beam: PyScriptObject) -> PyScriptObject:
    """Gets the beam set that the beam belongs to.

    Args:
        case: The case that the beam belongs to
        beam: The beam

    Returns
    -------
    The beam set containing the beam, or None if the beam is not found in the case
    """
    for plan in case.TreatmentPlans:
        for beam_set in plan.BeamSets:
            for b in beam_set.Beams:
                if b.Equals(beam):
                    return beam_set


def struct_set_and_exam_from_dose(case: PyScriptObject, dose: PyScriptObject) -> PyScriptObject:
    """Gets the structure set and the examination associated with the dose distribution

    Arguments
    ---------
    case: The case that the dose distribution belongs to
    dose: The dose distribution

    Returns
    -------
    The structure set associated with the dose distribution
    """
    if hasattr(dose, 'WeightedDoseReferences'):  # Plan dose, dose sum
        struct_set = dose.WeightedDoseReferences[0].DoseDistribution.ForBeamSet.GetStructureSet()
        exam = struct_set.OnExamination
    else:
        if hasattr(dose, 'ForBeamSet'):  # Beam set dose, perturbed dose, dose on additional set
            exam = dose.ForBeamSet.GetPlanningExamination()
        elif hasattr(dose, 'ForBeam'):  # Beam dose
            exam = beam_set_from_beam(case, dose.ForBeam).GetPlanningExamination()
        elif hasattr(dose, 'OfDoseDistribution'):  # Deformed dose
            exam = dose.OfDoseDistribution.ForBeamSet.GetPlanningExamination()
        struct_set = case.PatientModel.StructureSets[exam.Name]
    return struct_set, exam


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


def timestamp() -> str:
    """Gets a text timestamp

    Returns
    -------
    A timestamp in the format YYYY-MM-DD HH_MM_SS

    Example
    -------
    timestamp() -> "2022-05-03 15_05_06"
    """
    return datetime.now().strftime('%Y-%m-%d %H_%M_%S')


def export_dvh_curves():
    """Exports DVH curves from RayStation.

    RayStation includes functionality for exporting a DVH curve. This script extends this functionality:
    - Export multiple DVH curves at a time
    - Export DVH curves for any beam sets, beams, and eval doses in addition to plans
    - Export dose in Gy instead of cGy
    - Export volume in cc instead of %
    - Use a constant, user-specified dose or volume interval
    - Include selected dose statistics in addition to percent volume outside dose grid
    - Exclude ROIs from export
    - Export as ".txt" instead of ".dvh"
    """
    ## Get current variables
    # Patient
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    # Case
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    # Launch form for user input
    form = ExportDvhCurvesForm(case)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
 
    # Get user input from form
    roi_names = sorted(cb.Text for cb in form.roi_cbs if cb.Checked)
    dose_unit = form.dose_unit 
    vol_unit = form.vol_unit
    extras = [cb.Text for cb in form.extras_cbs if cb.Checked]
    file_exts = [cb.Text for cb in form.file_type_cbs if cb.Checked]
    if form.vol_interval_rb.Checked:
        vol_interval = float(form.vol_interval_tb.Text)
        dose_interval = None
    else:
        dose_interval = float(form.dose_interval_tb.Text)
        vol_interval = None
    
    # Iterate over each dose that the user wants a DVH curve for
    for cb, dose in form.dose_cbs.items():
        if not cb.Checked:  # Ignore unchecked doses
            continue

        dose.UpdateDoseGridStructures()  # Update geometries for the dose

        # Get the dose's structure set and exam
        struct_set, exam = struct_set_and_exam_from_dose(case, dose)

        # Dose display text
        dose_type, dose_name = cb.Text.split(' - ')  # E.g., "Plan" and "SBRT Lung_L"
        dose_txt = dose_type + ' dose - ' + dose_name  # E.g., "Plan dose: SBRT Lung_L"
        # For non-eval doses, add exam name to dose text
        if dose_type != 'Eval':
            dose_txt += ' (' + exam.Name + ')'

        # Create dose grid box geometry for computing % ROI volume outside dose grid
        dg_box = dose_grid_box(case, dose)
        dg_box_geom = struct_set.RoiGeometries[dg_box.Name]
        
        # Create temp ROI to hold intersection of ROI geometry and dose grid box
        temp_name = case.PatientModel.GetUniqueRoiName(DesiredName='Temp')
        temp = case.PatientModel.CreateRoi(Name=temp_name, Color='purple', Type='Control')
        
        # Text string that will be written to DVH file(s)
        # First rows of header
        string = '#PatientName:' + format_pt_name(patient) + '\n' \
               + '#PatientId:' + patient.PatientID + '\n' \
               + '#Dosename:' +  dose_txt + '\n' 

        # Add DVH data for each selected ROI
        for roi_name in roi_names:
            geom = struct_set.RoiGeometries[roi_name]
            if not geom.HasContours():  # Ignore empty geometries
                continue

            max_dose = dose.GetDoseStatistic(RoiName=roi_name, DoseType='Max')
            total_vol = geom.GetRoiVolume()

            string += '#RoiName:' + roi_name + '\n'
            
            # Compute % ROI volume outside dose grid
            temp.CreateAlgebraGeometry(Examination=exam, ExpressionA={'Operation': 'Union', 'SourceRoiNames': [roi_name], 'MarginSettings': {'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [dg_box.Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ResultOperation='Subtraction', ResultMarginSettings={'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})
            dose.UpdateDoseGridStructures()  # So we can get dose to temp geometry
            temp_geom = struct_set.RoiGeometries[temp.Name]
            if temp_geom.HasContours():
                temp_vol = temp_geom.GetRoiVolume()
                pct_vol_outside_dg = int(round(temp_vol / total_vol * 100))
            else:
                pct_vol_outside_dg = 0
            string += f'#Roi volume fraction outside grid: {pct_vol_outside_dg}%\n'

            string += '#Dose unit: ' + dose_unit + '\n'

            # Add user-selected extras to header
            if 'Volume unit' in extras:
                string += '#Volume unit: ' + vol_unit + '\n'

            if 'ROI volume' in extras:
                string += f'#Roi volume: {total_vol:.2f}\n'

            if 'Min dose' in extras:
                min_dose = dose.GetDoseStatistic(RoiName=roi_name, DoseType='Min')
                if dose_unit == 'cGy':
                    string += f'#Min dose: {int(round(min_dose))}\n'
                elif dose_unit == 'Gy':
                    string += f'#Min dose: {(min_dose / 100):.2f}\n'
                    
            if 'Max dose' in extras:
                if dose_unit == 'cGy':
                    string += f'#Max dose: {int(round(max_dose))}\n'
                elif dose_unit == 'Gy':
                    string += f'#Max dose: {(max_dose / 100):.2f}\n'

            if 'Mean dose' in extras:
                mean_dose = dose.GetDoseStatistic(RoiName=roi_name, DoseType='Average')
                if dose_unit == 'cGy':
                    string += f'#Mean dose: {int(round(mean_dose))}\n'
                elif dose_unit == 'Gy':
                    string += f'#Mean dose: {(mean_dose / 100):.2f}\n'

            # Compute doses and volumes for the geometry
            if dose_interval is not None:  # Use dose interval
                # Doses from max dose to zero
                dose_vals = np.arange(max_dose, -dose_interval, -dose_interval)
                dose_vals = np.array([max(0, val) for val in dose_vals])
                
                # Get dose values to pass to GetRelativeVolumeAtDoseValues
                # Converted to Gy if necessary
                if dose_unit == 'Gy':
                    dose_vals_arg = dose_vals * 100
                else:
                    dose_vals_arg = dose_vals

                # Volume at each dose
                vol_vals = dose.GetRelativeVolumeAtDoseValues(RoiName=roi_name, DoseValues=dose_vals_arg)
                
                # Convert relavtive volumes (proportions) to % or cc
                if vol_unit == '%':
                    vol_vals *= 100
                else:  # cc
                    vol_vals *= total_vol
            else:  # Use volume interval
                if vol_unit == '%':
                    vol_vals = np.arange(100, -vol_interval, -vol_interval)  # From max % volume (100%) to zero
                    vol_vals_arg = vol_vals / 100  # Convert to relative volume
                else:  # cc
                    # From max volume (whole geometry) to zero
                    vol_vals = np.arange(total_vol, -vol_interval, -vol_interval)
                    vol_vals = np.array([max(0, val) for val in vol_vals])
                    vol_vals_arg = vol_vals / total_vol  # Convert to relative volume

                # Dose at each relative volume
                dose_vals = dose.GetDoseAtRelativeVolumes(RoiName=roi_name, RelativeVolumes=vol_vals_arg)
                
                # Convert dose to Gy if necessary
                if dose_unit == 'Gy':
                    dose_vals /= 100

            string += '\n'.join(f'{dose_val:.3f}\t{vol_val:.3f}' for dose_val, vol_val in zip(dose_vals, vol_vals)) + '\n\n'  # E.g., "2000.321   10.780"

        # Delete unnecessary ROIs
        dg_box.DeleteRoi()
        temp.DeleteRoi()   

        # Create output directory if it doesn't exist
        if not os.path.isdir(OUTPUT_DIR):
            os.mkdir(OUTPUT_DIR)

        # Create and write to output file(s)
        # Output file name is "<dose type> dose <timestamp>.<extension>"
        filename = re.sub(r'[<>:"/\\\|\?\*]', '_', dose_txt) + ' ' + timestamp()
        for ext in file_exts:
            filepath = os.path.join(OUTPUT_DIR, filename + ext.lower())
            with open(filepath, 'w') as f:
                f.write(string)
