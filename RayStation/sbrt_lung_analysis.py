import clr
from collections import OrderedDict
from datetime import datetime
import random
import re
import sys
from typing import Dict, List, Optional

import pandas as pd  # Interpolation data from RTOG 0813 is read in as a DataFrame

from connect import *  # Interact w/ RS
from connect.connect_cpython import PyScriptObject

# Report uses ReportLab to create a PDF
from reportlab.lib.colors import Color, obj_R_G_B, toColor, black, blue, grey, green, lightgrey, orange, red, white, yellow
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.platypus.flowables import KeepTogether
from reportlab.platypus.tables import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

from scipy import interpolate  # Interpolate stats for a given PTV volume

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System import EventArgs
from System.Drawing import *
from System.Windows.Forms import *


class SbrtLungConstants(object):
    """Class that defines several useful constants for this remainder of the script"""

    # ----------------- Clinic-specific constants. CHANGE THESE! ----------------- #

    # Absolute path to the directory in which to create the PDF report
    # This directory does not have to already exist
    OUTPUT_DIR = os.path.join('T:', os.sep, 'Physics', 'Scripts', 'Output Files', 'SBRT Lung Analysis')

    # Filepath to the spreadsheet that does some of what this script does
    # This script uses the first table ("Theoretical") in that spreadsheet
    RTOG_INPUT_FILEPATH = os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-spreadsheets', 'RTOG 0813 Interpolation.xlsx')

    # --------------------------- No changes necessary --------------------------- #

    STATS = ['Traditional CI (R100%)', 'Paddick CI', 'GI (R50%)', 'D2cm [%]', 'Max dose @ appreciable volume [%]', 'V20Gy [%]', 'Max dose to External is inside PTV']  # Options for plan stats
    INTERP_STATS = ['GI (R50%)', 'D2cm [%]', 'V20Gy [%]']  # Plan stats that will be interpolated using the RTOG 0813 table

    STYLES = getSampleStyleSheet()  # Base styles (e.g., 'Heading1', 'Normal')
    
    # Paths to Adobe Reader on RS servers
    ADOBE_READER_PATHS = [os.path.join('C:', os.sep, 'Program Files (x86)', 'Adobe', 'Reader 11.0', 'Reader', 'AcroRd32.exe'), os.path.join('C:', os.sep, 'Program Files (x86)', 'Adobe', 'Acrobat Reader DC', 'Reader', 'AcroRd32.exe')]


def doses_on_addl_set(case: PyScriptObject, exam_names: List[str]) -> Dict[str, PyScriptObject]:
    """Gets all doses on additional set in the given case

    Arguments
    ---------
    case: The case whose doses on additional set to return
    exam_names: List of names on examinations that doses are on

    Returns
    -------
    A dictionary with dose names as keys and doses as values
    Dose names are in the format "<dose name> on <exam name>"
    """
    eval_doses = {}
    for fe in case.TreatmentDelivery.FractionEvaluations:
        for doe in fe.DoseOnExaminations:
            exam_name = doe.OnExamination.Name
            if exam_name in exam_names:
                for de in doe.DoseEvaluations:
                    if not de.Name and hasattr(de, 'ForBeamSet'):  # Ensure it's a dose on additional set
                        de_name = f'{de.ForBeamSet.DicomPlanLabel} on {exam_name}'
                        eval_doses[de_name] = de
    return eval_doses


def get_text_color(bkgrd_color: Color) -> Color:
    """Computes the appropriate text color (black or white) from the given background color

    Argument
    --------
    bkgrd_color: The background color

    Returns
    -------
    The black or white Color object
    """
    r, g, b = obj_R_G_B(bkgrd_color)
    return white if r * 0.00117 + g * 0.0023 + b * 0.00045 <= 0.72941 else black


def create_iso_roi(case: PyScriptObject, dose_dist: PyScriptObject, idl: float) -> PyScriptObject:
    """Creates an isodose line (IDL) ROI

    Arguments
    ---------
    case: The case in which to create the ROI
    dose_dist: The dose distribution from which to create the ROI
    idl: The dose level, in cGy

    Returns
    -------
    The IDL ROI
    """
    iso_roi_name = case.PatientModel.GetUniqueRoiName(DesiredName=f'IDL_{idl:.0f}')
    iso_roi = case.PatientModel.CreateRoi(Name=iso_roi_name, Type='Control')
    iso_roi.CreateRoiGeometryFromDose(DoseDistribution=dose_dist, ThresholdLevel=idl)
    dose_dist.UpdateDoseGridStructures()  # Must run to update new geometry
    return iso_roi


class SbrtLungAnalysisForm(Form):
    """Form that allows the user to choose settings for generation of a PDF report w/ RTOG 0813 statistics

    From a GUI, the user chooses the PTV, the stats to calculate, and the plan(s) and/or dose(s) on additional set for which to compute the stats 

    Plan / eval dose checkboxes are disabled with a highlighted message if the selected PTV is empty on their exams or if interpolated stats are selected and the PTV volume on their exams is outside the range in the RTOG 0813 interpolation table
    """
    def __init__(self, case: PyScriptObject, min_ptv_vol: float, max_ptv_vol: float, ptv_names: List[str], default_ptv_name: str, stats: List[str], default_stats: List[str], plan_names: List[str], default_plan_names: List[str], eval_doses: List[str], default_eval_doses: List[str]) -> None:
        """Initializes an SbrtLungAnalysisForm object

        Arguments
        ---------
        case: The case that the plans / eval doses belong to
        min_ptv_vol: The minimum PTV volume for which interpolated stats can be computed
        max_ptv_vol: The maximum PTV volume for which interpolated stats can be computed
        ptv_names: List of all PTV name choices
        default_ptv_name: PTV name that is checked by default
        stats: List of all stat choices
        default_stats: List of stats to check by default
        plan_names: List of all plan name choices
        default_plan_names: List of plan names to check by default
        eval_doses: List of names of choices of doses on additional set
        default_eval_doses: List of names of doses on additional set, to check by default
        """
        self._case = case
        self._min_ptv_vol, self._max_ptv_vol = min_ptv_vol, max_ptv_vol

        self._interp_stats_cked = any(stat in SbrtLungConstants.INTERP_STATS for stat in stats)  # "Helper" attribute. Do not restrict PTV volumes to between min_ptv_vol and max_ptv_vol if no interpolated stats are checked

        # Public attributes for main `sbrt_lung_analysis` function to access
        self.ptv_name = default_ptv_name
        self.stats = default_stats
        self.plan_names = default_plan_names
        self.eval_doses = default_eval_doses

        self._set_up_form()

        y = 15  # y-coordinate of next control

        # PTV radiobuttons
        self._ptv_gb = GroupBox()
        self._ptv_gb.AutoSize = True  # Enlarge GroupBox to accommodate radiobutton widths, if necessary
        self._ptv_gb.Location = Point(15, y)
        self._ptv_gb.Text = 'PTV:'
        rb_y = 15  # Vertical coordinate of radio button inside groupbox
        for ptv_name in ptv_names:
            rb = RadioButton()
            rb.AutoSize = True  # Enlarge RadioButton to accommodate text width, if necessary
            rb.Click += self._ptv_rb_Click
            rb.Checked = ptv_name == default_ptv_name  # By default, check only the default PTV name
            rb.Location = Point(15, rb_y)
            rb.Text = ptv_name
            self._ptv_gb.Controls.Add(rb)
            rb_y += rb.Height
        self._ptv_gb.Height = len(ptv_names) * 20  # Size groupbox to radio buttons
        self.Controls.Add(self._ptv_gb)
        y += self._ptv_gb.Height + 15
        
        # Add stats checkboxes
        self._stats_gb = GroupBox()
        self._stats_gb.AutoSize = True
        self._stats_gb.Location = Point(15, y)
        self._stats_gb.Text = 'Stat(s):'
        cb_y = 15  # Vertical coordinate of checkbox inside groupbox
        cb = CheckBox()
        cb.AutoSize = True
        cb.Text = 'Select all'
        cb.Location = Point(15, cb_y)
        cb.ThreeState = True
        if len(default_stats) == len(stats):  # All stats checked by default
            cb.CheckState = cb.Tag = CheckState.Checked  # Tag attribute is used to keep track of previous check state, so next check state can be determined
        elif not stats:  # No stats checked by default
            cb.CheckState = cb.Tag = CheckState.Unchecked
        else:  # Some, but not all, stats checked by default
            cb.CheckState = cb.Tag = CheckState.Indeterminate
        cb.Click += self._select_all_cb_Click
        self._stats_gb.Controls.Add(cb)
        cb_y += cb.Height
        for stat in stats:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Location = Point(30, cb_y)  # Indent with resoect to "Select all" checkbox
            cb.Checked = stat in default_stats
            cb.Click += self._item_cb_Click
            cb.Text = stat
            self._stats_gb.Controls.Add(cb)
            cb_y += cb.Height
        self._stats_gb.Height = len(default_stats) * 20
        self.Controls.Add(self._stats_gb)
        y += self._stats_gb.Height + 15

        # Associate plan / eval dose checkboxes with labels that display a message if they are disabled due to empty PTV geometry or PTV volume unsupported by interpolation
        self._ptv_vol_msg_lbls = {}

        # Add plans checkboxes
        self._plans_gb = GroupBox()
        self._plans_gb.AutoSize = True
        self._plans_gb.Location = Point(15, y)
        self._plans_gb.Text = 'Plan(s):'
        cb_y = 15
        cb = CheckBox()
        cb.AutoSize = True
        cb.Text = 'Select all'
        cb.Location = Point(15, cb_y)
        cb.ThreeState = True
        if len(default_plan_names) == len(plan_names):
            cb.CheckState = cb.Tag = CheckState.Checked  # Tag attribute is used to keep track of previous check state, so next check state can be determined
        elif not default_plan_names:
            cb.CheckState = cb.Tag = CheckState.Unchecked
        else:
            cb.CheckState = cb.Tag = CheckState.Indeterminate
        cb.Click += self._select_all_cb_Click
        self._plans_gb.Controls.Add(cb)
        cb_y += cb.Height
        for plan_name in plan_names:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Location = Point(30, cb_y)
            cb.Checked = plan_name in default_plan_names
            cb.Click += self._item_cb_Click
            cb.Text = plan_name
            self._plans_gb.Controls.Add(cb)
            cb_y += cb.Height
            self._ptv_vol_msg_lbls[cb] = self._create_ptv_vol_msg_lbl(cb)
        self._plans_gb.Height = len(plan_names) * 20
        self.Controls.Add(self._plans_gb)
        y += self._plans_gb.Height + 15

        # Add eval doses (doses on additional set) checkboxes
        if eval_doses:  # There is dose computed on additional set
            # Add eval dose checkboxes
            self._eval_doses_gb = GroupBox()
            self._eval_doses_gb.AutoSize = True
            self._eval_doses_gb.Location = Point(15, y)
            self._eval_doses_gb.Text = 'Dose(s) on additional set:'
            cb_y = 15
            cb = CheckBox()
            cb.AutoSize = True
            cb.Text = 'Select all'
            cb.Location = Point(15, cb_y)
            cb.ThreeState = True
            if len(default_eval_doses) == len(eval_doses):
                cb.CheckState = cb.Tag = CheckState.Checked  # Tag attribute is used to keep track of previous check state, so next check state can be determined
            elif not default_eval_doses:
                cb.CheckState = cb.Tag = CheckState.Unchecked
            else:
                cb.CheckState = cb.Tag = CheckState.Indeterminate
            cb.Click += self._select_all_cb_Click
            self._eval_doses_gb.Controls.Add(cb)
            cb_y += cb.Height
            for name in eval_doses:
                cb = CheckBox()
                cb.AutoSize = True
                cb.Location = Point(30, cb_y)
                cb.Checked = name in default_eval_doses
                cb.Click += self._item_cb_Click
                cb.Text = name
                self._eval_doses_gb.Controls.Add(cb)
                cb_y += cb.Height
                self._ptv_vol_msg_lbls[cb] = self._create_ptv_vol_msg_lbl(cb)
            self._eval_doses_gb.Height = len(eval_doses) * 20
            self.Controls.Add(self._eval_doses_gb)
            y += self._eval_doses_gb.Height + 15

        # Add "Generate Report" button
        self._generate_btn = Button()
        self._generate_btn.AutoSize = True
        self._generate_btn.Click += self._generate_btn_Click
        self._generate_btn.Location = Point(15, y)
        self._generate_btn.Text = 'Generate Report'
        self.AcceptButton = self._generate_btn  # Click button when "Enter" is pressed
        self.Controls.Add(self._generate_btn)

        self._ptv_rb_Click()  # Simulate clicking a PTV radio button
        self._set_generate_enabled()

    # ------------------------------ Event handlers ------------------------------ #

    def _ptv_rb_Click(self, sender: Optional[RadioButton] = None, event: Optional[EventArgs] = None) -> None:
        """Handles the event of clicking a PTV radio button

        Enables/disables plan and eval dose checkboxes based on whether the PTV has geometries on their exams

        `sender` and `event` args are optional because this method may or may not be called as an event handler
        """
        ptv_name = next(rb.Text for rb in self._ptv_gb.Controls if rb.Checked)  # The selected PTV name

        # Plans
        cbs = [cb for cb in self._plans_gb.Controls if isinstance(cb, CheckBox) and cb.Text != 'Select all']
        for cb in cbs:  # Iterate over all plan checkboxes that aren't "Select all"
            ptv_geom = self._case.TreatmentPlans[cb.Text].GetTotalDoseStructureSet().RoiGeometries[ptv_name]
            self._set_ptv_vol_msg(cb, ptv_geom)
        
        # Eval doses
        if hasattr(self, '_eval_doses_gb'):
            cbs = [cb for cb in self._eval_doses_gb.Controls if isinstance(cb, CheckBox) and cb.Text != 'Select all']
            for cb in cbs:
                exam_name = cb.Text.split(' on ')[1]  # Checkbox text is "<dose name> on <exam name>"
                ptv_geom = self._case.PatientModel.StructureSets[exam_name].RoiGeometries[ptv_name]
                self._set_ptv_vol_msg(cb, ptv_geom)

    def _select_all_cb_Click(self, sender: CheckBox, event: EventArgs) -> None:
        """Handles the event of clicking a "Select all" checkbox

        Sets new "Select all" checkstate based on previous checkstate
        Sets item checkbox states based on new "Select all" checkbox state
        """
        cbs = [cb for cb in sender.Parent.Controls if isinstance(cb, CheckBox) and cb.Text != 'Select all']
        if sender.Parent.Text == 'Stat(s):':
            self._interp_stats_cked = any(cb.Text in SbrtLungConstants.INTERP_STATS for cb in cbs if cb.Checked)
            self._ptv_rb_Click()
        if sender.Tag == CheckState.Checked:  # If checked, uncheck
            sender.CheckState = sender.Tag = CheckState.Unchecked
            for cb in cbs: 
                cb.Checked = False
        else:  # If unchecked or indeterminate, check and check all enabled item checkboxes
            for cb in cbs:
                cb.Checked = cb.Enabled
            if all(cb.Checked for cb in cbs):
                sender.CheckState = sender.Tag = CheckState.Checked   
            else:
                sender.CheckState = sender.Tag = CheckState.Indeterminate
        self._set_generate_enabled()  # "Generate Reports" enabled based on item selections

    def _item_cb_Click(self, sender: CheckBox, event: Optional[EventArgs] = None) -> None:
        """Handles the event that is clicking a non-"Select all" checkbox

        Sets the checkstate of the appropriate "Select all" checkbox based on the checkstates of all its item checkboxes
        Enables the "Select all" checkbox only if all item checkboxes are enabled

        `event` arg is optional because this method may or may not be called as an event handler
        """
        select_all = next(cb for cb in sender.Parent.Controls if isinstance(cb, CheckBox) and cb.Text == 'Select all')
        cbs = [cb for cb in sender.Parent.Controls if isinstance(cb, CheckBox) and cb.Text != 'Select all']  # Non-"Select all" checkboxes
        if sender.Parent.Text == 'Stat(s):':
            self._interp_stats_cked = any(cb.Text in SbrtLungConstants.INTERP_STATS for cb in cbs if cb.Checked)
            self._ptv_rb_Click()
        enabled_cbs = [cb for cb in cbs if cb.Enabled]
        enabled_cbs_cked = [cb.Checked for cb in enabled_cbs]
        select_all.Enabled = bool(enabled_cbs)  # Enable "Select all" only if all item checkboxes are enabled
        if not select_all.Enabled or not any(enabled_cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Unchecked
        elif all(enabled_cbs_cked):  # All item checkboxes are checked
            select_all.CheckState = select_all.Tag = CheckState.Checked
        else:  # Some, but not all, enabled item checkboxes are checked
            select_all.CheckState = select_all.Tag = CheckState.Indeterminate
        self._set_generate_enabled()

    def _set_generate_enabled(self) -> None:
        """Enables or disables the "Generate Report" button

        Enables the button only if a PTV is checked, at least one stat is checked, and at least one plan or eval dose is checked

        Sets public attributes for main `sbrt_lung_analysis` function to access
        """
        try:
            self.ptv_name = next(cb.Text for cb in self._ptv_gb.Controls if cb.Checked)
        except StopIteration:
            self.ptv_name = None
        self.stats = [cb.Text for cb in self._stats_gb.Controls if cb.Text != 'Select all' and cb.Checked]
        self.plan_names = [cb.Text for cb in self._plans_gb.Controls if isinstance(cb, CheckBox) and cb.Text != 'Select all' and cb.Checked]
        if hasattr(self, '_eval_doses_gb'):
            self.eval_dose_names = [cb.Text for cb in self._eval_doses_gb.Controls if isinstance(cb, CheckBox) and cb.Text != 'Select all' and cb.Checked]
        else:
            self.eval_dose_names = []
        self._generate_btn.Enabled = self.ptv_name is not None and self.stats and (self.plan_names or self.eval_dose_names)

    def _generate_btn_Click(self, sender: Button, event: EventArgs) -> None:
        """Handles the event of clicking the "Generate Report" button

        Sets public attributes for main `sbrt_lung_analysis` function to access
        """
        self.DialogResult = DialogResult.OK

    # ------------------------------- Other methods ------------------------------ #

    def _set_up_form(self) -> None:
        """Styles the Form"""
        self.Text = 'SBRT Lung Analysis'

        # Adapt form size to controls
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # Disallow user resizing
        self.StartPosition = FormStartPosition.CenterScreen  # Launch window in middle of screen

    def _create_ptv_vol_msg_lbl(self, cb: CheckBox) -> None:
        """Creates a label to go beside a plan / eval dose checkbox

        The label is to hold "highlighted" messages about unsupported PTV volumes

        Argument
        --------
        cb: The checkbox to associate the label with
        """
        l = Label()

        # Adapt size to text
        l.AutoSize = True
        l.AutoSizeMode = AutoSizeMode.GrowAndShrink

        l.BackColor = Color.Yellow  # "Highlight" the text
        l.Location = Point(cb.Location.X + cb.Width, cb.Location.Y)  # Immediately to the right of the checkbox in the groupbox
        l.Text = ''  # Start with no volume message. Text is set in constructor's call to _ptv_rb_Click
        cb.Parent.Controls.Add(l)
        
        return l

    def _set_ptv_vol_msg(self, cb: CheckBox, ptv_geom: PyScriptObject) -> None:
        """Sets the text of the label beside each plan / eval dose checkbox, and enables or disables the checkbox, based on whether the selected PTV on its exam has the appropriate volume

        An unsupported PTV volume is an empty geometry or, if any interpolated stats are checked, a PTV volume oytside the min and max supported volumes

        Arguments
        ---------
        cb: The plan or eval dose checkbox
        ptv_geom: The geometry of the selected PTV on the dose's exam
        """
        msg = ''  # Assume PTV volume OK. This text will effectively hide the label
        if not ptv_geom.HasContours():  # Empty geometry is a problem for all stats, not just interpolated ones
            msg = 'Empty PTV geometry'
        elif self._interp_stats_cked:  # Volume restrictions only apply if interpolation is required
            ptv_vol = ptv_geom.GetRoiVolume()
            if ptv_vol < self._min_ptv_vol:  # Volume too small
                msg = f'PTV volume = {ptv_vol:.2f} < {self._min_ptv_vol:.1f} cc'
            elif ptv_vol > self._max_ptv_vol:  # Volume too large
                msg = f'PTV volume = {ptv_vol:.2f} > {self._max_ptv_vol:.1f} cc'
        self._ptv_vol_msg_lbls[cb].Text = msg  # Set the text of the checkbox's label
        if msg:  # Disable and uncheck the checkbox if there is a volume message
            cb.Enabled = cb.Checked = False
        else:
            cb.Enabled = True
        self._item_cb_Click(cb)  # Simulate clicking the plan / eval dose checkbox so that the "Select all" checkbox can be enabled/disabled and checked/unchecked


def format_name(name: str) -> str:
    """Converts the patient's Name attribute into a better format for display

    Arguments
    ---------
    pt: The patient whose name to format

    Returns
    -------
    The formatted patient name

    Example
    -------
    Given some_pt with Name '^Jones^Bill^^M':
    format_pt_name(some_pt) -> 'Jones, Bill M'
    """
    parts = [part for part in re.split(r'\^+', name) if part != '']
    name = parts[0]
    if len(parts) > 0:
        name += ', ' + ' '.join(parts[1:])
    return name


def sbrt_lung_analysis_() -> None:
    """Generates and opens a PDF report of selected RTOG 0813 statistics for SBRT lung plan(s) and/or dose(s) on additional set

    From a GUI, the user chooses:
        - PTV (default is PTV in primary Rx of current beam set) 
        - RTOG 0813 statistics (default is all)
        - Plan(s) (default is current plan) (External, E-PTV_Ev20, and Lungs-CTV must be contoured on the planning exam)
        - Doses on additional sets, if present (default is none) (External, E-PTV_Ev20, and Lungs-CTV must be contoured on the additional set)
    """
    filepath = sbrt_lung_analysis(gui=True)

    # Open report
    for reader_path in SbrtLungConstants.ADOBE_READER_PATHS:
        try:
            os.system(f'START /B "{reader_path}" "{filepath}"')
            break
        except:
            continue
        

def sbrt_lung_analysis(**kwargs) -> str:
    """Generates a PDF report with selected RTOG 0813 statistics for the selected plan(s) and/or dose(s) on additional set

    If a GUI is used, displays errors in a dialog box and aborts script.
    Otherwise, returns the error message when it is generated, or the filepath to the report if there are no errors

    Keyword Arguments
    -----------------
    gui: bool
        True if the user should choose settings from a GUI, False otherwise
        Defaults to False
    ptv_name: str
        Name of the PTV to use for statistic calculations
        If `gui` is True, the name of the default checked PTV. User chooses from PTVs with geometries on any of the planning / eval dose exams with volumes between the min and max PTV volumes in the RTOG 0813 stats interpolation table
        Defaults to the PTV that the current beam set's primary Rx is to
    stats: List[str]
        List of statistic names to compute
        If `gui` is True, the default checked stats
        Choose from the following:
            - "Traditional CI (R100%)"
            - "Paddick CI"
            - "GI (R50%)"
            - "D2cm [%]"
            - "Max dose @ appreciable volume [%]"
            - "V20Gy [%]"
            - "Max dose to External is inside PTV"
        Defaults to all of the above
    plan_names: List[str]
        List of plan names for which to compute the stats
        If `gui` is True, the default checked plan names
        Defaults to all plan names with planning exams with normal tissue ("E-PTV_Ev20"), external, "Lungs-CTV"/"Lungs-ITV", and PTV geometries
    eval_dose_names: List[str]
        List of doses on additional set for which to compute the stats
        If `gui` is True, the default checked eval dose names
        Each eval dose name is in the format "<beam set name> on <exam name>" (e.g., "SBRT Rt Lung on Inspiration 12.9.20")
        Defaults to all eval doses on exams with normal tissue ("E-PTV_Ev20"), external, and "Lungs-CTV"/"Lungs-ITV", and PTV geometries

    PDF Report contains:
    - Table of color-coded computed stats for each selected plan and eval dose
    - If any interpolated stats were selected, table of RTOG 0813 interpolation stats with a row added for each selected plan and eval dose
    """
    # Get current variables
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()  # Exit script
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit() 
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
        sys.exit()
    try:
        beam_set = get_current('BeamSet')
    except:
        MessageBox.Show('There are no beam sets in the current plan. Click OK to abort the script.', 'No Beam Sets')
        sys.exit()  
    exam = plan.GetTotalDoseStructureSet().OnExamination

    gui = kwargs.get('gui', False)  # Are we using a GUI?

    # Report name
    if not os.path.isdir(SbrtLungConstants.OUTPUT_DIR):
        os.makedirs(SbrtLungConstants.OUTPUT_DIR)
    pt_name = format_name(patient.Name)
    filename = pt_name + ' ' + plan.Name + ' ' + datetime.now().strftime('%Y-%m-%d %H_%M_%S') + '.pdf'
    filepath = os.path.join(SbrtLungConstants.OUTPUT_DIR, re.sub(r'[<>:"/\\\|\?\*]', '_', filename))
    
    # Ensure Rx exists
    rx = beam_set.Prescription.PrimaryPrescriptionDoseReference
    if rx is None:
        msg = 'Beam set has no prescription.'
        if not gui:
            return msg
        MessageBox.Show(msg, 'No Prescription')
        sys.exit()

    # Ensure beam set is SBRT lung
    fx_pattern = beam_set.FractionationPattern
    if beam_set.Modality != 'Photons' or fx_pattern is None or fx_pattern.NumberOfFractions > 5 or rx.DoseValue < 600 * fx_pattern.NumberOfFractions or case.BodySite not in ['', 'Thorax']:
        MessageBox.Show('This is not an SBRT lung plan. Click OK to abort the script.', 'Not SBRT Lung')
        sys.exit()

    # Read in data
    # RTOG 0813 stats for interpolation. Read only the first table
    data = pd.read_excel(SbrtLungConstants.RTOG_INPUT_FILEPATH, engine='openpyxl', usecols='A:G')  
    min_ptv_vol, max_ptv_vol = data['PTV vol [cc]'].min(), data['PTV vol [cc]'].max()

    # Does the case contain the necessary ROIs?
    roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
    if 'E-PTV_Ev20' not in roi_names or ('Lungs-CTV' not in roi_names and 'Lungs-ITV' not in roi_names):
        msg = 'Case is missing ROI(s): E-PTV_Ev20 and/or Lungs-CTV/Lungs-ITV.'
        if not gui:
            return msg
        MessageBox.Show(msg, 'Missing ROI(s)')
        sys.exit()

    # Does the case have an external ROI?
    try:
        ext = next(roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == 'External')
    except:
        msg = 'Case has no external ROI.'
        if not gui:
            return msg
        MessageBox.Show(msg, 'No External ROI')
        sys.exit()

    # All PTV names
    all_ptv_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type.upper() == 'PTV']
    if not all_ptv_names:
        msg = 'There are no PTVs in the current case.'
        if not gui:
            return msg
        MessageBox.Show(msg, 'No PTVs')
        sys.exit()

    # All PTV names with volume in the range supported by the RTOG 0813 interpolation table, on at least one exam
    # Exams w/ the necessary geometries (necessary to select plans and eval doses)
    exam_names = []
    viable_ptv_names = []  # PTVs with geometries on at least one viable exam
    for e in case.Examinations:
        roi_geom_names = [geom.OfRoi.Name for geom in case.PatientModel.StructureSets[e.Name].RoiGeometries if geom.HasContours()]
        viable_ptv_names_exam = [ptv_name for ptv_name in all_ptv_names if ptv_name in roi_geom_names]
        if viable_ptv_names_exam and 'E-PTV_Ev20' in roi_geom_names and ('Lungs-CTV' in roi_geom_names or 'Lungs-ITV' in roi_geom_names) and ext.Name in roi_geom_names:
            exam_names.append(e.Name)
            viable_ptv_names.extend(viable_ptv_names_exam)
    viable_ptv_names = sorted(list(set(viable_ptv_names)), key=lambda x: x.lower())  # Remove duplicate PTV names and sort alphabetically (case insensitive)

    # Are there any PTV geometries?
    if not viable_ptv_names:
        msg = 'There are no exams with a geometry for each E-PTV_Ev20, External, either Lungs-CTV or Lungs-ITV, and at least one PTV.'
        if not gui:
            return msg
        MessageBox.Show(msg, 'No Viable Exams')
        sys.exit()
    
    # Default PTV name
    # If not provided, use the target of the current beam set's primary Rx if it has geometries
    # If this volume is inappropriate, use random PTV with geometries
    ptv_name = kwargs.get('ptv_name', rx.OnStructure.Name if rx.OnStructure.Name in viable_ptv_names else random.choice(viable_ptv_names))

    # Default plan stats
    stats = kwargs.get('stats', SbrtLungConstants.STATS)

    # All plan names
    all_plan_names = [p.Name for p in case.TreatmentPlans if p.TreatmentCourse is not None and p.TreatmentCourse.TotalDose is not None and p.TreatmentCourse.TotalDose.DoseValues is not None and p.GetTotalDoseStructureSet().OnExamination.Name in exam_names]
    if not all_plan_names:
        msg = 'There are no plans that have dose, and a geometry for each E-PTV_Ev20, External, either Lungs-CTV or Lungs-ITV, and at least one PTV.'
        if not gui:
            return msg
        MessageBox.Show(msg, 'No Viable Plans')
        sys.exit()

    # Default plan names
    if 'plan_names' in kwargs:
        plan_names = [plan_name for plan_name in kwargs.get('plan_names') if plan_name in all_plan_names]
    else:
        plan_names = all_plan_names[:]

    # All eval doses
    all_eval_doses = doses_on_addl_set(case, exam_names)

    # Default eval doses
    if 'eval_dose_names' in kwargs:
        eval_doses = {name: all_eval_doses[name] for name in kwargs.get('eval_dose_names') if name in all_eval_doses}
    else:
        eval_doses = all_eval_doses.copy()

    # Display GUI if applicable
    if gui:
        form = SbrtLungAnalysisForm(case, min_ptv_vol, max_ptv_vol, viable_ptv_names, ptv_name, SbrtLungConstants.STATS, stats, all_plan_names, plan_names, list(all_eval_doses.keys()), list(eval_doses.keys()))
        form.ShowDialog()
        if form.DialogResult != DialogResult.OK:  # User exited GUI
            sys.exit()
        ptv_name = form.ptv_name
        stats = form.stats
        plan_names = form.plan_names
        eval_doses = {eval_dose_name: all_eval_doses[eval_dose_name] for eval_dose_name in form.eval_dose_names} if all_eval_doses else {}

    ## Prepare PDF
    pdf = SimpleDocTemplate(filepath, pagesize=landscape(letter), bottomMargin=0.2 * inch, leftMargin=0.25 * inch, rightMargin=0.2 * inch, topMargin=0.2 * inch)  # 8.5 x 11", 0.2" top and bottom margin, 0.25" left and right margin

    # Heading
    #hdr = Paragraph(pt_name, style=hdg)
    hdr = Paragraph(pt_name, style=SbrtLungConstants.STYLES['Heading1'])
    mrn = Paragraph('MRN: ' + patient.PatientID, style=SbrtLungConstants.STYLES['Heading2'])
    desc = Paragraph('RTOG 0813 SBRT Lung Analysis', style=SbrtLungConstants.STYLES['Heading2'])

    # Plan/Eval dose key is a ReportLab Table
    # A single row: orange square, "Plan" text, blue square, "Evaluation Dose" text
    # Include each key item only if some of that items are selected
    key_1_data, key_1_style = [], []
    if plan_names:  # Some plans are selected
        key_1_data.extend(['', 'Plan'])
        key_1_style.extend([('BACKGROUND', (0, 0), (0, 0), orange), ('BOX', (0, 0), (0, 0), 0.5, black)])  # 1st cell: orange background, black outline
        if eval_doses:  # Some eval doses are selected in addition to the plans
            key_1_data.extend(['', 'Evaluation dose']) 
            key_1_style.extend([('BACKGROUND', (2, 0), (2, 0), blue), ('BOX', (2, 0), (2, 0), 0.5, black)])  # 3rd cell: blue background, black outline
    elif eval_doses:  # No plans are selected, but some eval doses are
        key_1_data.extend(['', 'Evaluation dose'])
        key_1_style.extend([('BACKGROUND', (0, 0), (0, 0), blue), ('BOX', (0, 0), (0, 0), 0.5, black)])  # 1st cell: blue background, black outline
    key_1 = Table([key_1_data], colWidths=[0.2 * inch, 1.25 * inch] * 2, rowHeights=[0.2 * inch], style=TableStyle(key_1_style), hAlign='LEFT')  # Left-align the table

    # Deviation key is a ReportLab Table
    key_2_data = ['', 'No deviation', '', 'Minor deviation', '', 'Major deviation']  # A single row: green square, "No deviation" text, yellow square, "Minor deviation" text, red square, "Major deviation" text
    key_2_style = [
        # 1st cell: green background, black outline
        ('BACKGROUND', (0, 0), (0, 0), green),
        ('BOX', (0, 0), (0, 0), 0.5, black),
        # 3rd cell: yellow background, black outline
        ('BACKGROUND', (2, 0), (2, 0), yellow),
        ('BOX', (2, 0), (2, 0), 0.5, black),
        # 5th cell: red background, black outline
        ('BACKGROUND', (4, 0), (4, 0), red),
        ('BOX', (4, 0), (4, 0), 0.5, black)
    ]
    key_2 = Table([key_2_data], colWidths=[0.2 * inch, 1.25 * inch] * 3, rowHeights=[0.2 * inch], style=TableStyle(key_2_style), hAlign='LEFT')  # Left-align the table

    ## Plan stats table

    # Table title
    plan_stats_title = Paragraph('Plan Stats', style=SbrtLungConstants.STYLES['Heading3'])

    # Header row
    plan_stats_data = ['Plan / Evaluation dose', 'PTV vol [cc]'] + stats
    plan_stats_data = [re.sub(r'(\d+(%|cm|Gy))', '<sub>\g<1></sub>', text) for text in plan_stats_data]  # Surround each subscript with HTML <sub> tags
    plan_stats_data = [[Paragraph(text, style=SbrtLungConstants.STYLES['Heading4']) for text in plan_stats_data]]
    plan_stats_style = [  # Center-align, middle-align, and black outline for all cells. Gray background for header row
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
    ]

    ## Compute all PTV vols (necessary now so we know where in tables to place plan rows)
    # Dict of {vol : [plan / eval dose names]}
    ptv_vols = {}
    for plan_name in plan_names:
        ptv_vol = case.TreatmentPlans[plan_name].GetTotalDoseStructureSet().RoiGeometries[ptv_name].GetRoiVolume()
        if ptv_vol in ptv_vols:
            ptv_vols[ptv_vol].append(plan_name)
        else:
            ptv_vols[ptv_vol] = [plan_name]
    for name in eval_doses:
        exam_name = name.split(' on ')[1]
        ptv_vol = case.PatientModel.StructureSets[exam_name].RoiGeometries[ptv_name].GetRoiVolume()
        if ptv_vol in ptv_vols:
            ptv_vols[ptv_vol].append(name)
        else:
            ptv_vols[ptv_vol] = [name]
    ptv_vols = [(vol, sorted(ptv_vols[vol])) for vol in sorted(ptv_vols)]  # Sort PTV volumes in ascending order
    
    ## Colors
    # Equally spaced hues based on number of hues needed
    # Constrain to brighter hues so that color is obvious 
    r, g, b = obj_R_G_B(orange)  # Extract R, G, and B components of ReportLab orange
    a = [0.502 + (0.9 - 0.502) / len(plan_names) * (len(plan_names) - i - 1) for i in range(len(plan_names))]  # A components of colors
    plan_colors = [toColor(f'rgba({r}, {g}, {b}, {a_})') for a_ in a]  # Plan rows will be an orange hue

    r, g, b = obj_R_G_B(blue)
    a = [0.502 + (0.9 - 0.502) / len(eval_doses) * (len(eval_doses) - i - 1) for i in range(len(eval_doses))]  # A components of colors
    eval_dose_colors = [toColor(f'rgba({r}, {g}, {b}, {a_})') for a_ in a]  # Eval dose rows will be an blue hue

    ## Interpolated cutoffs table

    # Table title
    interp_cutoffs_title = Paragraph('Interpolated Cutoffs', style=SbrtLungConstants.STYLES['Heading3'])

    # Header row
    interp_cutoffs_data = ['PTV vol [cc]']
    for stat in SbrtLungConstants.INTERP_STATS:
        interp_cutoffs_data.extend([stat + ' None', stat + ' Minor'])
    interp_cutoffs_data = [[Paragraph(re.sub(r'(\d+(%|cm|Gy))', '<sub>\g<1></sub>', text), style=SbrtLungConstants.STYLES['Heading4']) for text in interp_cutoffs_data]]
    interp_cutoffs_style = plan_stats_style[:]  # Interp vals table has same style as plan stats table
    
    # Add plans and eval doses to plan stats table and interp cutoffs table
    plan_stats_idx = 0  # Row number in plan stats table
    # Process any PTV volumes between this interp cutoffs row and the next
    for i, row in data.iterrows():
        while ptv_vols:
            ptv_vol, names = ptv_vols[0]  # This PTV volume's list of plans / eval doses
            if ptv_vol > row['PTV vol [cc]']:
                break
            for name in names:  # Add rows for each plan / eval dose with that PTV volume
                plan_stats_idx += 1
                plan_stats_row = [name, round(ptv_vol, 2)]  # e.g., ['SBRT Lung', 38.89]
                interp_cutoffs_row = [Paragraph(str(round(ptv_vol, 2)), style=SbrtLungConstants.STYLES['Normal'])]  # e.g., [38.89]
            
                if name in plan_names:
                    color = plan_colors.pop()  # Get next plan row color
                    dose_dist = case.TreatmentPlans[name].TreatmentCourse.TotalDose
                    exam_name = case.TreatmentPlans[name].GetTotalDoseStructureSet().OnExamination.Name
                    rx = sum(bs.Prescription.PrimaryPrescriptionDoseReference.DoseValue for bs in case.TreatmentPlans[name].BeamSets if bs.Prescription.PrimaryPrescriptionDoseReference is not None)  # Sum of Rx's from all beam sets that have an Rx
                    v20_vol = 2000
                else:  # Eval dose
                    # Dose distribution is fractional!
                    color = eval_dose_colors.pop()  # Get next eval dose row color
                    dose_dist = eval_doses[name]
                    exam_name = name.split(' on ')[1]
                    rx = float(dose_dist.ForBeamSet.Prescription.PrimaryPrescriptionDoseReference.DoseValue) / dose_dist.ForBeamSet.FractionationPattern.NumberOfFractions
                    v20_vol = 2000.0 / dose_dist.ForBeamSet.FractionationPattern.NumberOfFractions
                struct_set = case.PatientModel.StructureSets[exam_name]
                text_color = get_text_color(color)  # Color text black or white (based on background color `color`)?

                plan_stats_style.append(('BACKGROUND', (0, plan_stats_idx), (1, plan_stats_idx), color))
                bk_color = None  # Background color for the individual stat value

                # Compute each stat and add appropriately colored cell to plan stats table
                for j, stat in enumerate(stats):
                    if stat == 'Traditional CI (R100%)':
                        iso_roi = create_iso_roi(case, dose_dist, rx)  # Create ROI from 100% isodose
                        
                        plan_stat_val = struct_set.RoiGeometries[iso_roi.Name].GetRoiVolume() / ptv_vol  # Volume of 100% isodose geometry as proportion of PTV volume
                        plan_stat_val = round(plan_stat_val, 2)  # Round CI to 2 decimal places
                        bk_color = yellow if plan_stat_val < 1 else green if plan_stat_val <= 1.2 else yellow if plan_stat_val <= 1.5 else red
                        
                        iso_roi.DeleteRoi()

                    elif stat == 'Paddick CI':
                        iso_roi = create_iso_roi(case, dose_dist, rx)  # Create ROI from 100% isodose
                        # Create intersection of PTV and 100% isodose
                        intersect_roi_name = case.PatientModel.GetUniqueRoiName(DesiredName=f'PTV&IDL_{rx:.0f}')
                        intersect_roi = case.PatientModel.CreateRoi(Name=intersect_roi_name, Type='Control')
                        intersect_roi.CreateAlgebraGeometry(Examination=case.Examinations[exam_name], ExpressionA={ 'Operation': 'Union', 'SourceRoiNames': [ptv_name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [iso_roi.Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation='Intersection', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

                        iso_roi_vol = struct_set.RoiGeometries[iso_roi.Name].GetRoiVolume()
                        intersect_roi_vol = struct_set.RoiGeometries[intersect_roi.Name].GetRoiVolume()
                        
                        plan_stat_val = intersect_roi_vol * intersect_roi_vol / (ptv_vol * iso_roi_vol)
                        plan_stat_val = round(plan_stat_val, 2)  # Round CI to 2 decimal places
                        bk_color = yellow if plan_stat_val < 1 else green if plan_stat_val <= 1.2 else yellow if plan_stat_val <= 1.5 else red
                        
                        intersect_roi.DeleteRoi()
                        iso_roi.DeleteRoi()
                    
                    elif stat == 'GI (R50%)':
                        iso_roi = create_iso_roi(case, dose_dist, 0.5 * rx)  # 50% isodose
                        
                        plan_stat_val = struct_set.RoiGeometries[iso_roi.Name].GetRoiVolume() / ptv_vol  # Volume of 50% isodose geometry as proportion of PTV volume
                        plan_stat_val = round(plan_stat_val, 2)  # Round GI to 2 decimal places
                        
                        iso_roi.DeleteRoi()

                    elif stat == 'D2cm [%]':
                        normal_tissue_vol = case.PatientModel.StructureSets[exam_name].RoiGeometries['E-PTV_Ev20'].GetRoiVolume()  # Volume of everything except 2-cm expansion of PTV
                        rel_vol = 0.035 / normal_tissue_vol  # Appreciable volume (0.035 cc) as proportion of normal tissue volume
                        dose_at_rel_vol = dose_dist.GetDoseAtRelativeVolumes(RoiName='E-PTV_Ev20', RelativeVolumes=[rel_vol])[0]

                        plan_stat_val = dose_at_rel_vol / rx * 100  # Dose at relative volume, as percent of Rx
                        plan_stat_val = round(plan_stat_val, 1)  #Round D2cm to a single decimal place

                    elif stat == 'Max dose @ appreciable volume [%]':
                        ext_vol = case.PatientModel.StructureSets[exam_name].RoiGeometries[ext.Name].GetRoiVolume()  # Volume of external
                        rel_vol = 0.035 / ext_vol  # Appreciable volume (0.035 cc) as proportion of external volume
                        dose_at_rel_vol = dose_dist.GetDoseAtRelativeVolumes(RoiName=ext.Name, RelativeVolumes=[rel_vol])[0]
                        
                        plan_stat_val = dose_at_rel_vol / rx * 100  # Dose at relative volume, as percent of Rx
                        plan_stat_val = round(plan_stat_val, 1)  # Round to a single decimal place
                        bk_color = yellow if plan_stat_val < 123 else green if plan_stat_val <= 130 else yellow if plan_stat_val <= 135 else red

                    elif stat == 'V20Gy [%]':
                        try:
                            rel_vol = dose_dist.GetRelativeVolumeAtDoseValues(RoiName='Lungs-CTV', DoseValues=[v20_vol])[0]  # Proportion of Lungs-CTV that receives 2000 cGy total dose
                        except:
                            rel_vol = dose_dist.GetRelativeVolumeAtDoseValues(RoiName='Lungs-ITV', DoseValues=[v20_vol])[0]  # Proportion of Lungs-CTV that receives 2000 cGy total dose
                        
                        plan_stat_val = rel_vol * 100  # Volume as a percent
                        plan_stat_val = round(plan_stat_val, 2)  # Round V20Gy to 2 decimal places

                    else:  # Max dose to External is inside PTV
                        ext_max = dose_dist.GetDoseStatistic(RoiName=ext.Name, DoseType='Max')
                        ptv_max = dose_dist.GetDoseStatistic(RoiName=ptv_name, DoseType='Max')
                        
                        if ext_max == 0 or ptv_max == 0:  # External or PTV geometry is empty, or geometry has been updated since last voxel volume computation
                            plan_stat_val = 'N/A'
                            bk_color = grey
                        elif ext_max == ptv_max:  # Same max dose, so assume it's the same point
                            plan_stat_val = 'Yes'
                            bk_color = green
                        else:  # Different max doses, so assume they're at different points
                            plan_stat_val = 'No'
                            bk_color = red

                    # If this stat is to be interpolated, add the None and Minor interpolated values to the row in interp cutoffs table
                    if stat in SbrtLungConstants.INTERP_STATS:
                        none_dev = float(interpolate.interp1d(data['PTV vol [cc]'], data['{} None'.format(stat)])(ptv_vol))
                        minor_dev = float(interpolate.interp1d(data['PTV vol [cc]'], data['{} Minor'.format(stat)])(ptv_vol))
                        bk_color = green if plan_stat_val < none_dev else yellow if plan_stat_val < minor_dev else red
                        # Display no decimal places for V20Gy. Else, display 2 decimal places.
                        if stat == 'V20Gy [%]':
                            none_dev, minor_dev = int(none_dev), int(minor_dev)
                        else:
                            none_dev, minor_dev = round(none_dev, 2), round(minor_dev, 2)
                        interp_cutoffs_row.extend([Paragraph('<{}'.format(none_dev), style=SbrtLungConstants.STYLES['Normal']), Paragraph('<{}'.format(minor_dev), style=SbrtLungConstants.STYLES['Normal'])])

                    # Add stat to plan stats row
                    plan_stats_row.append(plan_stat_val)
                    plan_stats_style.append(('BACKGROUND', (j + 2, plan_stats_idx), (j + 2, plan_stats_idx), bk_color))
                    plan_stats_style.append(('TEXTCOLOR', (j + 2, plan_stats_idx), (j + 2, plan_stats_idx), get_text_color(bk_color)))

                # Add row to plan stats table
                plan_stats_row = [Paragraph(str(text), style=SbrtLungConstants.STYLES['Normal']) for text in plan_stats_row]
                plan_stats_data.append(plan_stats_row)

                # Add row to interp cutoffs table
                interp_cutoffs_data.append(interp_cutoffs_row)
                interp_cutoffs_style.append(('BACKGROUND', (0, i + plan_stats_idx), (-1, i + plan_stats_idx), color))
                interp_cutoffs_style.append(('TEXTCOLOR', (0, i + plan_stats_idx), (-1, i + plan_stats_idx), text_color))

            ptv_vols.pop(0)

        # Add non-plan row to interp cutoffs table
        interp_cutoffs_row = [round(row['PTV vol [cc]'], 1)]  # e.g., [1.8]
        for stat in SbrtLungConstants.INTERP_STATS:
            for dev in ['None', 'Minor']:
                interp_cutoffs_val = row['{} {}'.format(stat, dev)]  # e.g., 'V20Gy [%] None'
                # If V20Gy, display without decimal place. Otherwise, display a single decimal place
                if stat == 'V20Gy [%]':
                    interp_cutoffs_row.append(int(interp_cutoffs_val))
                else:
                    interp_cutoffs_row.append(round(interp_cutoffs_val, 1))
        interp_cutoffs_row = [Paragraph(str(text), style=SbrtLungConstants.STYLES['Normal']) for text in interp_cutoffs_row]
        interp_cutoffs_data.append(interp_cutoffs_row)

    # Finally create tables
    plan_stats_tbl = Table(plan_stats_data, style=TableStyle(plan_stats_style))
    interp_cutoffs_tbl = Table(interp_cutoffs_data, style=TableStyle(interp_cutoffs_style))

    #elems = [KeepTogether([hdr, spcr_sm, mrn, spcr_sm, desc]), spcr_lg, KeepTogether([key_1, spcr_sm, key_2]), spcr_lg, KeepTogether([plan_stats_title, spcr_sm, plan_stats_tbl])]  # List of elements to build PDF from
    elems = [KeepTogether([hdr, mrn, desc]), KeepTogether([key_1, Spacer(letter[1], 0.1 * inch), key_2]), KeepTogether([plan_stats_title, plan_stats_tbl])]  # List of elements to build PDF from
    
    # Add interp cutoffs table to PDF only if interp stat(s) were selected
    # Would be more efficient to just not build this table, but...
    if SbrtLungConstants.INTERP_STATS:
        #elems.extend([spcr_lg, KeepTogether([interp_cutoffs_title, spcr_sm, interp_cutoffs_tbl])])
        elems.extend([KeepTogether([interp_cutoffs_title, interp_cutoffs_tbl])])

    # Build PDF report
    pdf.build(elems)

    return filepath
