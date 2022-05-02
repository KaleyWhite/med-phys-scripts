import clr
import sys
from typing import List, Optional

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

from System import EventArgs  # For type hints
from System.Drawing import *
from System.Windows.Forms import *

sys.path.append(r'T:\Physics\KW\med-phys-scripts\RayStation')
from copy_plan_without_changes import copy_plan_without_changes


class ScaleClinicalGoalsForm(Form):
    """Class that allows the user to choose a reference dose and dose to scale to, for scaling clinical goals

    When both doses are valid, the scale factor is provided along with a preview of how a goal will be scaled
    """
    def __init__(self, example_txt: str, example_dose: int, rx: Optional[int] = None) -> None:
        """Initializes a ScaleClinicalGoalsForm object

        Arguments
        ---------
        example_txt: Text of the example goal for the preview, with a double underscore to stand in for the dose
        example_dose: The dose parameter of the example goal for the preview
        rx: The beam set's primary prescription
            Used as the default dose to scale to
        """
        self._dose = example_dose

        self._set_up_form()

        self._y = 15  # Vertical coordinate of next control

        # OK button
        self._ok = Button()
        self._set_up_ok_btn()

        # Script description/instructions
        l = self._add_lbl('Scale clinical goals to a new dose value.')
        self._y += l.Height + 10

        # Align textboxes by using uniform label width
        self.txt_box_x = max(TextRenderer.MeasureText(txt, Font(l.Font, FontStyle.Bold)).Width for txt in ['Reference dose:', 'Dose to scale to:', 'Scale factor:', 'Preview:']) + 5

        # Reference dose
        # Default value is current beam set Rx, if it is present
        self._ref_dose_tb = self._add_dose_input('Reference dose:')
        self._y += self._ref_dose_tb.Height

        # Dose to scale to
        # No default value
        self._scale_dose_tb = self._add_dose_input('Dose to scale to:', rx)
        self._y += self._scale_dose_tb.Height

        # Scale factor
        # "?"" until both dose textboxes have valid values
        self._add_lbl('Scale factor:', bold=True)
        self._scale_factor_lbl = self._add_lbl('?', x=self.txt_box_x)
        self._y += self._scale_factor_lbl.Height + 20

        # Preview shows an example of scaling a goal
        l = self._add_lbl('Preview:', bold=True)
        self._y += l.Height
        dose_str = str(round(self._dose, 2)).rstrip('0').rstrip('.')
        self._add_preview_objective(example_txt, dose_str)
        l = self._add_lbl('will be changed to', x=70)
        self._y += l.Height
        self._dose_lbl = self._add_preview_objective(example_txt, '?')  # Value for scaled example goal is "?" until a scale factor is computed
        
        self._set_up_ok_btn()

    # ------------------------------ Styling methods ----------------------------- #

    def _set_up_form(self) -> None:
        """Styles the Form"""
        # Adapt form size to controls
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  

        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = 'Scale Clinical Goals'
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)  # At least as wide as title plus room for "X" button, etc.

    def _set_up_ok_btn(self) -> None:
        """Styles and adds the "OK" button to the Form"""
        self.Controls.Add(self._ok)
        self.AcceptButton = self._ok  # Clicked when "Enter" is pressed
        self._ok.Enabled = False  # By default, no dose to scale to is provided
        self._ok.Text = 'OK'
        self._ok.Location = Point(self.ClientSize.Width - 50, self._y)  # Right align
        self._ok.Click += self._ok_Click
        
    def _add_lbl(self, lbl_txt: str, **kwargs) -> Label:
        """Adds a Label to the Form

        Arguments
        ---------
        lbl_txt: Text of the Label

        Keyword Arguments
        -----------------
        x: x-coordinate of new label
        bold: True if label text should be bold, False otherwise

        Returns
        -------
        The Label
        """
        x = kwargs.get('x', 15)
        bold = kwargs.get('bold', False)

        l = Label()
        l.AutoSize = True
        if bold:
            l.Font = Font(l.Font, FontStyle.Bold)
        l.Location = Point(x, self._y)
        l.Text = lbl_txt
        self.Controls.Add(l)
        return l

    def _add_dose_input(self, lbl_txt: str, dose: Optional[int] = None) -> TextBox:
        """Adds a set of controls for getting dose input from the user:
            - Label with description of dose field
            - TextBox
            - "cGy" Label

        Arguments
        ---------
        lbl_txt: Description of field
        dose: Starting value for text box. If None, text box starts out empty.

        Returns
        -------
        The TextBox that holds the dose input
        """
        # Label w/ description of field
        l = self._add_lbl(lbl_txt, bold=True)

        # TextBox
        tb = TextBox()
        tb.Location = Point(self.txt_box_x, self._y - 5)  # Approximately vertically centered with description label
        if dose is not None:
            tb.Text = str(dose)
        tb.TextChanged += self._dose_TextChanged
        tb.Size = Size(40, l.Height)  # Wide enough for 6 digits
        self.Controls.Add(tb)

        # 'cGy' label
        self._add_lbl('cGy', x=self.txt_box_x + 45)
        self._y += l.Height

        return tb

    def _add_preview_objective(self, string: str, dose: str) -> Label:
        """Displays an example goal with the given dose values
        
        Dose values are bold

        Arguments
        ---------
        string: Text to display, with "__" standing in for the dose value
        dose: Dose values to substitute for "__" in the string

        Returns
        -------
        A Labels displaying the dose

        Example
        -------
        self._add_preview_objective('Here's a dose: __.', 5000) -> Label with bold text "5000"
        Displays "Here's a dose: 5000." in the GUI
        """
        lbl_txts = string.split('__')  # E.g., 'At most 0.25 cm^3 volume at __ cGy dose' -> ['At most 0.25 cm^3 volume at ', ' cGy dose']
        
        x = 50  # Offset from left
        l = self._add_lbl(lbl_txts[0], x=x)
        self.Controls.Add(l)
        x += l.Width

        dose_l = self._add_lbl(dose, bold=True, x=x)
        dose_l.Width = 45  # Wide enough for 6 digits
        self.Controls.Add(dose_l)
        x += 45

        l = self._add_lbl(lbl_txts[1], x=x)
        self.Controls.Add(l)

        self._y += l.Height
        return dose_l

    # ------------------------------ Event handlers ------------------------------ #

    def _dose_TextChanged(self, sender: TextBox, event: EventArgs) -> None:
        """Event handler for changing the text of a dose input field
        
        Turns TextBox background red if input is invalid
        If both dose values valid, compute and display scale factor, change preview output, and enable OK button
        If either dose value invalid, display "?" for scale factor and preview output, and disable OK button
        """

        # Reference dose
        self.ref_dose = self._ref_dose_tb.Text
        # If empty, don't turn red, but the value is still not valid
        if self.ref_dose == '':
            self._ref_dose_tb.BackColor = Color.White
            ref_dose_ok = False
        else:
            try:
                self.ref_dose = float(self.ref_dose)
                if self.ref_dose != int(self.ref_dose) or self.ref_dose <= 0 or self.ref_dose > 100000:  # Number is fractional or out of range
                    raise ValueError
                self._ref_dose_tb.BackColor = Color.White
                ref_dose_ok = True
            except ValueError:
                self._ref_dose_tb.BackColor = Color.Red
                ref_dose_ok = False

        # Dose to scale to
        self.scale_dose = self._scale_dose_tb.Text
        # If empty, don't turn red, but the value is still not valid
        if self.scale_dose == '':
            self._scale_dose_tb.BackColor = Color.White
            scale_dose_ok = False
        else:
            try:
                self.scale_dose = float(self.scale_dose)
                if self.scale_dose != int(self.scale_dose) or self.scale_dose < 0 or self.scale_dose > 100000:  # Number is fractional or out of range
                    raise ValueError
                self._scale_dose_tb.BackColor = Color.White
                scale_dose_ok = True
            except ValueError:
                self._scale_dose_tb.BackColor = Color.Red
                scale_dose_ok = False
       
        # Both dose values are valid
        if ref_dose_ok and scale_dose_ok:
            self.scale_factor = self.scale_dose / self.ref_dose
            
            # Display at most 3 decimal places, with no trailing zeros
            display = str(round(self.scale_factor, 3)).rstrip('0')
            if display.endswith('.'):
                display = display[:-1]
            self._scale_factor_lbl.Text = display
            
            # Display rounded computed doses in preview
            self._dose_lbl.Text = f'{(self.scale_factor * self._dose):.0f}'
            self._ok.Enabled = True
        # Invalid dose value(s)
        else:
            self._scale_factor_lbl.Text = self._dose_lbl.Text = '?'  # Can't compute scale factor or preview dose value(s)
            self._ok.Enabled = False

    def _ok_Click(self, sender: Button, event: EventArgs) -> None:
        """Event handler for clicking the OK button"""
        self.DialogResult = DialogResult.OK


def goal_str(goal: PyScriptObject) -> str:
    """Creates a string representation of the clinical goal

    Attempts to mimic whatever RayStation does behind the scenes
    Dose is replaced with double underscores

    Arguments
    ---------
    goal: The goal to format as a string

    Returns
    -------
    The goal string

    Raises
    ------
    ValueError: If the dose type is not recognized

    Example
    -------
    Given some_goal V2000cGy < 0.25cm^3:
    goal_str(some_goal) -> "At most 0.25 cm^3 volume at 2000 cGy dose"
    """
    goal_type = goal.PlanningGoal.Type

    goal_criteria = 'At least' if goal.PlanningGoal.GoalCriteria == 'AtLeast' else 'At most'
    accept_lvl = goal.PlanningGoal.AcceptanceLevel
    param_val = goal.PlanningGoal.ParameterValue

    goal_txt = goal_criteria + ' '

    if goal_type == 'DoseAtAbsoluteVolume':
        goal_txt += '__ cGy dose at ' + str(param_val) + ' cm^3 volume'
    elif goal_type == 'DoseAtVolume':
        goal_txt += '__ cGy dose at ' + str(param_val * 100) + '% volume'
    elif goal_type == 'AbsoluteVolumeAtDose':
        goal_txt += str(accept_lvl) + ' cm^3 volume at __ cGy dose'
    elif goal_type == 'VolumeAtDose':
        goal_txt += str(accept_lvl * 100) + '% volume at __ cGy dose'
    elif goal_type == 'AverageDose':
        goal_txt += '__ cGy average dose'
    elif goal_type == 'ConformityIndex':
        goal_txt += 'a conformity index of ' + str(accept_lvl) + ' at __ cGy dose'
    elif goal_type == 'DoseAtPoint':  # DoseAtPoint
        goal_txt += '__ cGy dose at point'
    else:
        raise ValueError('Unrecognized goal type "' + goal_type + '"')

    return goal_txt


def goal_attr_to_scale(goal: PyScriptObject) -> str:
    """Determines the cname of the dose attribute of the goal

    Arguments
    ---------
    goal: The goal whose attribute name to return

    Returns
    -------
    The name of the dose attribute. Either "AcceptanceLevel" or "ParameterValue"

    Raises
    ------
    ValueError: If the dose type is not recognized or is "HomogeneityIndex"

    Example
    -------
    Given some_goal V2000cGy < 0.25cm^3:
    goal_attr_to_scale(some_goal) -> "ParameterValue"
    """
    goal_type = goal.PlanningGoal.Type
    if goal_type in ['DoseAtAbsoluteVolume', 'DoseAtVolume', 'AverageDose', 'DoseAtPoint']:
        return 'AcceptanceLevel' 
    if goal_type in ['AbsoluteVolumeAtDose', 'VolumeAtDose', 'ConformityIndex']:
        return 'ParameterValue' 
    if goal_type == 'HomogeneityIndex':
        raise ValueError('Dose of type "HomogeneityIndex" is not parametrized by dose')
    else:
        raise ValueError('Unrecognized goal type "' + goal_type + '"')
    


def scale_clinical_goals() -> None:
    """Scales the current plan's clinical goals to a new dose value

    Changes goals in place: does not add goals.
    User supplies reference dose and dose to scale to (default is primary Rx dose for current beam set, if it is provided), in a GUI.
    Each goal's dose (acceptance level or parameter value, depending on the goal type) are multiplied by a scale factor = dose to scale to / reference dose and rounded to the nearest integer

    If current plan is approved, user may scale goals in a copy
    """
    # Get current variables
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
        sys.exit()

    # If plan is approved, offer to change objectives/constraints on a copy
    if plan.Review is not None and plan.Review.ApprovalStatus == 'Approved':
        res = MessageBox.Show('Plan is approved, so clinical goals cannot be changed. Would you like to scale clinical goals on a copy of the plan?', 'Plan Is Approved', MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit()
        plan = copy_plan_without_changes()

    # Get clinical goals for current plan
    goals = list(plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions)
    if not goals:
        MessageBox.Show('The current plan has no clinical goals. Click OK to abort the script.', 'No Clinical Goals')
        sys.exit()
    scalable_goals = [goal for goal in goals if goal.PlanningGoal.Type != 'HomogeneityIndex']
    if not scalable_goals:
        MessageBox.Show('All of the current plan\'s clinical goals are for the homogeneity index - not parameterized by dose. Click OK to abort the script.', 'No Clinical Goals')
        sys.exit()

    # Get example goal for preview in GUI
    example_goal_txt = goal_str(scalable_goals[0])
    example_goal_dose = getattr(scalable_goals[0].PlanningGoal, goal_attr_to_scale(scalable_goals[0]))

    # Get default dose to scale to
    try:
        beam_set = get_current('BeamSet')
        rx = beam_set.Prescription.PrimaryPrescriptionDoseReference
        if rx is not None:
            rx = int(rx.DoseValue)  # Rx can't be fractional, so this is not truncating
    except:
        rx = None

    # Get scale factor from user input in a GUI
    form = ScaleClinicalGoalsForm(example_goal_txt, example_goal_dose, rx)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
    scale_factor = form.scale_factor

    # Scale each goal
    for goal in scalable_goals:
        attr_name = goal_attr_to_scale(goal)
        attr_val = getattr(goal.PlanningGoal, attr_name)
        setattr(goal.PlanningGoal, attr_name, round(attr_val * scale_factor))
