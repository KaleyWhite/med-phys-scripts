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


class ScaleObjectivesConstraintsForm(Form):
    """Class that allows the user to choose a reference dose and dose to scale to, for scaling objectives and constraints

    When both doses are valid, the scale factor is provided along with a preview of how an objective/constraint will be scaled
    """
    def __init__(self, example_obj: PyScriptObject, rx: Optional[int] = None) -> None:
        """Initializes a ScaleObjectivesConstraintsForm object

        Arguments
        ---------
        example_obj: Objective or contraint to be used in the preview
        rx: The beam set's primary prescription
            Used as the default reference dose
        """
        self._set_up_form()

        self._y = 15  # Vertical coordinate of next control

        # OK button
        self._ok = Button()
        self._set_up_ok_btn()

        # Script description/instructions
        l = self._add_lbl('Scale objectives and constraints to a new dose value.')
        self._y += l.Height + 10

        # Align textboxes by using uniform label width
        self.txt_box_x = max(TextRenderer.MeasureText(txt, Font(l.Font, FontStyle.Bold)).Width for txt in ['Reference dose:', 'Dose to scale to:', 'Scale factor:', 'Preview:']) + 5

        # Reference dose
        # Default value is current beam set Rx, if it is present
        self._ref_dose_tb = self._add_dose_input('Reference dose:', rx)
        self._y += self._ref_dose_tb.Height

        # Dose to scale to
        # No default value
        self._scale_dose_tb = self._add_dose_input('Dose to scale to:')
        self._y += self._scale_dose_tb.Height

        # Scale factor
        # "?" until both dose textboxes have valid values
        self._add_lbl('Scale factor:', bold=True)
        self._scale_factor_lbl = self._add_lbl('?', x=self.txt_box_x)
        self._y += self._scale_factor_lbl.Height + 20

        ## Preview
        # Shows an example of scaling an objective/constraint
        l = self._add_lbl('Preview:', bold=True)
        self._y += l.Height
        dfp = example_obj.DoseFunctionParameters
        if not hasattr(dfp, 'FunctionType'):  # Dose falloff
            string = f'Dose Fall-Off [H]__ cGy [L]__ cGy, Low dose distance {dfp.LowDoseDistance:.2f} cm'
            self._doses = [int(dfp.HighDoseLevel), int(dfp.LowDoseLevel)]  # Current dose values in the example objective/constraint
        else:
            self._doses = [int(dfp.DoseLevel)]  # Current dose values in the example objective/constraint
            if dfp.FunctionType == 'MinDose':
                string = 'Min dose __ cGy'
            elif dfp.FunctionType == 'MaxDose':
                string = 'Max dose __ cGy'
            elif dfp.FunctionType == 'MinDvh':
                string = f'Min DVH __ cGy to {dfp.PercentVolume}% volume'
            elif dfp.FunctionType == 'MaxDvh':
                string = f'Max DVH __ cGy to {dfp.PercentVolume}% volume'
            elif dfp.FunctionType == 'UniformDose':
                string = 'Uniform dose __ cGy'
            elif dfp.FunctionType == 'MinEud':
                string = f'Min EUD __ cGy, Parameter A {dfp.EudParameterA:.0f}'
            elif dfp.FunctionType == 'MaxEud':
                string = f'Max EUD __ cGy, Parameter A {dfp.EudParameterA:.0f}'
            else:  # UniformEud
                string = f'Target EUD __ cGy, Parameter A {dfp.EudParameterA:.0f}'
        self._add_preview_objective(string, self._doses)  # Labels containing current values for the example objective/constraint
        l = self._add_lbl('will be changed to', x=70)
        self._y += l.Height
        self._dose_lbls = self._add_preview_objective(string, ['?'] * len(self._doses))  # Values for scaled example objective/constraint are '?' until a scale factor is computed
        
        self._set_up_ok_btn()

    # ------------------------------ Styling methods ----------------------------- #

    def _set_up_form(self) -> None:
        """Styles the Form"""
        # Adapt form size to controls
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  

        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = 'Scale Objectives and Constraints'
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

    def _add_preview_objective(self, string: str, doses: List[int]) -> List[Label]:
        """Adds an example objective/constraint with the given dose values
        
        Dose values are bold

        Arguments
        ---------
        string: Text to display, with "__" standing in for each dose value
        doses: List of dose values to substitute for "__" in the string

        Returns
        -------
        A list of Labels displaying each dose text in `doses`

        Example
        -------
        self._add_preview_objective('Here's a dose: __. And another: __.', [5000, 6000]) -> [Label with bold text "5000", Label with bold text "6000"]
        Displays "Here's a dose: 5000. And another: 6000." in the GUI and returns
        """
        lbl_txts = string.split('__')  # E.g., 'Dose Fall-Off [H]__ cGy [L]__ cGy, Low dose distance 1.00 cm' -> ['Dose Fall-Off [H]', ' cGy [L]', ' cGy, Low dose distance 1.00 cm']
        dose_lbls = []
        x = 50  # Offset from left
        for i, txt in enumerate(lbl_txts):
            l = self._add_lbl(txt, x=x)
            self.Controls.Add(l)
            x += l.Width
            # There is dose to display
            if i != len(lbl_txts) - 1:
                l_2 = self._add_lbl(str(doses[i]), bold=True, x=x)
                l_2.Width = 45  # Wide enough for 6 digits
                self.Controls.Add(l_2)
                x += 45
                dose_lbls.append(l_2)
        self._y += l.Height
        return dose_lbls

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
            for i, dose_lbl in enumerate(self._dose_lbls):
                dose_lbl.Text = f'{(self.scale_factor * self._doses[i]):.0f}'
            self._ok.Enabled = True
        # Invalid dose value(s)
        else:
            self._scale_factor_lbl.Text = '?'  # Can't compute scale factor
            for dose_lbl in self._dose_lbls:  # Can't compute preview dose value(s)
                dose_lbl.Text = '?'
            self._ok.Enabled = False

    def _ok_Click(self, sender: Button, event: EventArgs) -> None:
        """Event handler for clicking the OK button"""
        self.DialogResult = DialogResult.OK


def scale_objectives_constraints() -> None:
    """Scales the current beam set's objectives and constraints to a new dose value

    (If there is only one beam set in the current plan, these objectives and constraints apply to the plan as a whole.)
    Changes objectives and constraints in place: does not add objectives and constraints.
    User supplies reference dose (default is Rx dose, if it is provided) and dose to scale to, in a GUI.
    All doses in objectives and constraints are multiplied by a scale factor = dose to scale to / reference dose

    If current plan is approved, user may scale objectives/constraints in a copy
    """
    # Get current variables
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
        sys.exit()
    try:
        beam_set = get_current('BeamSet')
    except:
        MessageBox.Show('There are no beam sets in the open plan. Click OK to abort the script.', 'No Beam Sets')
        sys.exit()

    # If plan is approved, offer to change objectives/constraints on a copy
    if plan.Review is not None and plan.Review.ApprovalStatus == 'Approved':
        res = MessageBox.Show('Plan is approved, so objectives and constraints cannot be changed. Would you like to scale objectives and constraints on a copy of the plan?', 'Plan Is Approved', MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit()
        plan = copy_plan_without_changes()

    # Get objectives and constraints for current beam set
    try:
        plan_opt = next(opt for opt in plan.PlanOptimizations if opt.OptimizedBeamSets[0].DicomPlanLabel == beam_set.DicomPlanLabel)
    except StopIteration:
        MessageBox.Show('There is no optimization for this beam set. Click OK to abort the script.', 'No Optimization')
        sys.exit()
    objectives = plan_opt.Objective
    constraints = plan_opt.Constraints
    scalable_constraints = [constraint for constraint in constraints if not hasattr(constraint.DoseFunctionParameters, 'PercentStdDeviation')]  # Uniform dose constraints do not have "PercentStdDeviation" attribute and are not scalable
    
    # Exit script if no objectives/constraints
    if objectives is None:  # No objectives
        if constraints.Count == 0:  # NO constraints, period
            MessageBox.Show('There are no objectives or constraints for the current beam set. Click OK to abort the script.', 'No Objectives/Constraints')
            sys.exit()
        if not scalable_constraints:  # No constraints that scaling the dose would affect
            MessageBox.Show('The current beam set has no objectives. All constraints are uniform dose constraints, which are not parameterized by dose. Click OK to abort the script.', 'No Objectives/Constraints')
            sys.exit()

    # Get default reference dose
    rx = beam_set.Prescription.PrimaryPrescriptionDoseReference
    if rx is not None:
        rx = int(rx.DoseValue)  # Rx can't be fractional, so this is not truncating

    # Example objective/constraint for preview in GUI
    example_obj = objectives.ConstituentFunctions[0] if objectives is not None else scalable_constraints[0]

    # Get scale factor from user input in a GUI
    form = ScaleObjectivesConstraintsForm(example_obj, rx)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
    scale_factor = form.scale_factor
    
    # List of objectives and/or constraints
    objectives_constraints = list(scalable_constraints)
    if objectives is not None:
        objectives_constraints.extend(list(objectives.ConstituentFunctions))

    # Scale each objective or constraint
    for o_c in objectives_constraints:
        dfp = o_c.DoseFunctionParameters
        for attr in dir(dfp):
            # Multiply each dose parameter by the scale factor
            if 'DoseLevel' in attr:
                val = getattr(dfp, attr)
                setattr(dfp, attr, val * scale_factor)
