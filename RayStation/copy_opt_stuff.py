import clr
from collections import OrderedDict
import sys
from typing import Dict, Tuple

from connect import *

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *

sys.path.append(r'T:\Physics\KW\med-phys-scripts\RayStation')
from copy_plan_without_changes import copy_plan_without_changes


class CopyOptStuffForm(Form):
    def __init__(self, curr_plan_name: str, plan_info: Dict[str, Tuple[bool, bool]]):
        """Initializes a CopyOptStuffForm object

        Arguments:
            curr_plan_name: Name of the current open plan
            plan_info: Dictionary that tells whether each plan has clinical goals, and/or objectives and/or constraints
                       Keys are plan names. 
                       Values are 2-tuples, where:
                       The first element is True if the plan has goals.
                       The second element is True if the plan has objectives and/or costraints.
        """
        self.curr_plan_name = curr_plan_name
        self.plan_info = plan_info

        # Plan names to copy goals and objectives from, respectively
        self.copy_goals_from, self.copy_objs_from = [], []
        self.copy_params_from = None

        self.y = 15  # Vertical coordinate of next control

        # Controls
        self.clear_existing_goals_cb = self.create_clear_checkbox('Clear existing clinical goals', 15)
        self.clear_existing_objs_cb = self.create_clear_checkbox('Clear existing objectives && constraints', 200)
        self.data_grid_view = DataGridView()
        self.copy_btn = Button()

        self.set_up_layout()
        # Redraw table
        #self.data_grid_view.Invalidate()
        self.Invalidate()

    def data_grid_view_CellContentClick(self, sender, event):
        # Event handler for clicking a checkbox
        # This allows the cell value to actually change (the CellValueChanged event to be fired)
        self.data_grid_view.CommitEdit(DataGridViewDataErrorContexts.Commit)

    def data_grid_view_CellPainting(self, sender, event):
        # Event handler for when the DGV is redrawn
        # Fills cells with gray instead of a checkbox if the cell corresponds to a False in the plan_info dict
        
        # Determine whether to color the cell gray or leave it alone
        # Only color non-header, read-only cells in a CheckBoxColumn
        if event.RowIndex >= 0 and event.ColumnIndex != self.plan_col.Index:
            if event.ColumnIndex in [self.copy_goals_col.Index, self.copy_objs_col.Index] and self.data_grid_view[event.ColumnIndex, event.RowIndex].ReadOnly:
                # Fill with gray
                brush = SolidBrush(event.CellStyle.BackColor)
                event.Graphics.FillRectangle(brush, event.CellBounds)

                # Replace the black outline that the filled rectangle covers
                event.Paint(event.CellBounds, DataGridViewPaintParts.Border)
            elif event.ColumnIndex == self.copy_params_col.Index:
                # Draw radio button
                #event.PaintBackground(event.ClipBounds, True)
                rb = Rectangle()
                rb.Width = rb.Height = 14
                rb.X = event.CellBounds.X + (event.CellBounds.Width - rb.Width) / 2
                rb.Y = event.CellBounds.Y + (event.CellBounds.Height - rb.Height) / 2

                rb_state = ButtonState.Checked if event.Value else ButtonState.Normal
                ControlPaint.DrawRadioButton(event.Graphics, rb, rb_state)
                event.Paint(event.CellBounds, DataGridViewPaintParts.Focus)
            
            # Skip the rest of the event handler
            event.Handled = True

    def data_grid_view_CellValueChanged(self, sender, event):
        # Event handler for cell value change in the DGV
        # Updates lists of plans to copy goals/objs from, and enables or disables the "OK" button accordingly

        # Only look at the CheckBoxColumns
        if event.ColumnIndex > 0:
            plan_name = self.data_grid_view[0, event.RowIndex]
            val = self.data_grid_view[event.ColumnIndex, event.RowIndex].Value
            # Select copy_goals_from, copy_objs_from, or copy_params_from list depending on which column the clicked checkbox belongs to
            if event.ColumnIndex in [self.copy_goals_col.Index, self.copy_objs_col.Index]:
                l = self.copy_goals_from if event.RowIndex == self.copy_goals_col.Index else self.copy_objs_from
                if val:  # The checkbox was checked
                    # Add plan name to list if it is not already there
                    if plan_name not in l:
                        l.append(plan_name)
                else:  # The checkbox was unchecked
                    # Remove plan from list if it is there
                    if plan_name in l:
                        l.remove(plan_name)
            elif val:
                self.copy_params_from = plan_name
            else:
                self.copy_params_from = None
            
            # Enable "OK" button only if at least one set of goals or objectives is checked
            self.copy_btn.Enabled = self.copy_goals_from or self.copy_objs_from or self.copy_params_from is not None

    def data_grid_view_CurrentCellDirtyStateChanged(self, sender, event):
        if self.data_grid_view.CurrentCell.ColumnIndex == self.copy_params_col.Index:
            for row in self.data_grid_view.Rows:
                if row.Index != self.data_grid_view.CurrentCell.RowIndex:
                    row.Cells[self.copy_params_col.Index].Value = False

    def copy_Click(self, sender, event):
        # Event handler for "OK" button
        self.DialogResult = DialogResult.OK

    def set_up_layout(self):
        # Styles controls and adds them to the Form
    
        self.Text = 'Copy Optimization Stuff'  # Form title

        # Adapt form size to controls
        # Make form at least as wide as title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink 
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

        # "Clear" checkboxes
        # Disabled if current plan has no goals/objectives 
        self.Controls.Add(self.clear_existing_goals_cb)
        self.Controls.Add(self.clear_existing_objs_cb)
        self.clear_existing_goals_cb.Enabled, self.clear_existing_objs_cb.Enabled = self.plan_info[self.curr_plan_name]
        self.y += self.clear_existing_goals_cb.Height + 15

        # Table
        # Each row is a plan and checkboxes for copying goals/objectives if they exist
        self.set_up_data_grid_view()

        # "OK" button
        self.Controls.Add(self.copy_btn)
        self.copy_btn.AutoSize = True
        self.copy_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.copy_btn.Enabled = False
        self.copy_btn.Location = Point(self.Width // 2, self.y)
        self.copy_btn.Text = 'OK'
        self.copy_btn.Click += self.copy_Click

    def set_up_data_grid_view(self):
        # Style and add a DGV to the Form

        # Basic settings
        self.Controls.Add(self.data_grid_view)
        self.data_grid_view.Location = Point(15, self.y)
        self.data_grid_view.ScrollBars = 0
        self.data_grid_view.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.data_grid_view.GridColor = Color.Black

        # Row settings
        self.data_grid_view.RowHeadersVisible = False
        self.data_grid_view.AllowUserToAddRows = False
        self.data_grid_view.AllowUserToDeleteRows = False
        self.data_grid_view.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells

        # Column settings
        self.data_grid_view.AllowUserToOrderColumns = False
        self.data_grid_view.AllowUserToResizeColumns = False
        self.data_grid_view.AutoGenerateColumns = False
        self.data_grid_view.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        self.data_grid_view.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        self.data_grid_view.ColumnHeadersDefaultCellStyle.Font = Font(self.data_grid_view.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
        self.data_grid_view.ColumnHeadersDefaultCellStyle.SelectionBackColor = self.data_grid_view.ColumnHeadersDefaultCellStyle.BackColor
        self.data_grid_view.ColumnHeadersDefaultCellStyle.SelectionForeColor = self.data_grid_view.ColumnHeadersDefaultCellStyle.ForeColor
        self.data_grid_view.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self.data_grid_view.EnableHeadersVisualStyles = False
    
        # Cell settings
        self.data_grid_view.DefaultCellStyle.SelectionBackColor = self.data_grid_view.DefaultCellStyle.BackColor
        self.data_grid_view.DefaultCellStyle.SelectionForeColor = self.data_grid_view.DefaultCellStyle.ForeColor
        
        # Event handlers
        self.data_grid_view.CellPainting += self.data_grid_view_CellPainting
        self.data_grid_view.CellContentClick += self.data_grid_view_CellContentClick
        self.data_grid_view.CellValueChanged += self.data_grid_view_CellValueChanged
        self.data_grid_view.CurrentCellDirtyStateChanged += self.data_grid_view_CurrentCellDirtyStateChanged

        self.add_data_grid_cols()
        self.populate_data_grid_view()

    def add_data_grid_cols(self):
        # Adds the three columns to the table: "Plan", "Copy Clinical Goals", and "Copy Objectives & Constraints"
        # "Plan" column
        self.plan_col = DataGridViewTextBoxColumn()
        self.plan_col.HeaderText = 'Plan'
        self.plan_col.ReadOnly = True  # User cannot change values
        self.data_grid_view.Columns.Add(self.plan_col)

        # "Copy" columns
        self.copy_goals_col = self.create_copy_col('Copy Clinical Goals')
        self.data_grid_view.Columns.Add(self.copy_goals_col)

        self.copy_objs_col = self.create_copy_col('Copy Objectives & Constraints')
        self.data_grid_view.Columns.Add(self.copy_objs_col)

        self.copy_params_col = self.create_copy_col('Copy Optimization Parameters')
        self.data_grid_view.Columns.Add(self.copy_params_col)

    def populate_data_grid_view(self):
        # Add rows to the DGV, from the plan_info dictionary

        # Populate rows
        for i, plan_name in enumerate(self.plan_info):
            print(plan_name)
            # Skip current plan b/c we don't want to copy to itself
            if plan_name == self.curr_plan_name:
                continue
            
            # Add row
            # Both checkboxes unchecked by default
            row = [plan_name, False, False, False]
            self.data_grid_view.Rows.Add(row)

            # Disable "Copy" checkbox if plan does not have goals/objectives
            has_goals, has_objs = self.plan_info[plan_name]
            row = self.data_grid_view.Rows[i]  # The row we just added
            if not has_goals:
                self.disable_checkbox(row.Cells['Copy Clinical Goals'])
            if not has_objs:
                self.disable_checkbox(row.Cells['Copy Objectives & Constraints'])

        # Resize table to contents
        self.data_grid_view.Width = sum(col.Width + 37 for col in self.data_grid_view.Columns)
        self.data_grid_view.Height = self.data_grid_view.Rows[0].Height * (self.data_grid_view.Rows.Count + 1)
        self.y += self.data_grid_view.Height + 15

    def create_clear_checkbox(self, txt, x):
        # Helper function that returns a "Clear ___" checkbox
        # txt: Text attribute
        # x: x-coordinate of checkbox

        cb = CheckBox()
        cb.AutoSize = True
        cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        cb.Location = Point(x, self.y)
        cb.Text = txt

        return cb

    def create_copy_col(self, txt):
        # Helper function that returns a "Copy ___" DGV column
        # txt: column name

        col = DataGridViewCheckBoxColumn()
        col.Name = col.HeaderText = txt

        return col

    def disable_checkbox(self, cell):
        # Helper function that "disables" a checkbox if the plan does not have goals/objectives to copy

        cell.Style.BackColor = Color.Gray
        cell.ReadOnly = True


def copy_clinical_goals(old_plan, new_plan):
    old_goals = old_plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions
    new_goals = new_plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions
    for goal in old_goals:
        goal_args = {
            'RoiName': goal.ForRegionOfInterest.Name,
            'GoalCriteria': goal.PlanningGoal.GoalCriteria,
            'GoalType': goal.PlanningGoal.Type,
            'AcceptanceLevel': goal.PlanningGoal.AcceptanceLevel,
            'ParameterValue': goal.PlanningGoal.ParameterValue,
            'Priority': goal.PlanningGoal.Priority,
        }
        new_goals.AddClinicalGoal(**goal_args)


# All code for copy objectives/constraints is taken from RaySearch's Scripting Guideline script Example_19_Copy_optimization_functions

# Define a function that will retrieve the necessary information
# and put it in a dictionary.
def get_arguments_from_function(function):
    dfp = function.DoseFunctionParameters
    arg_dict = {}
    arg_dict['RoiName'] = function.ForRegionOfInterest.Name
    arg_dict['IsRobust'] = function.UseRobustness
    arg_dict['Weight'] = dfp.Weight
    if hasattr(dfp, 'FunctionType'):
        if dfp.FunctionType == 'UniformEud':
            arg_dict['FunctionType'] = 'TargetEud'
        else:
            arg_dict['FunctionType'] = dfp.FunctionType
        arg_dict['DoseLevel'] = dfp.DoseLevel
        if 'Eud' in dfp.FunctionType:
            arg_dict['EudParameterA'] = dfp.EudParameterA
        elif 'Dvh' in dfp.FunctionType:
            arg_dict['PercentVolume'] = dfp.PercentVolume
    elif hasattr(dfp, 'HighDoseLevel'):
        # Dose fall-off function does not have function type attribute.
        arg_dict['FunctionType'] = 'DoseFallOff'
        arg_dict['HighDoseLevel'] = dfp.HighDoseLevel
        arg_dict['LowDoseLevel'] = dfp.LowDoseLevel
        arg_dict['LowDoseDistance'] = dfp.LowDoseDistance
    elif hasattr(dfp, 'PercentStdDeviation'):
        # Uniformity constraint does not have function type.
        arg_dict['FunctionType'] = 'UniformityConstraint'
        arg_dict['PercentStdDeviation'] = dfp.PercentStdDeviation
    else:
        # Unknown function type, raise exception.
        raise ('Unknown function type')

    return arg_dict


# Define a function that will use information in arg_dict to set
# optimization function parameters.
def set_function_arguments(function, arg_dict):
    dfp = function.DoseFunctionParameters
    dfp.Weight = arg_dict['Weight']
    if arg_dict['FunctionType'] == 'DoseFallOff':
        dfp.HighDoseLevel = arg_dict['HighDoseLevel']
        dfp.LowDoseLevel = arg_dict['LowDoseLevel']
        dfp.LowDoseDistance = arg_dict['LowDoseDistance']
    elif arg_dict['FunctionType'] == 'UniformityConstraint':
        dfp.PercentStdDeviation = arg_dict['PercentStdDeviation']
    else:
        dfp.DoseLevel = arg_dict['DoseLevel']
        if 'Eud' in dfp.FunctionType:
            dfp.EudParameterA = arg_dict['EudParameterA']
        elif 'Dvh' in dfp.FunctionType:
            dfp.PercentVolume = arg_dict['PercentVolume']


def copy_objectives_and_constraints(old_plan, new_plan):
    # Loop over all objectives and constraints in the original plan
    # to retrieve information. It is assumed that the original plan
    # only contains one beam set.
    po_original = original_plan.PlanOptimizations[0]
    arguments = [] # List to hold arg_dicts of all functions.

    # Get arguments from objective functions.
    if po_original.Objective != None:
        for cf in po_original.Objective.ConstituentFunctions:
            arg_dict = get_arguments_from_function(cf)
            arg_dict['IsConstraint'] = False
            arguments.append(arg_dict)

    # Get arguments from constraint functions.
    for cf in po_original.Constraints:
        arg_dict = get_arguments_from_function(cf)
        arg_dict['IsConstraint'] = True
        arguments.append(arg_dict)

    # Add optimization functions to the new plan.
    po_new = new_plan.PlanOptimizations[0]
    for arg_dict in arguments:
        with CompositeAction('Add Optimization Function'):
            f = po_new.AddOptimizationFunction(FunctionType = arg_dict['FunctionType'],
                                            RoiName = arg_dict['RoiName'],
                                            IsConstraint = arg_dict['IsConstraint'],
                                            IsRobust = arg_dict['IsRobust'])
            set_function_arguments(f, arg_dict)


def copy_opt_params(old_plan: PyScriptObject, new_plan: PyScriptObject) -> None:
    """Copies select optimization parameters from one beam set to another.

    Assumes only one beam set per plan optimization.

    Copies the following optimization parameters:
    - Autoscale to prescription
    - Maximum number of iterations and number of iterations in preparation phase
    - Optimization tolerance
    - Whether to compute final and intermediate dose
    For VMAT plans, also copies treatment setup settings:
    - Maximum leaf travel distance per degree, and whether it is used
    - Arc conversion properties per beam
    - Whether to use dual arcs
    - Final arc gantry spacing
    - Maximum arc delivery time

    This function is modified from the RayStation 11A Scripting Guideline.
    
    Args:
        patient: The patient that the beam sets belong to.
        old_beam_set: The beam set to copy optimization parameters from.
        new_beam_set: The beam set to copy optimization parameters to.
    """
    old_plan_opt = old_plan.PlanOptimizations[0]
    new_plan_opt = new_plan.PlanOptimizations[0]

    with CompositeAction('Copy Select Optimization Parameters'):
        new_plan_opt.AutoScaleToPrescription = old_plan_opt.AutoScaleToPrescription

        new_plan_opt.OptimizationParameters.Algorithm.MaxNumberOfIterations = old_plan_opt.OptimizationParameters.Algorithm.MaxNumberOfIterations
        new_plan_opt.OptimizationParameters.Algorithm.OptimalityTolerance = old_plan_opt.OptimizationParameters.Algorithm.OptimalityTolerance
        
        new_plan_opt.OptimizationParameters.DoseCalculation.ComputeFinalDose = old_plan_opt.OptimizationParameters.DoseCalculation.ComputeFinalDose
        new_plan_opt.OptimizationParameters.DoseCalculation.ComputeIntermediateDose = old_plan_opt.OptimizationParameters.DoseCalculation.ComputeIntermediateDose
        new_plan_opt.OptimizationParameters.DoseCalculation.IterationsInPreparationsPhase = old_plan_opt.OptimizationParameters.DoseCalculation.IterationsInPreparationsPhase

        old_beam_set = old_plan_opt.OptimizedBeamSets[0]
        if old_beam_set.Modality == 'Photons' and old_beam_set.PlanGenerationTechnique == 'Imrt' and old_beam_set.DeliveryTechnique == 'DynamicArc':
            for i, old_tx_setup_settings in enumerate(old_plan_opt.OptimizationParameters.TreatmentSetupSettings):
                new_tx_setup_settings = new_plan_opt.OptimizationParameters.TreatmentSetupSettings[i]
                
                new_tx_setup_settings.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree = old_tx_setup_settings.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree
                new_tx_setup_settings.SegmentConversion.ArcConversionProperties.UseMaxLeafTravelDistancePerDegree = old_tx_setup_settings.SegmentConversion.ArcConversionProperties.UseMaxLeafTravelDistancePerDegree

                for j, old_beam_set in enumerate(old_tx_setup_settings.BeamSettings):
                    old_beam_set_props = old_beam_set.ArcConversionPropertiesPerBeam
                    new_beam_set_props = new_tx_setup_settings.BeamSettings[j].ArcConversionPropertiesPerBeam
                    if old_beam_set_props.NumberOfArcs != new_beam_set_props.NumberOfArcs or old_beam_set_props.FinalArcGantrySpacing != new_beam_set_props.FinalArcGantrySpacing or old_beam_set_props.MaxArcDeliveryTime != new_beam_set_props.MaxArcDeliveryTime:
                        new_beam_set_props.EditArcBasedBeamOptimizationSettings(CreateDualArcs=old_beam_set_props.NumberOfArcs == 2, FinalGantrySpacing=old_beam_set_props.FinalArcGantrySpacing, MaxArcDeliveryTime=old_beam_set_props.MaxArcDeliveryTime) 


def copy_opt_stuff():
    """Copies clinical goals and/or objectives and constraints from other plan(s) to the current plan

    Goals and objectives are added in alphabetical order of selected plans (case insensitive)

    The user chooses the plans to copy goals/objectives from, from a GUI
    The user also chooses whether to clear the current plan's existing goals and objectives/constraints, if applicable

    Assumes that each plan has only one optimization
    """

    # Ensure that patient, case, and plan are open
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There are no plans in the current case. Click OK to abort the script.', 'No Plans')
        sys.exit()

    # Plan names that contain goals and objectives, respectively
    plan_info = OrderedDict()
    for p in case.TreatmentPlans:
        plan_opt = p.PlanOptimizations[0]
        plan_info[p.Name] = (p.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count > 0, plan_opt.Constraints.Count > 0 or plan_opt.Objective is not None)

    # Offer to copy the plan if the plan is approved
    if plan.Review is not None and plan.Review.ApprovalStatus == 'Approved':
        res = MessageBox.Show('The current plan is approved, it cannot be modified. Would you like to modify a copy of the plan?', 'Plan Is Approved', MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit()

        # Copy plan and switch to the copy
        plan = copy_plan_without_changes()
        patient.Save()
        plan.SetCurrent()

    # Get user input from GUI
    form = CopyOptStuffForm(plan.Name, plan_info)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()

    # Clear existing clinical goals if the user so desires
    if form.clear_existing_goals_cb.Checked:
        with CompositeAction('Clear Clinical Goals'):
            while plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count > 0:
                plan.TreatmentCourse.EvaluationSetup.DeleteClinicalGoal(FunctionToRemove=plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions[0])

    # Clear existing objectives and constraints if the user so desires
    if form.clear_existing_objs_cb.Checked:
        with CompositeAction('Clear Objectives and Constraints'):
            while plan.PlanOptimizations[0].Constraints.Count > 0:
                plan.PlanOptimizations[0].Constraints[0].DeleteFunction()
            while plan.PlanOptimizations[0].Objective.ConstituentFunctions.Count > 0:
                plan.PlanOptimizations[0].Objective.ConstituentFunctions[0].DeleteFunction()

    # Copy the goals/objectives
    copy_goals_from = sorted(form.copy_goals_from, key=lambda x: x.lower())
    copy_objs_from = sorted(form.copy_objs_from, key=lambda x: x.lower())
    for plan_name in copy_goals_from:
        copy_clinical_goals(case.TreatmentPlans[plan_name], plan)
    for plan_name in copy_objs_from:
        copy_objectives_and_constraints(case.TreatmentPlans[plan_name], plan)
    if form.copy_params_from is not None:
        copy_opt_params(case.TreatmentPlans[form.copy_params_from], plan)
