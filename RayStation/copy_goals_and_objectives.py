import clr
import sys

from connect import *

clr.AddReference('System.Windows.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Windows.drawing import *
from System.Windows.Forms import *

sys.path.append(r'T:\Physics\KW\med-phys-scripts\RayStation')
from copy_plan_without_changes import copy_plan_without_changes


class CopyGoalsAndObjectivesForm(Form):
    def __init__(self, plans_w_goals, plans_w_objs):
        self.Text = 'Add Clinical Goals'  # Form title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

        y = 15  # Vertical coordinate of next control

        clear_panel = Panel()
        clear_panel.Dock = DockStyle.Top
        self.Controls.Add(clear_panel)

        self.clear_existing_goals_cb = CheckBox()
        self.clear_existing_goals_cb.AutoSize = True
        self.clear_existing_goals_cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.clear_existing_goals_cb.Text = 'Clear existing clinical goals'
        clear_panel.Controls.Add(self.clear_existing_goals_cb)

        self.clear_existing_objs_cb = CheckBox()
        self.clear_existing_objs_cb.AutoSize = True
        self.clear_existing_objs_cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.clear_existing_objs_cb.Anchor = AnchorStyles.Right
        self.clear_existing_objs_cb.Text = 'Clear existing objectives & constraints'
        clear_panel.Controls.Add(self.clear_existing_objs_cb)

        plan_names = sorted(list(set(plans_w_goals + plans_w_objs)), key=lambda x: x.lower())

        self.data_grid_view = DataGridView()
        self.Controls.Add(self.data_grid_view)
        self.data_grid_view.AutoGenerateColumns = False
        self.data_grid_view.RowHeadersVisible = False
        self.data_grid_view.ColumnCount = 3

        plan_col = DataGridViewColumn()
        plan_col.HeaderText = 'Plan'
        plan_col.ReadOnly = True
        self.data_grid_view.Columns.Add(plan_col)

        copy_goals_col = DataGridViewCheckBoxColumn()
        copy_goals_col.HeaderText = 'Copy linical Goals'
        copy_goals_col.ThreeState = True
        self.data_grid_view.Columns.Add(copy_goals_col)

        copy_objs_col = DataGridViewCheckBoxColumn()
        copy_objs_column.HeaderText = 'Copy Objectives & Constraints'
        copy_objs_col.ThreeState = True
        self.data_grid_view.Columns.Add(copy_objs_col)

        for i, name in enumerate(plan_names):
            has_goals = name in plans_w_goals
            has_objs = name in plans_w_objs
            
            row = [name, has_goals, has_objs]
            self.data_grid_view.Items.Add(row)
            if not has_goals:
                copy_goals_col.Cells[i].Value = False
                copy_goals_col.Cells[i].Style.ForeColor = Color.DarkGray
                copy_goals_col.Cells[i].ReadOnly = True
            if not has_objs:
                copy_objs_col.Cells[i].Value = False
                copy_objs_col.Cells[i].Style.ForeColor = Color.DarkGray
                copy_objs_col.Cells[i].ReadOnly = True


def copy_goals_and_objectives():
    try:
        patient = get_current('Patient')
        try:
            plan = get_current('Plan')
        except:
            MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
            sys.exit()
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()

    plans_w_goals, plans_w_objs = [], []
    for p in case.TreatmentPlans:
        if p.Equals(plan):
            continue
        if p.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count > 0:
            plans_w_goals.append(p.Name)
        if p.PlanOptimizations[0].Constraints.Count > 0 or p.PlanOptimizations[0].Objective.ConstituentFunctions.Count > 0:
            plans_w_objs.append(p.Name)

    if not plans_w_goals and not plans_w_objs:
        MessageBox.Show('There are no other plans in the current case that have either clinical goals or objectives and constraints. Click OK to abort the script.', 'No Plans to Copy from')
        sys.exit()

    if plan.Review is not None and plan.Review.ApprovalStatus == 'Approved':
        res = MessageBox.Show('The current plan is approved, so clinical goals cannot be added. Would you like to add clinical goals to a copy of the plan?', 'Plan Is Approved', MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit()
        
        plans_w_goals.remove(plan)
        plans_w_goals.remove(plan)
        plan = copy_plan_without_changes()
        patient.Save()
        plan.SetCurrent()

    form = CopyGoalsAndObjectivesForm(plans_w_goals, plans_w_obj)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()

    if form.clear_existing_goals_cb.Checked:
        with CompositeAction('Clear Clinical Goals'):
            while plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count > 0:
                plan.TreatmentCourse.EvaluationSetup.DeleteClinicalGoal(FunctionToRemove=plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions[0])

    if form.clear_existing_obj_cb.Checked:
        with CompositeAction('Clear Objectives and Constraints'):
            while plan.PlanOptimizations[0].Constraints.Count > 0:
                plan.PlanOptimizations[0].Constraints[0].DeleteFunction()
            while plan.PlanOptimizations[0].Objective.ConstituentFunctions.Count > 0:
                plan.PlanOptimizations[0].Objective.ConstituentFunctions[0].DeleteFunction()

    #for 