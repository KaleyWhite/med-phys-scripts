"""NOTE: I minimally modified this script from CopyPlanToNewCTOrMergeBeamSets from RaySearch support.
My biggest changes:
- Convert from IronPython to CPython and WPF to WinForms
- Update from RayStation version 4 or 5, to 11A.
- Clean up styling: meaningful variable names, greater modularization, grammatically correct comments
"""

import clr
import sys
from typing import Optional

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')

from System import EventArgs, EventHandler  # For type hints
from System.Drawing import *
from System.Windows.Forms import *


BK_COLOR = ColorTranslator.FromHtml('#FFE5E5E5')  # TabPage background color
X = 20  # Default x-ccordinate of a control


# Classes to standardize styling/behavior of WinForms controls used multiple times in the form
# Each constructor sets common styles/behavior and adds the control to the given parent

class MyLabel(Label):
    def __init__(self, parent: Control, y: int, text: str = '') -> None:
        super().__init__()

        self.AutoSize = True
        self.Location = Point(X, y)
        self.Text = text
        parent.Controls.Add(self)


class MyComboBox(ComboBox):
    def __init__(self, parent: Control, y: int, on_selection_chg: EventHandler) -> None:
        super().__init__()
        print(type(on_selection_chg))
        self.DropDownStyle = ComboBoxStyle.DropDownList  # User cannot type in the control
        self.Height = 90
        self.Width = 187
        self.SelectedIndexChanged += on_selection_chg
        self.Location = Point(X, y)
        parent.Controls.Add(self)


class MyButton(Button):
    def __init__(self, parent: Control, x: int, on_click: EventHandler) -> None:
        super().__init__()

        self.Width = 75
        self.Click += on_click
        self.Location = Point(x, 272)
        parent.Controls.Add(self)


class SubmitButton(MyButton):
    def __init__(self, parent: Control, on_click: EventHandler) -> None:
        super().__init__(parent, 36, on_click)

        self.Text = 'Submit'
        self.Enabled = False


class CancelButton(MyButton):
    def __init__(self, parent: Control, on_click: EventHandler) -> None:
        super().__init__(parent, 116, on_click)

        self.Text = 'Cancel'


class MyWindow(Form):
    """Form that allows the user to set parameters for copying a plan or beam set(s)"""

    def __init__(self, case: PyScriptObject) -> None:
        self.case = case

        # Set form style
        self.Height = 371
        self.Width = 235
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # Disallow user resizing
        self.StartPosition = FormStartPosition.CenterScreen  # Launch window in middle of screen

        # Tab container
        tab_ctrl = TabControl()
        tab_ctrl.Height = 336
        tab_ctrl.Location = Point(0, 2)
        tab_ctrl.Width = 292
        self.Controls.Add(tab_ctrl)

        # Individual tabs
        tab_ctrl.Controls.Add(self.add_copy_tab())
        tab_ctrl.Controls.Add(self.add_merge_tab())
        
        self.populate_dataset_list()
        self.mode = ""

    def add_copy_tab(self) -> TabPage:
        # Returns the "Copy Plan To New CT" tab
    
        tab = TabPage('Copy Plan To New CT')
        tab.BackColor = BK_COLOR

        # Plan to copy
        MyLabel(tab, 21, 'Plan to Copy')
        self.plan_to_copy_cb = MyComboBox(tab, 52, self.plan_to_copy_cb_selection_chged)

        # New plan name
        MyLabel(tab, 79, 'New Plan Name')

        self.new_plan_name_tb = TextBox()
        self.new_plan_name_tb.Height = 23
        self.new_plan_name_tb.WordWrap = True
        self.new_plan_name_tb.Width = 187
        self.new_plan_name_tb.TextChanged += self.set_submit_1_enabled
        self.new_plan_name_tb.Location = Point(20, 110)
        tab.Controls.Add(self.new_plan_name_tb)

        # New dataset
        MyLabel(tab, 138, 'New Dataset')
        self.new_dataset_lb = MyComboBox(tab, 169, self.set_submit_1_enabled)

        # "Submit" button
        self.submit_btn = SubmitButton(tab, self.submit_btn_click)

        # "Cancel" button
        CancelButton(tab, self.cancel_btn_click)

        return tab

    def add_merge_tab(self) -> TabPage:
        # Returns the "Copy Beam Set" tab

        tab = TabPage('Copy Beam Set')
        tab.BackColor = BK_COLOR

        # Plan to copy from
        MyLabel(tab, 21, 'Plan to Copy From')
        self.plan_to_copy_from_cb = MyComboBox(tab, 52, self.plan_to_copy_from_selection_chged)

        # Plan to copy to
        MyLabel(tab, 79, 'Plan to Copy To')
        self.plan_to_copy_to_cb = MyComboBox(tab, 110, self.plan_to_copy_to_cb_selection_chged)

        # Beam set(s) to copy
        MyLabel(tab, 138, 'BeamSet(s) to Copy')

        self.beam_sets_lb = ListBox()
        self.beam_sets_lb.Location = Point(20, 169)
        self.beam_sets_lb.Height = 90
        self.beam_sets_lb.Width = 187
        self.beam_sets_lb.SelectedIndexChanged += self.beam_sets_lb_selection_chged
        self.beam_sets_lb.SelectionMode = SelectionMode.MultiExtended
        tab.Controls.Add(self.beam_sets_lb)

        # "Submit" button
        self.submit_btn_2 = SubmitButton(tab, self.submit_2_click)

        # "Cancel" button
        CancelButton(tab, self.cancel_btn_click)

        return tab

    def submit_btn_click(self, sender: SubmitButton, e: EventArgs) -> None:
        self.mode = 'copy'
        self.DialogResult = DialogResult.OK

    def cancel_btn_click(self, sender: CancelButton, e: EventArgs) -> None:
        self.DialogResult = DialogResult.Cancel
    
    def plan_to_copy_cb_selection_chged(self, sender: MyComboBox, e: EventArgs) -> None:
        self.set_submit_1_enabled()

    def plan_to_copy_to_cb_selection_chged(self, sender: MyComboBox, e: EventArgs) -> None:
        self.set_submit_2_enabled()

    def plan_to_copy_from_selection_chged(self, sender: MyComboBox, e: EventArgs) -> None:
        # Populates beam sets ComboBox with the beam set names in the selected plan to copy from
        if self.plan_to_copy_from_cb.Text != '':
            beam_set_names = [beam_set.DicomPlanLabel for beam_set in self.case.TreatmentPlans[self.plan_to_copy_from_cb.SelectedItem].BeamSets]
            self.beam_sets_lb.Items.Clear()
            for beam_set_name in sorted(beam_set_names):
                self.beam_sets_lb.Items.Add(beam_set_name)
        self.set_submit_2_enabled()

    def beam_sets_lb_selection_chged(self, sender: ListBox, e: EventArgs) -> None:
        self.set_submit_2_enabled()

    def submit_2_click(self, sender: SubmitButton, e: EventArgs) -> None:
        self.mode = 'merge'
        self.DialogResult = DialogResult.OK

    def set_submit_1_enabled(self, sender: Optional[SubmitButton], e: Optional[EventArgs]) -> None:
        # "Submit" button on copy plan tab is enabled if and only if all inputs on the tab are provided
        self.submit_btn.Enabled = self.new_plan_name_tb.Text != '' and self.new_dataset_lb.SelectedItem is not None and self.plan_to_copy_cb.SelectedItem is not None

    def set_submit_2_enabled(self, sender: Optional[SubmitButton], e: Optional[EventArgs]) -> None:
        # "Submit" button on copy pbeam set(s) tab is enabled if and only if all inputs on the tab are provided
        self.submit_btn_2.Enabled = self.plan_to_copy_to_cb.SelectedItem is not None and self.beam_sets_lb.SelectedItems.Count > 0 and self.plan_to_copy_from_cb.SelectedItem is not None

    def populate_dataset_list(self) -> None:
        # Initially populates the ComboBoxes on both tabs
        plans = [p.Name for p in self.case.TreatmentPlans]
        self.plan_to_copy_cb.Items.AddRange(plans)
        self.plan_to_copy_to_cb.Items.AddRange(plans)
        self.plan_to_copy_from_cb.Items.AddRange(plans)
        exams = [exam.Name for exam in self.case.Examinations]
        for e in sorted(exams):
            self.new_dataset_lb.Items.Add(e)


def tx_technique(bs: PyScriptObject) -> Optional[str]:
    # Returns a beam set's treatment technique, or None if it cannot be determined
    if bs.Modality == 'Photons':
        if bs.PlanGenerationTechnique == 'Imrt':
            if bs.DeliveryTechnique == 'SMLC':
                return 'SMLC'
            if bs.DeliveryTechnique == 'DynamicArc':
                return 'VMAT'
            if bs.DeliveryTechnique == 'DMLC':
                return 'DMLC'
        elif bs.PlanGenerationTechnique == 'Conformal':
            if bs.DeliveryTechnique == 'SMLC':
                return 'SMLC' # Changed from 'Conformal'. Failing with forward plans.
            if bs.DeliveryTechnique == 'Arc':
                return 'Conformal Arc'
    elif bs.Modality == 'Electrons' and bs.PlanGenerationTechnique == 'Conformal' and bs.DeliveryTechnique == 'SMLC':
        return 'ApplicatorAndCutout'


def create_ctrl_points(bs: PyScriptObject) -> None:
    # Copies control points from old beam set to new beam set
    # Since optimization must occur first, we optimize the single uniform dose objective for a dummy PTV
    # `bs` is the beam set that we are copying

    # Get plan optimizations for this beam set
    plan_opt = next(opt for opt in original_plan.PlanOptimizations if opt.OptimizedBeamSets.Count == 1 and opt.OptimizedBeamSets[0].DicomPlanLabel == bs.DicomPlanLabel)
    replan_opt = next(opt for opt in replan.PlanOptimizations if opt.OptimizedBeamSets.Count == 1 and opt.OptimizedBeamSets[0].DicomPlanLabel == replan_bs.DicomPlanLabel)

    # Remove an ROI named 'dummyPTV' if it exists
    try:
        roi = case.PatientModel.RegionsOfInterest['dummyPTV']
        roi.DeleteRoi()
    except:
        pass

    # Create OUR dummyPTV ROI
    with CompositeAction('Create dummy'):
        dummy = case.PatientModel.CreateRoi(Name='dummyPTV', Color='Red'
                , Type='Ptv', TissueName=None, RoiMaterial=None)

        dummy.CreateSphereGeometry(Radius=2,
                                   Examination=examination[0],
                                   Center=isocenter)

    # Remove any existing optimization objectives
    if replan_opt.Objective is not None:
        for f in replan_opt.Objective.ConstituentFunctions:
            f.DeleteFunction()

    # Add OUR optimization objectives
    with CompositeAction('Add objective'):
        function = replan_opt.AddOptimizationFunction(
            FunctionType='UniformDose',
            RoiName='dummyPTV',
            IsConstraint=False,
            RestrictAllBeamsIndividually=False,
            RestrictToBeam=None,
            IsRobust=False,
            RestrictToBeamSet=None,
            )

        function.DoseFunctionParameters.DoseLevel = bs.FractionDose.ForBeamSet.Prescription.PrimaryPrescriptionDoseReference.DoseValue
        function.DoseFunctionParameters.Weight = 90

    # Set the arc spacing and create the segments
    with CompositeAction('Dummy optimization'):
        replan_opt.OptimizationParameters.Algorithm.MaxNumberOfIterations = 1
        replan_opt.OptimizationParameters.DoseCalculation.IterationsInPreparationsPhase = \
            0
        for i, bs in \
            enumerate(plan_opt.OptimizationParameters.TreatmentSetupSettings[0].BeamSettings):
            replan_opt.OptimizationParameters.TreatmentSetupSettings[0].BeamSettings[i].ArcConversionPropertiesPerBeam.FinalArcGantrySpacing = \
                bs.ArcConversionPropertiesPerBeam.FinalArcGantrySpacing
        replan_opt.RunOptimization()


def copy_plan_to_new_ct_or_merge_beam_sets() -> None:
    """Copies a photon or electron plan to another exam or merges beam sets for plans on the same exam.
    
    The new exam must have an external geometry.
    """

    # Check if a patient is loaded
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('No patient is open. Click Ok to abort the script.', 'No Patient Open')
        sys.exit()

    # Check the version of RayStation
    case = get_current('Case')

    dialog = MyWindow(case)
    dialog.ShowDialog()
    
    if dialog.DialogResult != DialogResult.OK:
        sys.exit()

    reg = None

    # Are we copying or merging?
    if dialog.mode == 'copy':
        original_plan = case.TreatmentPlans[dialog.plan_to_copy_cb.SelectedItem]
        replan_ct_name = dialog.new_dataset_lb.SelectedItem
        replan_name = dialog.new_plan_name_tb.Text

        # Ensure external geometry on new exam so we can define dose grid
        if not any(geom.OfRoi.Type == 'External' and geom.HasContours() for geom in case.PatientModel.StructureSets[replan_ct_name].RoiGeometries):
            MessageBox.Show('There is no external geometry on the replan exam. Click OK to abort the script.', 'No External Geometry')
            sys.exit()

        # Unique replan name
        plan_names = [plan.Name for plan in case.TreatmentPlans]
        if replan_name in plan_names:
            base_replan_name = replan_name
            i = 1
            while replan_name in plan_names:
                replan_name = f'{base_replan_name} ({i})'
                i += 1

        # Old exam
        examination = [e for e in case.Examinations if e.Name == replan_ct_name]
        if original_plan.TreatmentCourse.TotalDose.DoseValues is None:
            MessageBox.Show('The plan "' + original_plan.Name + '" has no dose. Click OK to abort the script.', 'No Dose')
            sys.exit(1)
        original_plan_ct_name = \
            original_plan.GetTotalDoseStructureSet().OnExamination.Name
            
        # Does a registration exist? We will use this to maintain copy vs merge state
        reg = case.GetTransformForExaminations(
            FromExamination=original_plan_ct_name,
            ToExamination=replan_ct_name
        )
        if reg is None:
            MessageBox.Show('There is no registration between the old exam ("' + original_plan_ct_name + '") and the new exam ("' + replan_ct_name + '"). Click OK to abort the script.', 'No Registration')
            sys.exit()

        # Create the copy plan
        replan = case.AddNewPlan(PlanName=replan_name, PlannedBy='', Comment='',
                                    ExaminationName=replan_ct_name,
                                    AllowDuplicateNames=False)

        beam_sets = original_plan.BeamSets

    # If not copy, then must be merge
    else:
        original_plan = case.TreatmentPlans[dialog.plan_to_copy_from_cb.SelectedItem]
        beam_sets = [original_plan.BeamSets[item] for item in dialog.beam_sets_lb.SelectedItems]
        replan = case.TreatmentPlans[dialog.plan_to_copy_to_cb.SelectedItem]
        replan_ct_name = replan.GetTotalDoseStructureSet().OnExamination.Name

    # Copy beam sets
    for bs in beam_sets:
        name = bs.DicomPlanLabel
        machine_name = bs.MachineReference.MachineName
        modality = bs.Modality
        pt_position = bs.PatientPosition
        fractions = bs.FractionationPattern.NumberOfFractions
        technique = tx_technique(bs)

        # Unique name for beam set in replan
        replan_bs_names = [replan_bs.DicomPlanLabel for replan_bs in replan.BeamSets]
        if name in replan_bs_names:
            base_name = name
            i = 1
            while name in replan_bs_names:
                name = f'{base_name} ({i})'
                i += 1

        # Create the beamset
        replan_bs = replan.AddNewBeamSet(
            Name=name,
            ExaminationName=replan_ct_name,
            MachineName=machine_name,
            Modality=modality,
            TreatmentTechnique=technique,
            PatientPosition=pt_position,
            NumberOfFractions=fractions,
            CreateSetupBeams=True,
            Comment='From ' + bs.BeamSetIdentifier(),
            )  

        # Set dose grid
        dg = bs.GetDoseGrid()
        replan_bs.SetDefaultDoseGrid(VoxelSize={'x': dg.VoxelSize.x,
                                    'y': dg.VoxelSize.y, 'z': dg.VoxelSize.z})

        if modality == 'Electrons':
            for beam in bs.Beams:
                electron_beam_args = {
                    'BeamQualityId' : beam.BeamQualityId,
                    'Name' : beam.Name,
                    'GantryAngle' : beam.GantryAngle,
                    'CouchRotationAngle' : beam.CouchRotationAngle,
                    'ApplicatorName' : beam.Applicator.ElectronApplicatorName,
                    'InsertName' : beam.Applicator.Insert.Name,
                    'IsAddCutoutChecked' : True
                }

                iso = beam.Isocenter.Position
                isocenter = {
                    'x': iso.x,
                    'y': iso.y,
                    'z': iso.z
                }

                # Check whether copy or merge
                if reg is not None:
                    iso = case.TransformPointFromExaminationToExamination(
                        FromExamination=original_plan_ct_name,
                        ToExamination=replan_ct_name,
                        Point=isocenter
                    )
                    isocenter = {
                        'x': iso.x,
                        'y': iso.y,
                        'z': iso.z
                    }
                electron_beam_args['IsocenterData'] = bs.CreateDefaultIsocenterData(Position=isocenter)

                new_beam = replan_bs.CreateElectronBeam(**electron_beam_args)
                contour = [{'x': c.x, 'y': c.y} for c in beam.Applicator.Insert.Contour]
                new_beam.Applicator.Insert.Contour = contour
                new_beam.BeamMU = beam.BeamMU

        # Not an electron plan, so is it non-VMAT photon?
        elif technique != 'VMAT':
            for beam in bs.Beams:
                photon_beam_args = {
                    'BeamQualityId': beam.BeamQualityId,
                    'Name': beam.Name,
                    'GantryAngle': beam.GantryAngle,
                    'CouchRotationAngle': beam.CouchRotationAngle,
                    'CollimatorAngle': beam.InitialCollimatorAngle,
                    }

                iso = beam.Isocenter.Position
                isocenter = {
                    'x': iso.x,
                    'y': iso.y,
                    'z': iso.z
                }

                # Check whether copy or merge
                if reg is not None:
                    iso = case.TransformPointFromExaminationToExamination(
                        FromExamination=original_plan_ct_name,
                        ToExamination=replan_ct_name,
                        Point=isocenter
                    )
                    isocenter = {
                        'x': iso.x,
                        'y': iso.y,
                        'z': iso.z
                    }
                photon_beam_args['IsocenterData'] = bs.CreateDefaultIsocenterData(Position=isocenter)

                new_beam = replan_bs.CreatePhotonBeam(**photon_beam_args)
                for i, s in enumerate(beam.Segments):
                    new_beam.CreateRectangularField()
                    new_beam.Segments[i].LeafPositions = s.LeafPositions
                    new_beam.Segments[i].JawPositions = s.JawPositions
                    new_beam.BeamMU = beam.BeamMU
                    new_beam.Segments[i].RelativeWeight = s.RelativeWeight

        # Must be a VMAT
        else:
            for beam in bs.Beams:
                arc_beam_args = {
                    'ArcStopGantryAngle': beam.ArcStopGantryAngle,
                    'ArcRotationDirection': beam.ArcRotationDirection,
                    'BeamQualityId': beam.BeamQualityId,
                    'Name': beam.Name,
                    'GantryAngle': beam.GantryAngle,
                    'CouchRotationAngle': beam.CouchRotationAngle,
                    'CollimatorAngle': beam.InitialCollimatorAngle,
                    }
                
                iso = beam.Isocenter.Position
                isocenter = {
                    'x': iso.x,
                    'y': iso.y,
                    'z': iso.z
                }

                # check if copy or merge
                if reg is not None:
                    iso = case.TransformPointFromExaminationToExamination(
                        FromExamination=original_plan_ct_name,
                        ToExamination=replan_ct_name,
                        Point=isocenter
                    )
                    isocenter = {
                        'x': iso.x,
                        'y': iso.y,
                        'z': iso.z
                    }
                arc_beam_args['IsocenterData'] = bs.CreateDefaultIsocenterData(Position=isocenter)

                new_beam = replan_bs.CreateArcBeam(**arc_beam_args)
            
            # Can't directly create control points
            create_ctrl_points(bs)

            # Copy the old segments to the new segments
            for i, beam in enumerate(bs.Beams):
                replan_bs.Beams[i].BeamMU = beam.BeamMU

                for j, s in enumerate(beam.Segments):
                    replan_bs.Beams[i].Segments[j].LeafPositions = \
                        s.LeafPositions
                    replan_bs.Beams[i].Segments[j].JawPositions = s.JawPositions
                    replan_bs.Beams[i].Segments[j].DoseRate = s.DoseRate
                    replan_bs.Beams[i].Segments[j].RelativeWeight = \
                        s.RelativeWeight

    # Compute the new beams
    e_beam_sets = []
    for bs in replan.BeamSets:

        if bs.Modality == 'Electrons':
            e_beam_sets.append(bs.DicomPlanLabel)
        elif bs.Beams.Count != 0:
            bs.ComputeDose(ComputeBeamDoses=True, DoseAlgorithm=bs.AccurateDoseAlgorithm.DoseAlgorithm, ForceRecompute=True)
    if e_beam_sets:
        MessageBox.Show('The following are electron beam sets, so histories and Rx must be set before computing: ' + ', '.join(e_beam_sets), 'Cannot Compute e- Beam Sets')
