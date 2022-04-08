import clr
from collections import OrderedDict
import re
import sys
from typing import Dict, List

import pandas as pd  # Clinical Goals template data is read in as DataFrame
from connect import *

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')

from System.Drawing import *
from System.Windows.Forms import *

sys.path.append(r'T:\Physics\KW\med-phys-scripts\RayStation')
import copy_plan_without_changes

CLINICAL_GOALS_FILEPATH = r'T:\Physics\KW\med-phys-spreadsheets\Clinical Goals.xlsx'

# General and specific ROI names
# Some goals in templates have 'general' ROI names (e.g., 'Stomach and intestines') that correspond to multiple possible 'specific' ROI names
# Specific names from original TG-263 nomenclature spreadsheet (not the one w/ CRMC-created names)
SPECIFIC_ROIS = {
    'Aorta and major vessels': ['A_Aorta', 'A_Aorta_Asc', 'A_Aorta_Desc', 'A_Coronary', 'A_Pulmonary', 'GreatVes', 'V_Pulmonary', 'V_Venacava', 'V_Venacava_I', 'V_Venacava_S'],
    'Bag_Bowel': ['Bag_Bowel', 'Bowel', 'Bowel_Small', 'Colon_Sigmoid', 'Spc_Bowel'],
}
SPECIFIC_ROIS['Stomach and intestines'] = SPECIFIC_ROIS['Bag_Bowel'] + ['Stomach']


def format_warnings(warnings_dict: Dict[str, List[str]]):
    # Helper function that nicely formats a dictionary of strings into one long string
    # E.g., format_warnings({'A': ['B', 'C'], 'D': ['E']}) ->
    #     -  A
    #         -  B
    #         -  C
    #     -  D
    #         - E

    warnings_str = ''
    for k, v in warnings_dict.items():
        warnings_str += '\n\t-  ' + k
        for val in sorted(list(set(v))):
            warnings_str += '\n\t\t-  ' + val
    warnings_str += '\n'
    return warnings_str


class AddClinicalGoalsForm(Form):
    """Form that allows user to select the clinical goals tenplates to apply"""

    def __init__(self, templates: List[str], default_template: str):
        """Initializes an AddClinicalGoalsForm object.

        Args:
            data: The clinical goals sets from the spreadsheet
            default_template: The name of the default selected template
        """

        self.Text = 'Add Clinical Goals'  # Form title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

        y = 15  # Vertical coordinate of next control

        # Instructions
        lbl = Label()
        lbl.AutoSize = True
        lbl.AutoSizeMode = AutoSizeMode.GrowAndShrink
        lbl.Text = 'Select the clinical goals template(s) to apply.'
        self.Controls.Add(lbl)
        y += lbl.Height + 5
        
        # Remove existing goals?
        self.clear_existing_cb = CheckBox()  
        self.clear_existing_cb.AutoSize = True
        self.clear_existing_cb.Checked = True
        self.clear_existing_cb.Location = Point(15, y)
        self.clear_existing_cb.Text = 'Clear existing goals'
        self.Controls.Add(self.clear_existing_cb)
        y += self.clear_existing_cb.Height + 15

        # Templates choices
        self.choose_templates_lb = ListBox()
        self.choose_templates_lb.Location = Point(15, y)
        self.choose_templates_lb.SelectionMode = SelectionMode.MultiExtended
        self.choose_templates_lb.Visible = True

        self.choose_templates_lb.HorizontalScrollbar = False
        self.choose_templates_lb.VerticalScrollbar = False
        
        self.choose_templates_lb.Items.AddRange(templates)

        selected_idx = templates.index(default_template)
        self.choose_templates_lb.SetSelected(selected_idx, True)
        
        self.choose_templates_lb.Height = self.choose_templates_lb.PreferredHeight
        self.choose_templates_lb.Width = max(TextRenderer.MeasureText(template, self.choose_templates_lb.Font).Width for template in templates)
        
        self.Controls.Add(self.choose_templates_lb)
        y += self.choose_templates_lb.Height + 15

        self.ok = Button()
        self.ok.Click += self.ok_clicked
        self.ok.Location = Point(15, y)
        self.ok.Text = 'OK'
        self.AcceptButton = self.ok
        self.Controls.Add(self.ok)

    def set_ok_enabled(self, sender=None, event=None):
        # Enable or disable 'OK' button
        # Enable only if at least one template is selected
        self.ok.Enabled = self.choose_templates_lb.SelectedItems.Count > 0

    def ok_clicked(self, sender, event):
        # Event handler for 'OK' button click
        self.DialogResult = DialogResult.OK


def add_clinical_goals():
    """Applies clinical goals template(s) from spreadsheet, to the current plan
    
    User chooses templates from a GUI
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
        MessageBox.Show('The current plan has no beam sets. Click OK to abort the script.', 'No Beam Sets')
        sys.exit()
    patient = get_current('Patient')
    case = get_current('Case')

    if plan.Review is not None and plan.Review.ApprovalStatus == 'Approved':
        res = MessageBox.Show('Plan is approved, so clinical goals cannot be added. Would you like to add goals to a copy of the plan?', 'Plan Is Approved', MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit()

        # Copy plan and switch to new plan and corresponding beam set
        new_plan = copy_plan_without_changes.copy_plan_without_changes()
        patient.Save()
        new_plan.SetCurrent()
        plan = get_current('Plan')
        plan.BeamSets[beam_set.DicomPlanLabel].SetCurrent()
        beam_set = get_current('BeamSet')

    struct_set = plan.GetTotalDoseStructureSet()  # Geometries on the planning exam

    # Ensure that this is a photon beam set
    if beam_set.Modality != 'Photons':
        MessageBox.Show('The current beam set is not a photon beam set. Click OK to abort the script.', 'Incorrect Modality')
        sys.exit()  # Exit with an error

    # Ensure that beam set machine is commissioned
    machine = beam_set.MachineReference.MachineName
    if get_current('MachineDB').GetTreatmentMachine(machineName=machine) is None:
        MessageBox.Show('Machine "' + machine + '" is uncommissioned. Click OK to abort the script.', 'Uncommissioned Machine')
        sys.exit(1)

    warnings = ''  # Warnings to display at end of script (if there were any)

    # Rx and # fx
    fx = rx_val = None
    rx = beam_set.Prescription
    if rx is not None:
        rx = rx.PrimaryPrescriptionDoseReference
        if rx is not None:
            rx_val = rx.DoseValue
            fx = beam_set.FractionationPattern
            if fx is not None:
                fx = fx.NumberOfFractions

    # Default selected template depends on current plan type
    default_template = 'Mobius Conventional'
    if fx is not None and rx_val is not None and beam_set.PlanGenerationTechnique == 'Imrt' and beam_set.DeliveryTechnique == 'DynamicArc' and rx_val / fx >= 600:  # SABR/SBRT/SRS >=6 Gy/fx
        is_sabr = True
        if fx == 1:
            default_template = 'Mobius 1 Fx SRS'
        elif fx == 3:
            default_template = 'Mobius 3 Fx SBRT'
        elif fx == 5:
            default_template = 'Mobius 5 Fx SBRT'
    else:
        is_sabr = False

    # Read data from "Clinical Goals" spreadsheet
    # Read all sheets, ignoring "Planning Priority" and "Visualization Priority" columns
    # Dictionary of sheet name : DataFrame
    # First sheet is "How to Use"
    rdr = pd.ExcelFile(CLINICAL_GOALS_FILEPATH)
    sheets = sorted(sheet for sheet in rdr.sheet_names[1:] if not sheet.endswith('DNU'))
    data = pd.read_excel(rdr, sheet_name=sheets, engine='openpyxl', usecols=['ROI', 'Goal', 'Notes'])  # Default xlrd engine does not support xlsx
    
    # Get options from user
    form = AddClinicalGoalsForm(sheets, default_template)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:  # 'OK' button was not clicked
        sys.exit()
    clear_existing = form.clear_existing_cb.Checked
    template_names = list(form.choose_templates_lb.SelectedItems)

    # Text of checked RadioButtons
    # Clear existing Clinical Goals
    if clear_existing:
        with CompositeAction('Clear Clinical Goals'):
            while plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count > 0:
                plan.TreatmentCourse.EvaluationSetup.DeleteClinicalGoal(FunctionToRemove=plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions[0])

    # If Rx is specified, add Dmax goal
    if rx_val is not None:
        try:
            ext = next(roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == 'External')  # Select external ROI
            # If there is an external (there will only be one), add Dmax goal
            d_max = 1.25 if is_sabr else 1.1
            try:
                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ext.Name, GoalCriteria='AtMost', GoalType='DoseAtAbsoluteVolume', ParameterValue=0.03, AcceptanceLevel=d_max * rx_val)  # e.g., D0.03 < 4400 for 4000 cGy non-SBRT plan
            except:  # Clinical goal already exists
                pass
        except StopIteration:
            pass
        
        for dose_rx in beam_set.Prescription.PrescriptionDoseReferences:
            # If Rx is to volume of PTV, add PTV D95%, V95%, D100%, and V100%
            if dose_rx.PrescriptionType == 'DoseAtVolume' and dose_rx.OnStructure.Type == 'Ptv':
                dose_rx_val = dose_rx.DoseValue
                ptv = dose_rx.OnStructure
            
                try:
                    plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ptv.Name, GoalCriteria='AtLeast', GoalType='DoseAtVolume', ParameterValue=0.95, AcceptanceLevel=dose_rx_val)  # D95% >= Rx
                except:  # Clinical goal already exists
                    pass
                try:
                    plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ptv.Name, GoalCriteria='AtLeast', GoalType='VolumeAtDose', ParameterValue=0.95 * dose_rx_val, AcceptanceLevel=1)  # V95% >= 100%
                except:
                    pass
                try:
                    plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ptv.Name, GoalCriteria='AtLeast', GoalType='VolumeAtDose', ParameterValue=dose_rx_val, AcceptanceLevel=0.95)  # V100% >= 95%
                except:
                    pass
                try:
                    plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ptv.Name, GoalCriteria='AtLeast', GoalType='DoseAtVolume', ParameterValue=1, AcceptanceLevel=0.95 * dose_rx_val)  # D100% >= 95%
                except:
                    pass
                # If PTV is derived from CTV, add CTV D100% and V100%
                if ptv.DerivedRoiExpression is not None:
                    for r in struct_set.RoiGeometries[ptv.Name].GetDependentRois():
                        if case.PatientModel.RegionsOfInterest[r].Type == 'Ctv':
                            try:
                                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=r, GoalCriteria='AtLeast', GoalType='DoseAtVolume', ParameterValue=1, AcceptanceLevel=dose_rx_val)  # D100% >= 100%
                            except:
                                pass
                            try:
                                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=r, GoalCriteria='AtLeast', GoalType='VolumeAtDose', ParameterValue=dose_rx_val, AcceptanceLevel=1)  # V100% >= 100%
                            except:
                                pass

    # Information that will be displayed as warnings later
    # All invalid goals are in format '<ROI name>: <goal>', e.g., 'Liver:  V21 Gy < (v-700) cc', except ipsi/contra goals ('<ROI name>: <goal> <Ipsilateral|Contralateral>')
    invalid_goals = OrderedDict()  # Goals in template that are in an invalid format
    empty_spare = OrderedDict()  # Volume-to-spare goals that cannot be added due to empty geometry
    lg_spare_vol = OrderedDict()  # Volume to spare is larger than ROI volume
    no_ipsi_contra = OrderedDict()  # Whether or not ipsilateral/contralateral goals could not be added due to indeterminable Rx side
    no_nodal_ptv = OrderedDict()  # If a nodal PTV does not exist, goals for a nodal PTV 

    ## Determine Rx side, used when adding ispi/contra goals

    if hasattr(rx, 'OnStructure'):
        struct = rx.OnStructure
        if struct.OrganData is not None and struct.OrganData.OrganType == 'Target':  # Rx is to ROI
            rx_ctr = struct_set.RoiGeometries[struct.Name].GetCenterOfRoi().x  # R-L center of ROI
        else:  # Rx is to POI
            rx_ctr = struct_set.PoiGeometries[struct.Name].Point.x
    elif hasattr(rx, 'OnDoseSpecificationPoint'):  # Rx is to site
        if rx.OnDoseSpecificationPoint is not None:  # Rx is to DSP
            rx_ctr = rx.OnDoseSpecificationPoint.Coordinates.x
        else:  # Rx is to site that is not a DSP
            dose_dist = plan.TreatmentCourse.TotalDose
            if dose_dist.DoseValues is not None and dose_dist.DoseValues.DoseData is not None:
                rx_ctr = dose_dist.GetCoordinateOfMaxDose().x
    else:
        try:
            ini_laser_iso = next(poi for poi in case.PatientModel.PointsOfInterest if poi.Type == 'InitialLaserIsocenter')
            ini_laser_iso_geom = struct_set.PoiGeometries[ini_laser_iso.Name]
            if ini_laser_iso_geom.Point is not None and abs(ini_laser_iso_geom.Point.x) != float('inf'):
                rx_ctr = ini_laser_iso_geom.Point.x
            else:
                rx_ctr = None
        except StopIteration:
            rx_ctr = None

    ## Apply templates
    for template_name in template_names:
        goals = data[template_name]

        # 'Fine-tune' the goals to apply
        # Check fractionation, Rx, body site, side, etc.

        ## Add goals   
                
        goals['ROI'] = pd.Series(goals['ROI']).fillna(method='ffill')  # Autofill ROI name (due to vertically merged cells in spreadsheet)
    
        invalid_goals_template, empty_spare_template, lg_spare_vol_template, no_ipsi_contra_template, no_nodal_ptv_template = [], [], [], [], []
        roi_regex = r'^{}(_[LR])?(\^.+)?( \(\d+\))?$'
        for _, row in goals.iterrows():  # Iterate over each row in DataFrame
            args = {}  # dict of arguments for ApplyTemplates
            roi = row['ROI']  # e.g., 'Lens'
            rois = []
            for r in case.PatientModel.RegionsOfInterest:
                if re.match(roi_regex.format(roi), r.Name, re.IGNORECASE) or (roi in SPECIFIC_ROIS and any(re.match(roi_regex.format(specific_roi), r.Name, re.IGNORECASE) for specific_roi in SPECIFIC_ROIS[roi])):
                    rois.append(r.Name)
                else:
                    # See if ROI in template is a PRV but doesn't specify a numerical expansion
                    prv_in_roi = re.search('PRV', roi, re.IGNORECASE)
                    if prv_in_roi is None or prv_in_roi.end() != len(roi):
                        continue
                    prv_in_r = re.search('PRV', r.Name, re.IGNORECASE)
                    if prv_in_r is None:
                        continue
                    roi_base = roi[:prv_in_roi.start()].strip('_')
                    r_base = r.Name[:prv_in_r.start()].strip('_')
                    if roi_base.lower() != r_base.lower():
                        continue
                    rois.append(r.Name)
            if not rois:  # ROI in goal does not exist in case
                continue

            goal = re.sub(r'\s', r'', row['Goal'])  # Remove spaces in goal
            invalid_goal = roi + ':\t' + row['Goal']

            # If present, notes may be Rx, body site, body side, or info irrelevant to script
            notes = row['Notes']
            if not pd.isna(notes):  # Notes exist
                notes_lower = notes.lower()
                # Goal only applies to specific Rx
                m = re.match(r'([\d\.]+) gy', notes_lower)
                if m is not None: 
                    notes_lower = int(float(m.group(1)) * 100)  # Extract the number and convert to cGy
                    if rx != notes_lower:  
                        continue

                # Goal only applies to certain fx(s)
                elif notes_lower.endswith('fx'):
                    fxs = [int(elem.strip(',')) for elem in notes_lower[:-3].split(' ') if elem.strip(',').isdigit()]
                    if fx not in fxs:
                        continue
                
                # Ipsilateral objects have same sign on x-coordinate (so product is positive); contralateral have opposite signs (so product is negative)
                elif notes_lower in ['ipsilateral', 'contralateral']:
                    if rx_ctr is None:
                        no_ipsi_contra_template.append('{} ({})'.format(invalid_goal, notes))
                    else:
                        rois = [r for r in rois if (notes_lower == 'ipsilateral' and rx_ctr * struct_set.RoiGeometries[r].GetCenterOfRoi().x > 0) or (notes_lower == 'contralateral' and rx_ctr * struct_set.RoiGeometries[r].GetCenterOfRoi().x < 0)]  # Select the ipsilateral or contralateral matching ROIs
                # Otherwise, irrelevant info

            # Visualization Priority (note that this is NOT the same as planning priority)
            args['Priority'] = template_names.index(template_name) + 1
            
            ## Parse dose and volume amounts from goal. Then add clinical goal for volume or dose.

            # Regexes to match goal
            dose_amt_regex = '''(
                                    (?P<dose_pct_rx>[\d.]+%)?
                                    (?P<dose_rx>Rx[pn]?)|
                                    (?P<dose_amt>[\d.]+)
                                    (?P<dose_unit>c?Gy)
                                )'''  # e.g., 95%Rx or 20Gy
            dose_types_regex = '(?P<dose_type>max|min|mean|median)'
            vol_amt_regex = '''(
                                    (?P<vol_amt>[\d.]+)
                                    (?P<vol_unit>%|cc)|
                                    (\(v-(?P<spare_amt>[\d.]+)\)cc)
                            )'''  # e.g., 67%, 0.03cc, or v-700cc
            sign_regex = '(?P<sign><|>)'  # > or <
            dose_regex = 'D(' + dose_types_regex + '|' + vol_amt_regex + ')' + sign_regex + dose_amt_regex  # e.g., D0.03cc<110%Rx, Dmedian<20Gy
            vol_regex = 'V' + dose_amt_regex + sign_regex + vol_amt_regex  # e.g., V20Gy<67%

            # Need separate regexes b/c we can't have duplicate group names in a single regex
            # Remove whitespace from regex (left in above for readability) before matching
            vol_match = re.match(re.sub(r'\s', r'', vol_regex), goal)
            dose_match = re.match(re.sub(r'\s', r'', dose_regex), goal)
            m = vol_match if vol_match is not None else dose_match  # If it's not a volume, should be a dose

            if not m:  # Invalid goal format -> add goal to invalid goals list and move on to next goal
                invalid_goals_template.append(invalid_goal)
                continue

            args['GoalCriteria'] = 'AtMost' if m.group('sign') == '<' else 'AtLeast'  # GoalCriteria depends on sign

            # Extract dose: an absolute amount or a % of Rx
            dose_rx = m.group('dose_rx')
            if dose_rx:  # % of Rx
                if rx_val is None:
                    continue
                dose_pct_rx = m.group('dose_pct_rx')  # % of Rx
                if dose_pct_rx is None:  # Group not present. % of Rx is just specified as 'Rx'
                    dose_pct_rx = 100
                else:  # A % is specified, so make sure format is valid
                    try:
                        dose_pct_rx = float(dose_pct_rx[:-1])  # Remove percent sign and convert to float
                        if dose_pct_rx < 0:  # % out of range -> add goal to invalid goals list and move on to next goal
                            invalid_goals_template.append(invalid_goal)
                            continue
                    except:  # % is non numeric -> add goal to invalid goals list and move on to next goal
                        invalid_goals_template.append(invalid_goal)
                        continue
                # Find appropriate Rx (to primary or nodal PTV)
                if dose_rx == 'Rxn':  # Use 2ry Rx (to nodal PTV)
                    rx_n = [rx_n for rx_n in beam_set.Prescription.PrescriptionsDoseReferences if 'PTVn' in rx_n.OnStructure.Name]
                    if rx_n:  # Found a nodal PTV (should never be more than one)
                        dose_rx = rx_n[0]
                    else:  # There is no nodal PTV, so add goal to list of goals that could not be added to nodal PTV, and move on to next goal
                        no_nodal_ptv_template.append(invalid_goal)
                else:  # Primary Rx
                    dose_rx = rx
                dose_amt = dose_pct_rx / 100 * rx_val  # Get absolute dose based on % Rx
            else:  # Absolute dose
                try:
                    dose_amt = float(m.group('dose_amt'))  # Account for scaling to template Rx if user selected this option (remember that `scaling_factor` is 1 otherwise)
                except:  # Given dose amount is non-numeric  -> add goal to invalid goals list and move on to next goal
                    invalid_goals_template.append(invalid_goal)
                    continue
                if m.group('dose_unit') == 'Gy':  # Covert dose from Gy to cGy
                    dose_amt *= 100
            if dose_amt < 0 or dose_amt > 100000:  # Dose amount out of range  -> add goal to invalid goals list and move on to next goal
                invalid_goals_template.append(invalid_goal)
                continue

            # Extract volume: an absolute amount, a % of ROI volume, or an absolute amount to spare
            dose_type = vol_unit = spare_amt = None
            vol_amt = m.group('vol_amt')
            if vol_amt:  # Absolute volume or % of ROI volume
                try:
                    vol_amt = float(vol_amt)
                except:  # Given volume is non numeric -> add goal to invalid goals list and move on to next goal
                    invalid_goals_template.append(invalid_goal)
                    continue

                vol_unit = m.group('vol_unit')
                if vol_unit == '%':  # If relative volume, adjust volume amount
                    if vol_amt > 100:  # Given volume is out of range -> add goal to invalid goals list and move on to next goal
                        invalid_goals_template.append(invalid_goal)
                        continue
                    vol_amt /= 100  # Convert percent to proportion

                if vol_amt < 0 or vol_amt > 100000:  # Volume amount out of range supported by RS -> add goal to invalid goals list and move on to next goal
                    invalid_goals_template.append(invalid_goal)
                    continue
            else:  # Volume to spare or dose type
                spare_amt = m.group('spare_amt')
                if spare_amt:  # Volume to spare
                    try:
                        spare_amt = float(spare_amt)
                    except:  # Volume amount is non numeric -> add goal to invalid goals list and move on to next goal
                        invalid_goals_template.append(invalid_goal)
                        continue
                    
                    geom = struct_set.RoiGeometries[roi]
                    if not geom.HasContours():  # Cannot add volume to spare goal for empty geometry -> add goal to list of vol-to-spare goals for empty geometries
                        empty_spare_template.append(invalid_goal)
                        continue
                    if spare_amt < 0:  # Negative spare amount -> add goal to invalid goals list and move on to next goal
                        invalid_goals_template.append(invalid_goal)
                        continue
                    if spare_amt > geom.GetRoiVolume():
                        lg_spare_vol_template.append(invalid_goal)
                        continue
                else:  # Dose type: Dmax, Dmean, or Dmedian
                    dose_type = m.group('dose_type')

            # D...
            if goal.startswith('D'):
                # Dmax = D0.035
                if dose_type == 'max':
                    args['GoalType'] = 'DoseAtAbsoluteVolume'
                    args['ParameterValue'] = 0.03
                # Dmin = Max volume at that dose is everything but 0.035 cc
                elif dose_type == 'min':
                    args['GoalType'] = 'AbsoluteVolumeAtDose'
                    args['ParameterValue'] = dose_amt
                    args['AcceptanceLevel'] = struct_set.RoiGeometries[roi].GetRoiVolume() - 0.035
                # Dmean => 'AverageDose'
                elif dose_type == 'mean':
                    args['GoalType'] = 'AverageDose'
                # Dmedian = D50%
                elif dose_type == 'median':
                    args['GoalType'] = 'DoseAtVolume'
                    args['ParameterValue'] = 0.5
                # Absolute or relative dose
                else:
                    args['ParameterValue'] = vol_amt
                    if vol_unit == '%':
                        args['GoalType'] = 'DoseAtVolume'
                    else:
                        args['GoalType'] = 'DoseAtAbsoluteVolume'
                args['AcceptanceLevel'] = dose_amt
            # V...
            else:
                args['ParameterValue'] = dose_amt
                if vol_unit == '%':
                    args['GoalType'] = 'VolumeAtDose'
                else:
                    args['GoalType'] = 'AbsoluteVolumeAtDose'
                if not spare_amt:
                    args['AcceptanceLevel'] = vol_amt
                
            # Add Clinical Goals
            for roi in rois:
                roi_args = args.copy()
                roi_args['RoiName'] = roi
                if spare_amt:
                    if not struct_set.RoiGeometries[roi].HasContours():
                        continue
                    total_vol = struct_set.RoiGeometries[roi].GetRoiVolume()
                    roi_args['AcceptanceLevel'] = total_vol - spare_amt
                try:
                    plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(**roi_args)
                except:
                    pass

        if invalid_goals_template:
            invalid_goals[template_name] = invalid_goals_template
        if empty_spare_template:
            empty_spare[template_name] = empty_spare_template
        if lg_spare_vol_template:
            lg_spare_vol[template_name] = lg_spare_vol_template
        if no_ipsi_contra_template:
            no_ipsi_contra[template_name] = no_ipsi_contra_template
        if no_nodal_ptv_template:
            no_nodal_ptv[template_name] = no_nodal_ptv_template

    # Add warnings about clinical goals that were not added
    if invalid_goals:
        warnings += 'The following clinical goals could not be parsed so were not added:'
        warnings += format_warnings(invalid_goals)
    if empty_spare:
        warnings += 'The following clinical goals could not be added due to empty geometries:'
        warnings += format_warnings(empty_spare)
    if lg_spare_vol:
        warnings += 'The following clinical goals could not be added because the volume to spare is larger than the ROI volume:'
        warnings += format_warnings(lg_spare_vol)
    if no_ipsi_contra:
        warnings += 'There is no Rx, so ipsilateral/contralateral structures could not be determined. Ipsilateral/contralateral clinical goals were not added:'
        warnings += format_warnings(no_ipsi_contra)
    if no_nodal_ptv:
        warnings += 'No nodal PTV was found, so the following clinical goals were not added:'
        warnings += format_warnings(no_nodal_ptv)

    # Add template names to plan comments
    new_comments = 'Clinical Goals template(s) were applied:\n' + '\n'.join(f'{template_names.index(name) + 1}. {name}' for name in template_names)
    if plan.Comments == '':
        plan.Comments = new_comments
    else:
        plan.Comments = plan.Comments + '\n' + new_comments

    # Display warnings if there were any
    if warnings != '':
        MessageBox.Show(warnings, 'Warnings')
    