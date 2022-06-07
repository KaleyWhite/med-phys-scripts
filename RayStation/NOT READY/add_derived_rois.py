import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
import sys
from random import randint
from re import search

import pandas as pd
from connect import *
from System.Drawing import *
from System.Windows.Forms import *


case = exam = None
to_update = []  # List of ROI names whose derived geometries need updating
targets = {target_type: None for target_type in ["CTV", "GTV", "PTV"]}  # The target of each type that will be used in derived ROI expressions. Set by user when needed


class ChooseTargetForm(Form):
    # Helper form that allows the user to choose a target from a provided list
    # Used when there are multiple targets of the given type in the current case

    def __init__(self, target_type, target_rois):
        self.target_type = target_type

        # Set up Form
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow# User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = "Choose Target"
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        y = 15  # Vertical coordinate of next control

        l = Label()
        l.AutoSize = True
        l.Location = Point(15, y)
        l.Text = "There are multiple {}s in the current case.\nChoose the one to use in derived ROI expressions.".format(target_type.upper())
        self.Controls.Add(l)
        y += l.Height

        self.gb = GroupBox()
        self.gb.AutoSize = True
        self.gb.Location = Point(15, y)
        rb_y = 20
        for target_roi in target_rois:
            rb = RadioButton()
            rb.AutoSize = True
            rb.Checked = False
            rb.Click += self.set_ok_enabled
            rb.Location = Point(15, rb_y)
            rb.Text = target_roi.Name
            self.gb.Controls.Add(rb)
            rb_y += rb.Height
        self.Controls.Add(self.gb)
        y += self.gb.Height + 15

        self.ok = Button()
        self.ok.Click += self.ok_clicked
        self.ok.Enabled = False
        self.ok.Location = Point(15, y)
        self.ok.Text = "OK"
        self.Controls.Add(self.ok)

        self.ShowDialog()

    def set_ok_enabled(self, sender, event):
        self.ok.Enabled = any(rb.Checked for rb in self.gb.Controls)

    def ok_clicked(self, sender, event):
        global targets

        targets[self.target_type.upper()] = [case.PatientModel.RegionsOfInterest[rb.Text] for rb in self.gb.Controls if rb.Checked][0]
        self.DialogResult = DialogResult.OK


def set_target(target_type):
    # Helper function that updates the `targets` dictionary for the given target type

    if targets[target_type.upper()] is None:  # We have not yet set the target of that type
        target_rois = [roi for roi in case.PatientModel.RegionsOfInterest if roi.Type.lower() == target_type.lower()]  # All ROIs of that type
        if not target_rois:  # There are not any ROIs of that type, so leave `None` in the dictionary and return
            return
        if len(target_rois) == 1:  # Exactly one ROI of that type
            targets[target_type.upper()] = target_rois[0]
        else:  # Multiple ROIs of that type, so user chooses the one they want to use in derived ROI expressions
            form = ChooseTargetForm(target_type, target_rois)
            if form.DialogResult != DialogResult.OK:  # "OK" was never clicked (maybe user clicked the "X" button)
                sys.exit()


def name_item(item, l, max_len=sys.maxsize):
    # Helper function that generates a unique name for `item` in list `l` (case insensitive)
    # Limit name to `max_len` characters
    # E.g., name_item("Isocenter Name A", ["Isocenter Name A", "Isocenter Na (1)", "Isocenter N (10)"]) -> "Isocenter Na (2)"

    l_lower = [l_item.lower() for l_item in l]
    copy_num = 0
    old_item = item
    while item.lower() in l_lower:
        copy_num += 1
        copy_num_str = " ({})".format(copy_num)
        item = "{}{}".format(old_item[:(max_len - len(copy_num_str))].strip(), copy_num_str)
    return item[:max_len]



def create_roi_if_absent(roi_name, roi_type="Organ"):
    # Helper function that returns the "latest" UNAPPROVED ROI with the given name, as determined by the copy number in the name
    # If no such ROI is found, return a new ROI with the given name and type
    # `roi_name`: "Base" ROI name. Assume TG-263 compliance.
    # `roi_type`: If a new ROI is created, it is of this type.

    # CRMC standard ROI colors
    colors_filename = r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Data\TG263 Nomenclature with CRMC Colors.csv"
    colors = pd.read_csv(colors_filename)

    roi = get_latest_roi(roi_name, unapproved_only=True)  # Unapproved ROI with this base name and the latest copy number
    if roi is None:  # No unapproved ROIs with that base name exist
        if roi_name in colors["TG263-Primary Name"]:  # The TG-263 name has been catalogued in the spreadsheet
            color = colors.loc[colors["TG263-Primary Name"] == roi_name, "Color"]
        else:  # TG-263 name not yet catalogued, so add it
            new_row = {"N Characters": len(roi_name), "TG263-Primary Name": roi_name, "Named by CRMC": "Y"}
            # Get unique color
            while True:
                color = "255; {}; {}; {}".format(randint(0, 255), randint(0, 255), randint(0, 255))
                if color not in colors["Color"]:
                    new_row["Color"] = color
                    break
            new_row.update({col: "" for col in ["Target Type", "Major Category", "Minor Category", "Anatomic Group", "TG-263-Reverse Order Name", "Description", "FMAID", "Possible Alternate Names"]})
            colors = colors.append(new_row, ignore_index=True)  # Add new row to DataFrame
            colors.sort_values(by="TG263-Primary Name", inplace=True)  # Sort by TG-263 name
            colors.to_csv(colors_filename, index=False)  # Overwrite colors spreadsheet

        # Create new ROI
        color = color.replace(";", ",")  # E.g., "255; 1; 2; 3" -> "255, 1, 2, 3"
        roi_name = name_item(roi_name, [r.Name for r in case.PatientModel.RegionsOfInterest], 16)
        roi = case.PatientModel.CreateRoi(Name=roi_name, Type=roi_type, Color=color)
    
    return roi


def get_latest_roi(base_roi_name, **kwargs):
    # Helper function that returns the ROI with the given "base name" and largest copy number
    # kwargs:
    # unapproved_only: If True, consider only the ROIs that are not part of any approved structure set in the case
    # non_empty_only: If True, consider only the ROIs with geometries on the exam

    unapproved_only = kwargs.get("unapproved_only", False)
    non_empty_only = kwargs.get("non_empty_only", False)

    base_roi_name = base_roi_name.lower()

    rois = case.PatientModel.RegionsOfInterest
    if unapproved_only:
        approved_roi_names = set(geom.OfRoi.Name for ss in case.PatientModel.StructureSets for approved_ss in ss.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures)
        rois = [roi for roi in rois if roi.Name not in approved_roi_names]
    if non_empty_only:
        rois = [roi for roi in rois if case.PatientModel.StructureSets[exam.Name].RoiGeometries[roi.Name].HasContours()]
    
    latest_copy_num = copy_num = -1
    latest_roi = None
    for roi in rois:
        roi_name = roi.Name.lower()
        if roi_name == base_roi_name:
            copy_num = 0
        else:
            m = search(" \((\d+)\)".format(base_roi_name), roi_name)
            if m:  # There is a copy number
                grp = m.group()
                length = min(16 - len(grp), len(base_roi_name))
                if roi_name[:length] == base_roi_name[:length]:
                    copy_num = int(m.group(1))
        if copy_num > latest_copy_num:
            latest_copy_num = copy_num
            latest_roi = roi
    return latest_roi


def organ_minus_target(target_type, organ_names):
    # Helper function that sets an ROI algebra expression for organs minus the given target type
    # `target_type`: "CTV", "GTV", or "PTV"
    # `organ_names`: List of names of ROIs from which to subtract the target

    global to_update

    # Get the target ROI
    set_target(target_type)
    target = targets[target_type]
    if target is None:
        return
    
    # Create each derived ROI, including left and right sides
    for organ_name in organ_names:
        for ext in ["", "_L", "_R"]:  # E.g., "Chestwall" -> "Chestwall", "Chestwall_L", "Chestwall_R"
            temp_name = "{}{}".format(organ_name, ext)
            organ = get_latest_roi(temp_name)
            if organ is None:  # Organ does not exist in current case, so ignore it
                continue
            derived = create_roi_if_absent("{}-{}".format(temp_name, target_type.upper()))  # E.g., "Chestwall_L-PTV"
            derived.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [organ.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [target.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Subtraction", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
            to_update.append(derived.Name)


def union_or_intersection(derived_name, source_names, union=True):
    # Helper function that sets an ROI algebra expression for the union or intersection of the given ROIs
    # `derived_name`: Name of the union or intersection ROI
    #                 Necessary to specify because 16-character limit may cause name to differ fro standard.
    # `source_names`: List of names of ROIs to be unioned or intersected
    # `union`: True if the source ROIs should be unioned, False if they should be intersected

    global to_update

    latest_source_names = []
    for source_name in source_names:
        if source_name in targets:  # We don't regulate target names by TG-263, so rely on the set target of the given type
            set_target(source_name)
            latest_source = targets[source_name]
        else:
            latest_source = get_latest_roi(source_name)
        if latest_source is None:  # Source ROI absent in current case, so do not create this derived ROI
            return
        latest_source_names.append(latest_source.Name)
    derived = create_roi_if_absent(derived_name)
    derived.SetAlgebraExpression(ExpressionA={ 'Operation': "Union" if union else "Intersection", 'SourceRoiNames': latest_source_names, 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } })
    to_update.append(derived.Name)


def planning_region_volume(exp, source_name, derived_name=None):
    # Helper function that sets a margin expression for the given organ
    # `exp`: Uniform expansion amount, in mm
    # `source_name`: Name of the ROI to expand
    # `derived_name`: Name of the PRV ROI. If None, new ROI name is source name + "PRV" + two-digit expansion. 
    #                 Necessary to specify because 16-character limit may cause name to differ fro standard.

    global to_update

    if derived_name is None:
        derived_name = "{}_PRV{}".format(source_name, str(exp).zfill(2))  # E.g., "OpticNrv_PRV03"
    exp /= 10  # mm -> cm

    for ext in ["", "_L", "_R"]:
        temp_name = "{}{}".format(source_name, ext)
        source = get_latest_roi(temp_name)
        if source is None:  # Source ROI does not exist in current case, so ignore it
            continue
        if ext != "":  # This is an L or R ROI
            if len(derived_name) == 15:  # Only room for 1 more charcter in the name, so omit the underscore (e.g., "LbeTemporalPRV1L" instead of "LbeTemporalPRV1_L")
                derived_name = "{}{}".format(derived_name, ext[-1])
            else:  # There is room for the underscore
                derived_name = "{}{}".format(derived_name, ext)
        derived = create_roi_if_absent(derived_name, "Control")  # PRVs are controls, not OARs
        derived.SetMarginExpression(SourceRoiName=source.Name, MarginSettings={ 'Type': "Expand", 'Superior': exp, 'Inferior': exp, 'Anterior': exp, 'Posterior': exp, 'Right': exp, 'Left': exp })
        to_update.append(derived.Name)


def get_tx_technique(bs):
    # Helper function that returns the treatment technique for the given beam set.
    # SMLC, VMAT, DMLC, 3D-CRT, conformal arc, or applicator and cutout
    # Return "?" if treatment technique cannot be determined
    # Code modified from RS support

    if bs.Modality == "Photons":
        if bs.PlanGenerationTechnique == "Imrt":
            if bs.DeliveryTechnique == "SMLC":
                return "SMLC"
            if bs.DeliveryTechnique == "DynamicArc":
                fx_pattern = bs.FractionationPattern
                if fx_pattern is not None:
                    num_fx = fx_pattern.NumberOfFractions
                    if num_fx in [3, 5]:
                        rx = bs.Prescription.PrimaryPrescription
                        if rx is not None and rx.DoseValue / num_fx >= 6:
                            if num_fx == 3:  # The only 3 Fx VMAT that CRMC currently does is SRS brain
                                return "SRS" 
                            return "SBRT"  # The only 5 Fx VMAT that CRMC currently does is SRS brain
                return "VMAT"
            if bs.DeliveryTechnique == "DMLC":
                return "DMLC"
        elif bs.PlanGenerationTechnique == "Conformal":
            if bs.DeliveryTechnique == "SMLC":
                # return "SMLC" # Changed from "Conformal". Failing with forward plans.
                return "Conformal"
                # return "3D-CRT"
            if bs.DeliveryTechnique == "Arc":
                return "Conformal Arc"
    elif bs.Modality == "Electrons":
        if bs.PlanGenerationTechnique == "Conformal":
            if bs.DeliveryTechnique == "SMLC":
                return "ApplicatorAndCutout"
    return "?"


def add_derived_rois():
    """Add derived ROI geometries from the "Clinical Goals" spreadsheet

    Types of derived geometries created:
    - Bilateral organ sums (e.g., "Lungs" = "Lung_L" + "Lung_R")
    - Other, misc. sums (e.g., "Jejunum_Ileum" = "Jejunum" + "Ileum")
    - Overlaps (intersections) (e.g., "BneMdbl&JntTM&PTV" = "Bone_Mandible" & "Joint_TM" & PTV)
    - Target exclusions (subtracting a target volume from an organ) (e.g., "Brain-PTV" = "Brain" - PTV)
    - Planning regions volumes (PRVs) (e.g., "Brainstem_PRV03" = 3mm expansion of "Brainstem")
    - "E-PTV_Ev20" (SBRT plans only)

    If source ROI(s) do not exist in the current case, derived ROI(s) that depend on them are not created.
    
    Source ROIs are the ROIs with the "latest" name.
    Example:
    We are creating "Lungs" from "Lung_L" and "Lung_R". ROIs "Lung_L" and "Lung_L (1)" exist. "Lung_L (1)" has the highest copy number, so it is used.

    Derived geometries are created for the latest unapproved ROIs with the desired derived name. If none such ROIs exist, a new ROI is created.
    Example:
    We are creating "Lungs". ROIs "Lungs" and "Lungs (1)" exist, but "Lungs (1)" is approved, so "Lungs" is used. If both "Lungs" and "Lungs (1)" were approved, we would create a new ROI "Lungs (2)".

    Assumptions
    -----------
    All ROI names (perhaps excluding targets) are TG-263 compliant
    """

    global case, exam, to_update

    # Get current objects
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)
    try:
        exam = get_current("Examination")
    except:
        MessageBox.Show("There are no examinations in the current case. Click OK to abort script.", "No Examinations")
        sys.exit(1)

    to_update = []

    # Sums
    l_r_sums = ["Kidney", "Lung", "Parotid"]
    for l_r_sum in l_r_sums:
        union_or_intersection("{}s".format(l_r_sum), ["{}_L".format(l_r_sum), "{}_R".format(l_r_sum)])

    other_sums = {
        "BneMndble_JointTM": ["Bone_Mandible", "Joint_TM"],
        "Jejunum_Ileum": ["Jejunum", "Ileum"]
    }
    for sum_roi_name, base_roi_names in other_sums.items():
        union_or_intersection(sum_roi_name, base_roi_names)

    # Overlaps
    overlaps = {
        "BnMdbl&JntTM&PTV": ["Bone_Mandible", "Joint_TM", "PTV"]
    }
    for overlap_name, base_roi_names in overlaps.items():
        union_or_intersection(overlap_name, base_roi_names, union=False)

    # OAR expansions
    prvs = [
        (1, "Brainstem"),
        (1, "Lobe_Temporal", "LbeTemporalPRV1"),
        (1, "OpticNrv"),
        (1, "Tongue_All"),
        (3, "Brainstem"),
        (5, "CaudaEquina"),
        (3, "OpticChiasm", "OpticChiasm_PRV3"),
        (5, "Glnd_LcrimalPRV5"),
        (5, "Skin"),
        (5, "SpinalCord"),
        (7, "Lens")
    ]
    for prv in prvs:
        if len(prv) == 2:
            planning_region_volume(prv[0], prv[1])
        else:
            planning_region_volume(prv[0], prv[1], prv[2]) 

    # Target exclusion
    target_exclude = {
        "CTV": ["Lungs"],
        "GTV": ["Liver", "Lungs"],
        "PTV": ["Brain", "Cavity_Oral", "Chestwall", "Glottis", "Larynx", "Liver", "Lungs", "Stomach"]
    }
    for target, base_names in target_exclude.items():
        organ_minus_target(target, base_names)

    # E-PTV_Ev20
    try:
        plan = get_current("Plan")
        if plan.BeamSets.Count > 0:
            if any(get_tx_technique(bs) == "SBRT" for bs in plan.BeamSets):
                ext = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"]
                if ext:
                    set_target("PTV")
                    ptv = targets["PTV"]
                    if ptv is not None:
                        nt = create_roi_if_absent("E-PTV_Ev20", "Control")
                        nt.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [ext[0].Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [ptv.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0.2, 'Inferior': 0.2, 'Anterior': 0.2, 'Posterior': 0.2, 'Right': 0.2, 'Left': 0.2 } }, ResultOperation="Subtraction", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                        to_update.append(nt.Name)
    except:
        pass

    # Update derived geometries if any were added
    if to_update:
        case.PatientModel.UpdateDerivedGeometries(RoiNames=to_update, Examination=exam)
