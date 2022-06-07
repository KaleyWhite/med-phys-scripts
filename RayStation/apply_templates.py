# For GUI
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import os
import random
import re
import sys
sys.path.append(os.path.join("T:", "Physics - T", "Scripts", "RayStation"))

import pandas as pd  # Clinical Goals template data is read in as DataFrame
from connect import *  # Interact w/ RS

# For GUI
from System.Drawing import *
from System.Windows.Forms import *

from CenterCouchScript import *


# Get current variables
try:
    case = get_current("Case")
except:
    MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
    sys.exit(1)
try:
    plan = get_current("Plan")
except:
    MessageBox.Show("There is no plan loaded. Click OK to abort script.", "No Plan Loaded")
    sys.exit(1)
patient_db = get_current("PatientDB")
struct_set = plan.GetStructureSet()  # Geometries on the planning exam
exam = struct_set.OnExamination  # Planning exam

# Choices for the GUI to display
mds = ["Jiang", "Other"]  # All MDs that we have templates for
plan_types = ["3D", "IMRT", "SBRT", "SRS", "VMAT"]
body_sites = ["Brain", "Breast", "Esophagus", "GI", "Gyn", "H&N", "Lung", "Pancreas", "PB", "Pelvis", "Pros", "Other"]
templates = ["Clinical Goals", "Colorwash", "Structure"]  # Types of templates that the user may apply


def apply_struct_template(template_name, init_option=None):
    # Helper function that applies structure template named `template_name` using initialization option `init_option` (e.g., "EmptyGeometries" or "AlignImageCenters")
    # If `init_option` is None, use atlas-based initialization
    
    template = patient_db.LoadTemplatePatientModel(templateName=template_name)  # Must load template before it can be used
    approved_geom_names = [geom.OfRoi.Name for approved_ss in struct_set.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures]  # All ROI names approved on the planning exam
    src_roi_names = [roi.Name for roi in template.PatientModel.RegionsOfInterest if roi.Name not in approved_geom_names]  # All unapproved, underived ROI names in the template
    src_poi_names = [poi.Name for poi in template.PatientModel.PointsOfInterest]  # All POI names in the template
    
    if init_option is not None:  # Don't use atlas-based initialization
        src_exam_name = template.PatientModel.StructureSets[0].OnExamination.Name  # Doesn't matter which source exam we use, so just use the first one
        case.PatientModel.CreateStructuresFromTemplate(SourceTemplate=template, SourceExaminationName=src_exam_name, SourceRoiNames=src_roi_names, SourcePoiNames=src_poi_names, TargetExamination=exam, InitializationOption=init_option)
    else:
        # All exam names in template
        src_exam_names = [exam_.Name for exam_ in template.StructureSetExaminations]
        
        # Atlas-based initialization
        # General consensus is 15 fusions: any more doesn't improve accuracy, just increases time
        # Try to apply the template with all included ROIs. If some are not in the image, exclude them and try again.
        try:
            case.PatientModel.CreateStructuresFromAtlas(SourceTemplate=template, SourceExaminationsNames=src_exam_names, SourceRoiNames=src_roi_names, SourcePoiNames=src_poi_names, TargetExamination=exam, NrOfFusionAtlases=15)
        except Exception as e:
            roi_names_exclude = str(e).split("re-run: ")[1].split()  # Extract ROI names from error message, which looks something like, "The following ROI(s) are not sufficiently included in image. Unselect them for segmentation and re-run: Bladder Rectum"
            src_roi_names = [roi_name for roi_name in src_roi_names if roi_name not in roi_names_exclude]  # Remove those ROI names from source ROI names list
            case.PatientModel.CreateStructuresFromAtlas(SourceTemplate=template, SourceExaminationsNames=src_exam_names, SourceRoiNames=src_roi_names, SourcePoiNames=src_poi_names, TargetExamination=exam, NrOfFusionAtlases=15)

    # Update derived geometries that were either added by the template or empty before template was applied
    derived_roi_names = [roi.Name for roi in template.PatientModel.RegionsOfInterest if roi.Name in src_roi_names and roi.DerivedRoiExpression is not None]
    case.PatientModel.UpdateDerivedGeometries(RoiNames=derived_roi_names, Examination=exam, AreEmptyDependenciesAllowed=True)


def format_list(l):
    # Helper function that returns a nicely formatted string of elements in a list
    # E.g., format_list(["A", "B", None]) -> "A, B, and None"

    if len(l) == 1:
        return l[0]
    if len(l) == 2:
        return "{} and {}".format(l[0], l[1])
    l_ = ["'{}'".format(item) if isinstance(item, str) else item for item in l]
    return "{}, and {}".format(", ".join(l_[:-1]), l_[-1])


class ApplyTemplatesForm(Form):
    # Form that allows user to select MD, treatment technique, and body site from a GUI, to be used as template selection criteria
    # User also selects the types of template to apply (e.g., clinical goals)

    def __init__(self, selected_md, selected_plan_type, selected_body_site):
        # Parameters are the default checked item for each radio button group

        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ("X out of") it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen
        self.Text = "Apply Templates"  # Form title
        y = 15  # Vertical coordinate of next control
        
        ## MD

        # MD GroupBox
        self.md_gb = GroupBox()
        self.md_gb.AutoSize = True
        self.md_gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.md_gb.Text = "MD:"
        self.md_gb.Location = Point(15, y)
        self.Controls.Add(self.md_gb)
        
        # MD RadioButtons
        rb_y = 15  # Vertical coordinate of next radio button
        for md in mds:
            rb = RadioButton()
            rb.Click += self.set_ok_enabled
            rb.Text = md
            rb.Checked = md == selected_md  # Select default MD
            rb.Location = Point(15, rb_y)
            self.md_gb.Controls.Add(rb)
            rb_y += 20
        self.md_gb.Height = len(mds) * 20  # Leave 20 units for each radio button (autosized groupbox height tends to be off a little)
        y += self.md_gb.Height + 15  # Leave 15 units between groupbox and next groupbox

        ## Plan types

        # Plan type GroupBox
        self.plan_type_gb = GroupBox()
        self.plan_type_gb.AutoSize = True
        self.plan_type_gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.plan_type_gb.Text = "Plan type:"
        self.plan_type_gb.Location = Point(15, y)
        self.Controls.Add(self.plan_type_gb)

        # Plan type RadioButtons
        rb_y = 15
        for plan_type in plan_types:
            rb = RadioButton()
            rb.Click += self.set_ok_enabled
            rb.Text = plan_type
            rb.Checked = plan_type == selected_plan_type  # Select default plan type
            rb.Location = Point(15, rb_y)
            self.plan_type_gb.Controls.Add(rb)
            rb_y += 20
        self.plan_type_gb.Height = len(plan_types) * 20
        y += self.plan_type_gb.Height + 15

        ## Body sites

        # Body site GroupBox
        self.body_site_gb = GroupBox()
        self.body_site_gb.AutoSize = True
        self.body_site_gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.body_site_gb.Text = "Body site:"
        self.body_site_gb.Location = Point(15, y)
        self.Controls.Add(self.body_site_gb)

        # Body site RadioButtons
        rb_y = 15
        for body_site in body_sites:
            rb = RadioButton()
            rb.Click += self.set_ok_enabled
            rb.Text = re.sub("&", "&&", body_site)  # Escape "&" as "&&" so that it is displayed (e.g., "H&N")
            rb.Checked = body_site == selected_body_site  # Select default body site
            rb.Location = Point(15, rb_y)
            self.body_site_gb.Controls.Add(rb)
            rb_y += 20
        self.body_site_gb.Height = len(body_sites) * 20
        y += self.body_site_gb.Height + 15

        ## Template types

        # Templates ListBox
        # Listbox is easier than CheckBox b/c multiple selection, including select all, is "built in"
        gb = GroupBox()
        gb.AutoSize = True
        gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        gb.Text = "Template(s) to apply:"
        gb.Location = Point(15, y)
        self.Controls.Add(gb)

        # Templates ListBox
        self.templates_lb = ListBox()
        self.templates_lb.Location = Point(15, 15)
        self.templates_lb.SelectionMode = SelectionMode.MultiExtended  # User may select multiple items
        # Add and select all template types
        for i, template in enumerate(templates):
            self.templates_lb.Items.Add(template)
            self.templates_lb.SetSelected(i, True)
        self.templates_lb.Height = self.templates_lb.PreferredHeight  # "Autosize" based on number of items
        gb.Controls.Add(self.templates_lb)
        y += gb.Height + 15

        # Set uniform groupbox width
        width = max(gb.Width for gb in self.Controls)  # All controls so far are groupboxes
        for gb in self.Controls:
            gb.MinimumSize = Size(width, 0)  # For autosized controls, set MinimumSize instead of Size

        # OK button
        self.ok = Button()
        self.ok.Click += self.ok_clicked
        self.ok.Text = "OK"
        self.ok.Location = Point(15, y)
        self.AcceptButton = self.ok
        self.set_ok_enabled()
        self.Controls.Add(self.ok)

        self.ShowDialog()  # Launch window

    def set_ok_enabled(self, sender=None, event=None):
        # Enable or disable "OK" button
        # Enable only if at least one item is selected in each GroupBox
        # Is called in __init__ not as event handler, so `sender` and `event` need default arguments

        self.ok.Enabled = all(any(rb.Checked for rb in gb.Controls) for gb in [self.md_gb, self.plan_type_gb, self.body_site_gb]) and self.templates_lb.SelectedItems.Count > 0

    def ok_clicked(self, sender, event):
        # Event handler for "OK" button click

        self.DialogResult = DialogResult.OK


class ChooseTemplatesForm(Form):
    # Form that allows user to choose the template name they want to use
    # Instantiated in main function if no or multiple template names match user's selections
    
    def __init__(self, template_names, template_type, template_names_match):
        # `template_names`: names to display in ListBox
        # `template_type`: type of template (e.g., "clinical goals"), to make prompt more specfic
        # `template_names_match`: True if multiple templates match, False if no templates match

        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.StartPosition = FormStartPosition.CenterScreen  # Position Form in middle of screen
        self.Text = "Multiple Matching Templates" if template_names_match else "No Matching Templates"
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ("X out of") it
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)  # At least as wide as title plus some (e.g., for "X" button)
        y = 15  # Vertical coordinate of next control

        # Label w/ prompt
        lbl = Label()
        lbl.AutoSize = True  # Grow label if necessary to accommodate text
        lbl.Text = "Multiple {} templates match your selections.\nChoose the template(s) you want to use.".format(template_type) if template_names_match else "No {} template matches your selections.\nChoose the template(s) you want to use.".format(template_type)
        lbl.Location = Point(15, y)
        self.Controls.Add(lbl)
        y += lbl.Height + 10

        # ListBox from which to choose template(s)
        self.lb = ListBox()
        self.lb.SelectionMode = SelectionMode.MultiExtended  # User may select multiple items
        self.lb.Items.AddRange(template_names)  # Add all template names but don't select any
        #self.lb.SelectedIndexChanged += self.set_ok_enabled
        self.lb.Location = Point(15, y)
        self.lb.Height = self.lb.PreferredHeight  # "Autosize" ListBox based on number of items
        self.lb.Width = max(TextRenderer.MeasureText(item, self.lb.Font).Width for item in self.lb.Items)  # "Autosize" listbox width to widest item
        self.Controls.Add(self.lb)
        y += self.lb.Height + 15

        # "OK" button
        self.ok = Button()
        self.ok.Location = Point(15, y)
        self.ok.Click += self.ok_clicked
        self.ok.Text = "OK"
        #self.ok.Enabled = False  # Since no templates are selected by default, OK button starts out disabled
        self.AcceptButton = self.ok  # Pressing the Enter key is the same as clicking the OK button
        self.Controls.Add(self.ok)

        self.ShowDialog()  # Launch window

    def ok_clicked(self, sender, event):
        # Even handler for OK button click

        self.DialogResult = DialogResult.OK


class TemplateRxForm(Form):
    # Form that allows user to specify what to do if plan Rx doesn't match Rx in template name
    # User chooses whether or not to scale the goal (default is no)
    # If there are multiple template Rx's, user chooses which ones to use

    def __init__(self, template_name, plan_rx, template_rxs):
        # `template_name`: Name of the template (necessary only for prompt)
        # `plan_rx`: Rx specified in the plan
        # `template_rxs`: Rx's specified in template

        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.StartPosition = FormStartPosition.CenterScreen  # Position Form in middle of screen
        self.Text = "Rx Mismatch"
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ("X out of") it
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, self.Font).Width + 100, 0)  # Form is at least as wide as title, plus some
        self.y = 15  # Vertical coordinate of next control

        # Labels explaining scaling to plan Rx
        if len(template_rxs) == 1:  # 1 Rx specified in template
            self.label("The prescription in template '{}' ({}) does not match the plan prescription ({}).".format(template_name, plan_rx, template_rxs[0]))
        else:  # Multiple Rx's specified in template
            self.label("None of the prescriptions in template '{}' matches the plan prescription ({}).".format(template_name, plan_rx))
        self.label("You may scale the doses in the template goals to the plan prescription. The scale factor is the ratio of the plan prescription to the prescription specified in the goal.")
        self.label("Example:", True) 
        self.label("    Template goal is 'Dmean < 30 Gy' for 5040 Gy.")
        self.label("    Plan prescription is 5000 cGy.")
        self.label("    The scaled goal is 'Dmean < (5000 / 5400 * 30) Gy' -> 'Dmean < 27.78 Gy'.")
        self.label("What would you like to do?")

        self.y += 5

        # Choose whether to scale goals to plan Rx
        self.no_scale_rb = RadioButton()
        self.no_scale_rb.AutoSize = True
        self.no_scale_rb.Checked = True
        self.no_scale_rb.Text = "Use template with no changes"
        self.no_scale_rb.Location = Point(30, self.y)
        self.Controls.Add(self.no_scale_rb)
        self.y += self.no_scale_rb.Height

        self.scale_rb = RadioButton()
        self.scale_rb.AutoSize = True
        self.scale_rb.Text = "Scale Clinical Goals to plan Rx"
        self.scale_rb.Location = Point(30, self.y)
        self.Controls.Add(self.scale_rb)
        self.y += self.scale_rb.Height + 15
            
        if len(template_rxs) > 1:  # Choose template Rx(s) to use
            self.label("The template contains goals for the following prescriptions.")
            self.label("Which prescription(s)' goals would you like to use?")
            self.y += 5
            
            self.lb = ListBox()
            self.lb.SelectionMode = SelectionMode.MultiExtended  # User may select multiple Rx's
            self.lb.Items.AddRange([str(template_rx) for template_rx in template_rxs])  # No Rx's selected by default
            self.lb.SelectedIndexChanged += self.set_ok_enabled
            self.lb.Location = Point(30, self.y)
            self.lb.Height = self.lb.PreferredHeight  # "Autosize" ListBox to number of items
            self.lb.Width = max(TextRenderer.MeasureText(item, self.lb.Font).Width for item in self.lb.Items)  # "Autosize" listbox width to widest item
            self.Controls.Add(self.lb)
            self.y += self.lb.Height + 15

        # OK button
        self.ok = Button()
        self.ok.Location = Point(15, self.y)
        self.ok.Click += self.ok_clicked
        self.ok.Text = "OK"
        self.set_ok_enabled()
        self.AcceptButton = self.ok
        self.Controls.Add(self.ok)

        self.ShowDialog()  # Launch window

    def label(self, text, bold=False):
        # Add a Label to the form

        lbl = Label()
        lbl.AutoSize = True
        if bold:
            lbl.Font = Font(lbl.Font, FontStyle.Bold)
        lbl.Location = Point(15, self.y)
        lbl.Text = text
        self.Controls.Add(lbl)
        self.y += lbl.Height
    
    def set_ok_enabled(self, sender=None, event=None):
        # Enable or disable "OK" button
        # Enable only if user has checked whether or not to scale to Rx, and has selected template Rx's to use (if applicable)
        # Is called from __init__ not as an event handler, so need default args for `sender` and `event`

        self.ok.Enabled = (self.no_scale_rb.Checked or self.scale_rb.Checked) and (not hasattr(self, "lb") or self.lb.SelectedItems.Count > 0)
    
    def ok_clicked(self, sender, event):
        # Event handler for OK button click

        self.DialogResult = DialogResult.OK


def apply_templates(**kwargs):
    """Apply colorwash, structure, and/or clinical goals templates to plan

    The templates to apply are determined based on the keyword arguments (if `gui` is False) or the selected GUI options (if `gui` is True).
    If the same structure is part of multiple selected structure sets, the geometry comes from the first selected structure set.
    Info on template types:
        - Colorwash is "CRMC Standard Dose Colorwash", but you will not see the template name in RS "Edit color table" window (applying a colorwash template is an unscriptable action)
        - Structure
          Template name depends on patient position, plan type, and body site: "TG-263 [plan type ][patient position ("Prone" or "Supine")]<body site>"
          Apply couch template ("Elekta Couch" or "Elekta Prone Couch") if planning exam structure set is unapproved, and each couch structure either does not exist or is empty. Use initialization option "AlignImageCenters"
          Apply structure templates by atlas-based segmentation. Ignore approved ROIs. Update all derived geometries added by the template.
          Recolor targets.
            * GTV: yellow shade
            * CTV: orange shade
            * PTV: red shade
        - Clinical Goals
          Read clinical goals from "template" (worksheet) in Excel workbook "Clinical Goals.xlsx"
          Add spreadsheet goals as well as:
            * If Rx is to volume of PTV:
                - D95% and V95% goals for PTV
                - If PTV is derived from CTV, D100% and V100% for CTV
            * If Rx is provided, Dmax (D0.03)
          Notes column in Excel file may specify:
            * Body site, if template applies to multiple body sites
            * Rx, if template applies to multiple Rx's
            * "Ipsilateral" or "Contralateral"
            * Other info unused by this function
          If Rx's are specified in a template name, only the goals with the plan Rx are applied. If plan Rx does not match any of the Rx's in the template, user may choose to scale goals to plan Rx and, if applicable, the template Rx's to use.
          If there is ambiguity in which template(s) to apply, user chooses template(s) from a GUI.

    Keyword Arguments
    -----------------
    gui: bool
        True if user should choose settings from a GUI, False otherwise
        If True, other keyword arguments are ignored
        Note that even if False, other GUIs may appear to resolve ambiguity in selecting template info
        Defaults to False
    md: str
        Name of the MD: last name or "Other"
        Defaults to MD specified in plan (or "Other" if no MD is specified or MD not in `mds` list)
    plan_type: str
        Type of plan
        Should be in `plan_types` list
        Defaults to an algorithmically chosen plan type
    body_site: str
        Body site of the plan
        Should be in `body_sites` list
        Defaults to an algorithmically chosen body site
    selected_templates: List[str]
        List of types of templates to apply
        Each element should be in `templates` list
        Defaults to all supported templates

    Assumptions
    -----------
    No plan contains multiple nodal PTVs.
    Primary Rx is not to nodal PTV.
    Nodal PTV name contains "PTVn".
    In the clinical goals spreadsheet, Rx to primary PTV is "Rx" or "Rxp", and Rx to nodal PTV is "Rxn".
    There are no POIs without geometries.
    """

    try:
        beam_set = get_current("BeamSet")
    except:
        MessageBox.Show("The current plan has no beam sets. Click OK to abort script.", "No Beam Sets")
        sys.exit(1)

    # Ensure that this is a photon beam set
    if beam_set.Modality != "Photons":
        MessageBox.Show("The current beam set is not a photon beam set. Click OK to abort script.", "Incorrect Modality")
        sys.exit(1)  # Exit with an error

    warnings = ""  # Warnings to display at end of script (if there were any)

    ## Determine default options

    # Default selected MD is MD last name if MD is specified and in `mds` list, "Other" otherwise
    if case.Physician.Name is not None:
        selected_md = case.Physician.Name.split("^")[0]  # MD last name
        if selected_md not in mds:  # Unsupported MD (e.g., "Sidrys" since he doesn't have any custom templates) -> use "Other"
            selected_md = "Other"
    else:  # Unspecified MD -> use "Other"
        selected_md = "Other"
    
    # Default selected plan type is based on fractionation and delivery technique of current beam set
    # This code based on code from RS support
    fx = beam_set.FractionationPattern
    if fx is not None:
        fx = fx.NumberOfFractions
    if beam_set.PlanGenerationTechnique == "Imrt":
        if beam_set.DeliveryTechnique == "DynamicArc":
            if fx in [1, 3]:
                selected_plan_type = "SRS"
            elif fx == 5:
                selected_plan_type = "SBRT"
            else:
                selected_plan_type = "VMAT"
        else:
            selected_plan_type = "IMRT"
    else:
        selected_plan_type = "3D"

    # Default selected body site is the site specified in case name, body site, plan name, plan comments, or beam set name (checked in that order)
    selected_body_site = [site for site in body_sites if any(site in item for item in [case.CaseName, case.BodySite, plan.Name, plan.Comments, beam_set.DicomPlanLabel])]
    selected_body_site = selected_body_site[0] if len(selected_body_site) == 1 else "Other"  # "Other" if no or multiple identified body sites

    # Get options from user
    gui = kwargs.get("gui", False)  # Default: no GUI
    if gui:
        form = ApplyTemplatesForm(selected_md, selected_plan_type, selected_body_site)
        if form.DialogResult != DialogResult.OK:  # "OK" button was not clicked
            sys.exit()

        # Text of checked RadioButtons
        # We can assume each groupbox/listbox has a selection since OK button is disabled otherwise
        md = [rb.Text for rb in form.md_gb.Controls if rb.Checked][0]
        plan_type = [rb.Text for rb in form.plan_type_gb.Controls if rb.Checked][0]
        body_site = [rb.Text for rb in form.body_site_gb.Controls if rb.Checked][0]

        selected_templates = form.templates_lb.SelectedItems
    else:
        md = kwargs.get("md", selected_md)
        if md not in mds:  
            raise ValueError("Invalid `md` keyword argument: '{}'. Valid values are {}.".format(md, mds + [None]))

        # Use algorithmically determined defaults (from above) if plan/body site not provided
        plan_type = kwargs.get("plan_type", selected_plan_type)
        if plan_type not in plan_types:  # None or invalid
            raise ValueError("Invalid `plan_type` keyword argument: '{}'. Valid values are {}.".format(plan_type, format_list(plan_types + [None])))
        body_site = kwargs.get("body_site", selected_body_site)
        if body_site not in body_sites:
            raise ValueError("Invalid `body_site` keyword argument: '{}'. Valid values are {}.".format(body_site, format_list(body_sites + [None])))

        selected_templates = kwargs.get("selected_templates", templates)  # Default to all supportd template types
        if not isinstance(selected_templates, list) and not isinstance(selected_templates, tuple):  # Use default if `selected_templates` is not a list or a tuple
            raise ValueError("Invalid `selected_templates` keyword argument: '{}'. Must be list, tuple, or None.".format(selected_templates))
        invalid_templates = [template for template in selected_templates if template not in templates]
        if invalid_templates:
            raise ValueError("The following template types in keyword argument `selected_templates` are not supported: {}. Value values are {}.".format(format_list(invalid_templates), format_list(templates)))

    # Apply colorwash template
    if "Colorwash" in selected_templates:
        case.CaseSettings.DoseColorMap.ColorTable = patient_db.LoadTemplateColorMap(templateName="CRMC Standard Dose Colorwash").ColorMap.ColorTable

    # Apply structure template(s)
    if "Structure" in selected_templates:
        # Pt position is "Prone" or "Supine"
        # e.g., "HeadFirstSupine" -> "Supine"
        pt_pos = beam_set.PatientPosition.split("First")[-1]
        
        # Select possible templates
        all_struct_templates = [info["Name"] for info in patient_db.GetPatientModelTemplateInfo() if info["Name"].startswith("TG-263")]  # All structure template names
        struct_templates = [template for template in all_struct_templates if re.match("TG-263 ({} )?({} )?{}*".format(plan_type, pt_pos, body_site), template)]  # Structure template names that fit user's selection criteria; e.g., "TG-263 Brain - Basic"
        
        # Select from possible template(s)
        if len(struct_templates) != 1:  # Not exactly 1 matching template
            if struct_templates:  # Multiple matching templates
                form = ChooseTemplatesForm(sorted(struct_templates), "structure", True)  # Choose from multiple matching templates, with the appropriate prompt
            else:  # No matching templates
                form = ChooseTemplatesForm(sorted(all_struct_templates), "structure", False)  # Choose from all templates, with the appropriate prompt
            if form.DialogResult != DialogResult.OK:  # User didn't click "OK"
                sys.exit()
            struct_templates = form.lb.SelectedItems  # Values that user selected in the Form
        
        ## Apply couch template

        # Load template
        template_name = "Elekta Prone Couch" if pt_pos == "Prone" else "Elekta Couch"
        template = patient_db.LoadTemplatePatientModel(templateName=template_name)
        
        # Only apply couch template if planning exam structure set is not approved, and all couch structures either don't exist or are empty on planning exam
        existing_roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
        if struct_set.ApprovedStructureSets.Count == 0 and all(struct.Name not in existing_roi_names or not struct_set.RoiGeometries[struct.OfRoi.Name].HasContours() for struct in template.PatientModel.RegionsOfInterest):
            with CompositeAction("Apply Couch Template"):
                apply_struct_template(template_name, "AlignImageCenters")
                await_user_input("Translate couch structures so that the top of the shell is at the top of the linac table.")  # User translates couch
                center_couch()  # Move couch to RL center
        
        # Apply selected templates
        for template in struct_templates:
            with CompositeAction("Apply Template '{}'".format(template)):
                apply_struct_template(template)

        # Recolor targets
        # Hues are approximately equally spaced for each target type
        colors = {"Gtv": Color.Yellow, "Ctv": Color.Orange, "Ptv": Color.Red}
        with CompositeAction("Recolor targets"):
            for target_type, color in colors.items():
                targets = [roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == target_type]  # All ROIs of that type
                for target in targets:
                    try:
                        new_a = 128 + i * (255.0 - 128) / len(targets)  # Constrain to brigher hues so that color is obvious
                        target.Color = Color.FromArgb(new_a, color.R, color.G, color.B)
                    except:  # ROI is approved, so can't change color
                        pass

    # Apply Clinical Goals template(s)
    if "Clinical Goals" in selected_templates:
        if plan.Review is not None and plan.Review.ApprovalStatus == "Approved":
            warnings += "Plan is approved, so clinical goals could not be added.\n"
        else:
            # All ROI names in current case, with extra info removed
            # According to TG-263, "extra info" is specified after a carat
            # Remove extra info to make ROI name match ROI names in clinical goals templates
            case_rois = [roi.Name.split("^")[0] for roi in case.PatientModel.RegionsOfInterest]

            # Clear existing Clinical Goals
            with CompositeAction("Clear Clinical Goals"):
                while plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count > 0:
                    plan.TreatmentCourse.EvaluationSetup.DeleteClinicalGoal(FunctionToRemove=plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions[0])

            # If Rx is specified, add Dmax goal
            rx_ = beam_set.Prescription.PrimaryDosePrescription
            rx = None  # Assume no Rx dose value
            if rx_ is not None:
                rx = int(rx_.DoseValue)

                # Add Dmax goal
                d_max = 1.25 if plan_type == "SBRT" else 1.1
                ext = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"]  # Select external ROI
                if ext:  # If there is an external (there will only be one), add Dmax goal
                    try:
                        plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ext[0].Name, GoalCriteria="AtMost", GoalType="DoseAtAbsoluteVolume", AcceptanceLevel=d_max * rx, ParameterValue=0.03)  # e.g., D0.03 < 4400 for 4000 cGy non-SBRT plan
                    except:  # Clinical goal already exists
                        pass

                # If Rx is to volume of PTV, add CTV D100%, CTV V100%, PTV D95%, and PTV V95%
                if rx_.PrescriptionType == "DoseToVolume" and rx_.OnStructure.Type == "Ptv":  # Rx is to volume of PTV
                    # Add PTV goals
                    ptv = rx_.OnStructure
                    try:
                        plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ptv.Name, GoalCriteria="AtLeast", GoalType="DoseAtVolume", ParameterValue=0.95, AcceptanceLevel=rx, Priority=1)  # D95%
                    except:  # Clinical goal already exists
                        pass
                    try:
                        plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ptv.Name, GoalCriteria="AtLeast", GoalType="VolumeAtDose", ParameterValue=0.95 * rx, AcceptanceLevel=1, Priority=1)  # V95%
                    except:
                        pass

                    # If PTV is derived, CTV is the target that PTV is derived from
                    ptv_alg = ptv.DerivedRoiExpression
                    # Go "deeper" until you find the CTV
                    if ptv_alg:
                        while ptv_alg.Children:
                            ptv_alg = ptv_alg.Children[0]
                        if ptv_alg.RegionOfInterest.Type == "Ctv":
                            ctv = ptv_alg.RegionOfInterest
                            try:
                                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ctv.Name, GoalCriteria="AtLeast", GoalType="DoseAtVolume", ParameterValue=1, AcceptanceLevel=rx, Priority=1)  # D100%
                            except:
                                pass
                            try:
                                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ctv.Name, GoalCriteria="AtLeast", GoalType="VolumeAtDose", ParameterValue=rx, AcceptanceLevel=1, Priority=1)  # V100%
                            except:
                                pass

            # Read all sheets, ignoring "Planning Priority" column
            # Dictionary of sheet name : DataFrame
            data = pd.read_excel(os.path.join("T:", "Physics - T", "Scripts", "Data", "Clinical Goals.xlsx"), columns=["ROI", "Goal", "Visualization Priority", "Notes"], sheet_name=None)
            
            # Select possible templates
            # Template names that match MD, plan type, and body site, or don't specify these
            # All non-SBRT/-SRS plans can use "Conventional" templates
            template_names = {name: goals for name, goals in data.items() if (md in name or not any(md_name in name for md_name in mds)) and \
                                                              (plan_type in name or not any(plan_type_name in name for plan_type_name in plan_types)) and \
                                                              (plan_type not in ["SRS", "SBRT"] or "Conventional" not in name) and \
                                                              (body_site in name or not any(body_site_name in name for body_site_name in body_sites))}
            
            # Select from possible templates
            if len(template_names) != 1:  # 1 matching template
                if template_names:  # Multiple matching templates
                    form = ChooseTemplatesForm(sorted(template_names), "clinical goals", True)  # Choose from multiple matching templates, with the appropriate prompt
                else:
                    form = ChooseTemplatesForm(sorted(data), "clinical goals", False)  # Choose from all templates, with the appropriate prompt
                if form.DialogResult != DialogResult.OK:  # User didn't click "OK"
                    sys.exit()
                template_names = {name: data[name] for name in form.lb.SelectedItems}

            # Information that will be displayed as warnings later
            invalid_goals = []  # Goals in template that are in an invalid format
            empty_spare = []  # Volume-to-spare goals that cannot be added due to empty geometry (format: "<ROI name>:  <goal>", e.g., "Liver:  V21 Gy < (v-700) cc")
            no_ipsi_contra = False  # Whether or not ipsilateral/contralateral goals could not be added due to indeterminable Rx side
            no_nodal_ptv = []  # If a nodal PTV does not exist, goals for a nodal PTV 

            ## Determine Rx side, used when adding ispi/contra goals
            
            # Get initial laser isocenter, just in case it is needed
            ini_laser_iso = [poi for poi in case.PatientModel.PointsOfInterest if poi.Type == "InitialLaserIsocenter"]
            rx_ctr = struct_set.PoiGeometries[ini_laser_iso[0].Name].Point.x if ini_laser_iso else None
            
            if hasattr(rx_, "OnStructure"):
                struct = rx_.OnStructure
                if struct.OrganData is not None and struct.OrganData.OrganType == "Target":  # Rx is to ROI
                    rx_ctr = struct_set.RoiGeometries[struct.Name].GetCenterOfRoi().x  # R-L center of ROI
                else:  # Rx is to POI
                    rx_ctr = struct_set.PoiGeometries[struct.Name].Point.x
            elif hasattr(rx_, "OnDoseSpecificationPoint"):  # Rx is to site
                if rx_.OnDoseSpecificationPoint is not None:  # Rx is to DSP
                    rx_ctr = rx_.OnDoseSpecificationPoint.Coordinates.x
                else:  # Rx is to site that is not a DSP
                    dose_dist = plan.TreatmentCourse.TotalDose
                    if dose_dist.DoseValues is not None and dose_dist.DoseValues.DoseData is not None:
                        rx_ctr = dose_dist.GetCoordinateOfMaxDose().x

            ## Apply templates
            for template_name, goals in template_names.items():
                # "Fine-tune" the goals to apply
                # Check fractionation, Rx, body site, side, etc.
                template_rxs = []  # Rx value in Notes column of a goal to use
                scale = False  # Scale goal to Rx in template_rxs?
                if md != "Other":  
                    # Check # Fx in MD-specific template (irrelevant for Mobius templates)
                    fx = beam_set.FractionationPattern
                    if fx is not None:  # Don't check anything if fractionation isn't specified in beam set
                        fx = fx.NumberOfFractions
                        data_fx = re.search("(\d+) Fx", template_name)  # Find "__ Fx" in template name
                        if data_fx is not None:  # Template name contains Fx
                            data_fx = int(data_fx.group(1))  # Template number of fractions
                            if fx != data_fx:  # Fx mismatch
                                msg = "The number of fractions in template '{}' ({}) does not match the planned number of fractions ({}). Apply the template anyway?".format(template_name, data_fx, fx)
                                res = MessageBox.Show(msg, "Incorrect Fractionation", MessageBoxButtons.YesNo)
                                if res == DialogResult.No:  # User does not want to use template w/ incorrect Fx
                                    continue  # Skip this template

                    # Check Rx in template name, if plan contains Rx
                    if rx is not None:
                        data_rx = re.search("(([\d\.]+ )+)Gy", template_name)  # Find "__ Gy" in template name (e.g., "40.4 50.4 cGy")
                        if data_rx:
                            data_rx = [int(float(data_rx) * 100) for data_rx in data_rx.group(1).split()]  # All Rx's specified in template name, converted to cGy
                            if rx in data_rx:  # Plan Rx is specified in template name
                                template_rxs.append(rx)  # Only use goals with plan Rx or no specified Rx
                            else:  # Plan Rx is not specified in template name 
                                form = TemplateRxForm(template_name, rx, data_rx)  # User chooses to ignore Rx mismatch (default) or scale goals to plan Rx
                                if hasattr(form, "lb"):  # Multiple Rx's in template name, so user chose which ones to use
                                    template_rxs.extend([int(item) for item in form.lb.SelectedItems])
                                scale = form.scale_rb.Checked  # Did user check "scale"?
                ## Add goals   
                        
                goals["ROI"] = pd.Series(goals["ROI"]).fillna(method="ffill")  # Autofill ROI name (due to vertically merged cells in spreadsheet)
            
                with CompositeAction("Apply Clinical Goals Template '{}'".format(template_name)):
                    for _, row in goals.iterrows():  # Iterate over each row in DataFrame
                        args = {}  # dict of arguments for ApplyTemplates
                        scale_factor = 1  # Assume we're not scaling this goal
                        roi = row["ROI"]  # e.g., "Lens"
                        rois = [r for r in case_rois if re.match("^{}(_[LR])?$".format(roi), r)]  # Matching ROIs (account for side); e.g., "Lens", "Lens_L", "Lens_R"
                        if not rois:  # ROI in goal does not exist in case
                            continue

                        # If present, notes may be Rx, body site, body side, or info irrelevant to script
                        notes = row["Notes"]
                        if not pd.isna(notes):  # Notes exist
                            # Goal only applies to specific Rx
                            m = re.match("([\d\.]+) Gy", notes)
                            if m is not None: 
                                notes = int(float(m.group(1)) * 100)  # Extract the number and convert to cGy
                                if notes in template_rxs:  # Use the goal and scale if necessary
                                    if scale:
                                        scale_factor = float(rx) / notes  # Plan Rx / template Rx (convert an operand to float to avoid integer division)
                                else:  # Skip the goal
                                    continue
                            # Goal is for a body site that is not the chosen site
                            elif notes in template_name and notes != body_site:  
                                continue
                            
                            # Ipsilateral objects have same sign on x-coordinate (so product is positive); contralateral have opposite signs (so product is negative)
                            elif notes in ["Ipsilateral", "Contralateral"]:
                                if rx_ctr is None:
                                    no_ipsi_contra = True
                                else:
                                    rois = [r for r in rois if (notes == "Ipsilateral" and rx_ctr * struct_set.RoiGeometries[r].GetCenterOfRoi().x > 0) or (notes == "Contralateral" and rx_ctr * struct_set.RoiGeometries[r].GetCenterOfRoi().x < 0)]  # Select the ipsilateral or contralateral matching ROIs
                            # Otherwise, irrelevant info

                        # Visualization Priority (note that this is NOT the same as planning priority)
                        if not pd.isna(row["Visualization Priority"]):  
                            args = {"Priority": row["Visualization Priority"]}
                        
                        goal = re.sub("\s", "", row["Goal"])  # Remove spaces in goal
                        
                        ## Parse dose and volume amounts from goal. Then add clinical goal for volume or dose.

                        # Regexes to match goal
                        dose_amt_regex = """(
                                                (?P<dose_pct_rx>[\d.]+%)?
                                                (?P<dose_rx>Rx[pn]?)|
                                                (?P<dose_amt>[\d.]+)
                                                (?P<dose_unit>c?Gy)
                                            )"""  # e.g., 95%Rx or 20Gy
                        dose_types_regex = "(?P<dose_type>max|mean|median)"
                        vol_amt_regex = """(
                                                (?P<vol_amt>[\d.]+)
                                                (?P<vol_unit>%|cc)|
                                                (\(v-(?P<spare_amt>[\d.]+)\)cc)
                                        )"""  # e.g., 67%, 0.03cc, or v-700cc
                        sign_regex = "(?P<sign><|>)"  # > or <

                        dose_regex = """D
                                        ({}|{})
                                        {}
                                        {}""".format(dose_types_regex, vol_amt_regex, sign_regex, dose_amt_regex)  # e.g., D0.03cc<110%Rx, Dmedian<20Gy

                        vol_regex = """V
                                        {}
                                        {}
                                        {}
                                    """.format(dose_amt_regex, sign_regex, vol_amt_regex)  # e.g., V20Gy<67%

                        # Need separate regexes b/c we can't have duplicate group names in a single regex
                        # Remove whitespace from regex (left in above for readability) before matching
                        vol_match = re.match(re.sub("\s", "", vol_regex), goal)
                        dose_match = re.match(re.sub("\s", "", dose_regex), goal)
                        match = vol_match if vol_match is not None else dose_match  # If it's not a volume, should be a dose

                        if not match:  # Invalid goal format -> add goal to invalid goals list and move on to next goal
                            invalid_goals.append(goal)
                            continue

                        args["GoalCriteria"] = "AtMost" if match.group("sign") == "<" else "AtLeast"  # GoalCriteria depends on sign

                        # Extract dose: an absolute amount or a % of Rx
                        dose_rx = match.group("dose_rx")
                        if dose_rx:  # % of Rx
                            dose_pct_rx = match.group("dose_pct_rx")  # % of Rx
                            if dose_pct_rx is None:  # Group not present. % of Rx is just specified as "Rx"
                                dose_pct_rx = 100
                            else:  # A % is specified, so make sure format is valid
                                try:
                                    dose_pct_rx = float(dose_pct_rx[:-1])  # Remove percent sign and convert to float
                                    if dose_pct_rx < 0:  # % out of range -> add goal to invalid goals list and move on to next goal
                                        invalid_goals.append(goal)
                                        continue
                                except:  # % is non numeric -> add goal to invalid goals list and move on to next goal
                                    invalid_goals.append(goal)
                                    continue
                            # Find appropriate Rx (to primary or nodal PTV)
                            if dose_rx == "Rxn":  # Use 2ry Rx (to nodal PTV)
                                rx_n = [rx_n for rx_n in beam_set.Prescription.DosePrescriptions if "PTVn" in rx_n.OnStructure.Name]
                                if rx_n:  # Found a nodal PTV (should never be more than one)
                                    dose_rx = rx_n[0]
                                else:  # There is no nodal PTV, so add goal to list of goals that could not be added to nodal PTV, and move on to next goal
                                    no_nodal_ptv.append(goal)
                            else:  # Primary Rx
                                dose_rx = rx_
                            dose_amt = dose_pct_rx / 100 * dose_rx.DoseValue  # Get absolute dose based on % Rx
                        else:  # Absolute dose
                            try:
                                dose_amt = float(match.group("dose_amt")) * scale_factor  # Account for scaling to template Rx if user selected this option (remember that `scaling_factor` is 1 otherwise)
                            except:  # Given dose amount is non-numeric  -> add goal to invalid goals list and move on to next goal
                                invalid_goals.append(goal)
                                continue
                            if match.group("dose_unit") == "Gy":  # Covert dose from Gy to cGy
                                dose_amt *= 100
                        if dose_amt < 0 or dose_amt > 100000:  # Dose amount out of range  -> add goal to invalid goals list and move on to next goal
                            invalid_goals.append(goal)
                            continue

                        # Extract volume: an absolute amount, a % of ROI volume, or an absolute amount to spare
                        dose_type = vol_unit = spare_amt = None
                        vol_amt = match.group("vol_amt")
                        if vol_amt:  # Absolute volume or % of ROI volume
                            try:
                                vol_amt = float(vol_amt)
                            except:  # Given volume is non numeric -> add goal to invalid goals list and move on to next goal
                                invalid_goals.append(goal)
                                continue

                            vol_unit = match.group("vol_unit")
                            if vol_unit == "%":  # If relative volume, adjust volume amount
                                if vol_amt > 100:  # Given volume is out of range -> add goal to invalid goals list and move on to next goal
                                    invalid_goals.append(goal)
                                    continue
                                vol_amt /= 100  # Convert percent to proportion

                            if vol_amt < 0 or vol_amt > 100000:  # Volume amount out of range supported by RS -> add goal to invalid goals list and move on to next goal
                                invalid_goals.append(goal)
                                continue
                        else:  # Volume to spare or dose type
                            spare_amt = match.group("spare_amt")
                            if spare_amt:  # Volume to spare
                                try:
                                    spare_amt = float(spare_amt)
                                except:  # Volume amount is non numeric -> add goal to invalid goals list and move on to next goal
                                    invalid_goals.append(goal)
                                    continue
                                
                                geom = struct_set.RoiGeometries[roi]
                                if not geom.HasContours():  # Cannot add volume to spare goal for empty geometry -> add goal to list of vol-to-spare goals for empty geometries
                                    empty_spare.append("{}:\t{}".format(roi, goal))
                                    continue
                                if spare_amt < 0 or spare_amt > geom.GetRoiVolume():  # Spare amount out of range -> add goal to invalid goals list and move on to next goal
                                    invalid_goals.append(goal)
                                    continue
                            else:  # Dose type: Dmax, Dmean, or Dmedian
                                dose_type = match.group("dose_type")

                        # D...
                        if goal.startswith("D"):
                            # Dmax = D0.035
                            if dose_type == "max":
                                args["GoalType"] = "DoseAtAbsoluteVolume"
                                args["ParameterValue"] = 0.03
                            # Dmean => "AverageDose"
                            elif dose_type == "mean":
                                args["GoalType"] = "AverageDose"
                            # Dmedian = D50%
                            elif dose_type == "median":
                                args["GoalType"] = "DoseAtVolume"
                                args["ParameterValue"] = 0.5
                            # Absolute or relative dose
                            else:
                                args["ParameterValue"] = vol_amt
                                if vol_unit == "%":
                                    args["GoalType"] = "DoseAtVolume"
                                else:
                                    args["GoalType"] = "DoseAtAbsoluteVolume"
                            args["AcceptanceLevel"] = dose_amt
                        # V...
                        else:
                            args["ParameterValue"] = dose_amt
                            if vol_unit == "%":
                                args["GoalType"] = "VolumeAtDose"
                            else:
                                args["GoalType"] = "AbsoluteVolumeAtDose"
                            if not spare_amt:
                                args["AcceptanceLevel"] = vol_amt
                            
                        # Add Clinical Goals
                        for roi in rois:
                            roi_args = args.copy()
                            roi_args["RoiName"] = roi
                            if spare_amt:
                                total_vol = struct_set.RoiGeometries[roi].GetRoiVolume()
                                roi_args["AcceptanceLevel"] = total_vol - spare_amt
                            plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(**roi_args)

            # Add warnings about clinical goals that were not added
            if invalid_goals:
                warnings += "The following clinical goals could not be parsed so were not added:\n\t-  {}\n".format("\n\t-  ".join(invalid_goals))
            if empty_spare:
                warnings += "The following clinical goals could not be added due to empty geometries:\n\t-  {}\n".format("\n\t-  ".join(empty_spare))
            if no_ipsi_contra:
                warnings += "There is no Rx, so ipsilateral/contralateral structures could not be determined. Ipsilateral/contralateral clinical goals were not added."
            if no_nodal_ptv:
                warnings += "No nodal PTV was found, so the following clinical goals were not added:\n\t-  {}\n".format("\n\t-  ".join(no_nodal_ptv))

    # Display warnings if there were any
    if warnings != "":
        MessageBox.Show(warnings, "Warnings")

    sys.exit()  # For some reason, script won't exit on its own if warnings are displayed
