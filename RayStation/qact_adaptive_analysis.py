import clr
import re
import sys
from typing import List, Optional

from connect import *  # Interact with RayStation
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System import EventArgs  # For type hints
from System.Drawing import *
from System.Windows.Forms import *


QACT_EXAM_DATE_REGEX = r'(\d{1,2}[/\-\. ]\d{1,2}[/\-\. ](\d{2}|\d{4}))|(\d{6}|\d{8})'


def add_date_to_exam_name(exam) -> None:
    """Adds a date to the end of an exam name, if a date does not exist

    A date in the exam name matches the regular expression `QACT_EXAM_DATE_REGEX`

    The date is in the format "YYYY-MM-DD"
    The date is taken from the series acqusition date in the exam's DICOM. If this is not provided, the study acquisition date is used. If neither acquisition date is present, the exam name is unchanged

    Argument
    --------
    exam: The exam whose name to append a date to

    Example
    -------
    Exam name 'An Exam' becomes 'An Exam 2022-06-03' if the exam's series or study date is 6/3/2022
    """
    if re.search(QACT_EXAM_DATE_REGEX, exam.Name) is None:  # Exam name does not contain a date
        dcm = exam.GetAcquisitionDataFromDicom()
        date = dcm['SeriesModule']['SeriesDateTime']  # We prefer the series's acqusition date over the study's
        if date is None:  # Try the study's acqusition date if the series doesn't have one
            date = dcm['StudyModule']['StudyDateTime']
        if date is not None:  # There is a series or study acquisition date
            # Append "YYYY-MM-DD" date to exam name
            # Date is a System.DateTime object
            exam.Name = f'{exam.Name} {date.Year}-{date.Month:0>2}-{date.Day:0>2}'


class QACTAdaptiveAnalysisForm(Form):
    """Windows Form that allows user to select parameters for an analysis of whether a replan is needed on an adaptive exam

    User chooses:
    - TPCT (reference exam)
    - QACT (floating)
    - Geometries to copy versus deform

    Support geometries are not allowed to be deformed
    """
    def __init__(self, tpct, qacts, tpct_geoms):
        self.Text = 'QACT Adaptive Analysis'

        # Adapt form size to contents
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        
        self.MouseClick += self._MouseClick

        y = 15  # Vertical coordinate of next control

        # TPCT name
        tpct_lbl = Label()
        tpct_lbl.AutoSize = True
        tpct_lbl.Font = Font(tpct_lbl.Font, FontStyle.Bold)
        tpct_lbl.Location = Point(15, y)
        tpct_lbl.Text = 'TPCT:    '
        self.Controls.Add(tpct_lbl)
        tpct_name_lbl = Label()
        tpct_name_lbl.AutoSize = True
        tpct_name_lbl.Location = Point(15 + tpct_lbl.Width, y)
        tpct_name_lbl.Text = tpct
        tpct_name_lbl.UseMnemonic = False  # Display ampersands
        y += tpct_lbl.Height + 15
        self.Controls.Add(tpct_name_lbl)

        # QACT
        qact_lbl = Label()
        qact_lbl.AutoSize = True
        qact_lbl.Font = Font(qact_lbl.Font, FontStyle.Bold)
        qact_lbl.Location = Point(15, y)
        qact_lbl.Text = 'QACT:    '
        self.Controls.Add(qact_lbl)
        if len(qacts) == 1:  # If only one possible QACT, display the name
            self.qact = qacts[0]#.Name
            qact_name_lbl = Label()
            qact_name_lbl.AutoSize = True
            qact_name_lbl.Location = Point(15 + qact_lbl.Width, y)
            qact_name_lbl.Text = self.qact
            qact_name_lbl.UseMnemonic = False  # Display ampersands
            self.Controls.Add(qact_name_lbl)
        else:  # If multiple possible QACTs, allow user to choose
            #qacts = [qact.Name for qact in qacts]
            self.qact_combo = ComboBox()
            self.qact_combo.DropDownStyle = ComboBoxStyle.DropDownList
            self.qact_combo.Items.AddRange(qacts)
            self.qact_combo.Location = Point(15 + qact_lbl.Width, y)
            self.qact_combo.SelectedIndex = 0
            self.qact_combo.Width = max(TextRenderer.MeasureText(qact, self.qact_combo.Font).Width for qact in qacts) + 25
            self.Controls.Add(self.qact_combo)
        y += tpct_lbl.Height + 15

        #geom_names = [geom.OfRoi.Name for geom in ss.RoiGeometries if geom.HasContours()]
        lbl_width = self._get_lbl_width([geom.OfRoi.Name for geom in tpct_geoms]) * 2  # Double the necessary width for better formatting
        
        # "Geometry" label
        geom_lbl = Label()
        geom_lbl.AutoSize = True
        geom_lbl.MinimumSize = Size(lbl_width, 0)
        geom_lbl.Font = Font(geom_lbl.Font, FontStyle.Bold)
        geom_lbl.Location = Point(15, y)
        geom_lbl.Text = 'Geometry'
        self.Controls.Add(geom_lbl)

        # "Copy" label
        copy_lbl = Label()
        copy_lbl.AutoSize = True
        copy_lbl.Font = Font(copy_lbl.Font, FontStyle.Bold)
        copy_x = 15 + lbl_width
        copy_lbl.Location = Point(copy_x, y)
        copy_lbl.Text = 'Copy'
        self.Controls.Add(copy_lbl)

        # "Deform" label
        deform_lbl = Label()
        deform_lbl.AutoSize = True
        deform_lbl.Font = Font(deform_lbl.Font, FontStyle.Bold)
        deform_x = copy_x + int(copy_lbl.Width * 1.5)
        deform_lbl.Location = Point(deform_x, y)
        deform_lbl.Text = 'Deform'
        self.Controls.Add(deform_lbl)
        y += int(deform_lbl.Height * 1.25)  # Point() cannot take a float

        # "Select all" copy checkbox
        self.copy_all_cb = CheckBox()
        self.copy_all_cb.AutoSize = True
        self.copy_all_cb.IsThreeState = True
        self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Checked
        self.copy_all_cb.MinimumSize = Size(deform_x - copy_x, 0)  # At least as wide as deform label + space between deform and copy labels
        self.copy_all_cb.Click += self._copy_all_cb_Click
        self.copy_all_cb.Location = Point(copy_x, y)
        self.copy_all_cb.Text = 'All'
        self.Controls.Add(self.copy_all_cb)

        # "Select all" deform checkbox
        self.deform_all_cb = CheckBox()
        self.deform_all_cb.AutoSize = True
        self.deform_all_cb.IsThreeState = True
        self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Unchecked
        self.deform_all_cb.MinimumSize = Size(deform_x - copy_x, 0)  # At least as wide as deform label + space between deform and copy labels
        self.deform_all_cb.Click += self._deform_all_cb_Click
        self.deform_all_cb.Location = Point(deform_x, y)
        self.deform_all_cb.Text = 'All'
        self.Controls.Add(self.deform_all_cb)
        y += int(self.deform_all_cb.Height * 1.25)  # Point() cannot take a float

        # Zebra-striped geometry rows (alternate LightGray and the default Label background color)
        colors = [Color.LightGray, geom_lbl.BackColor]
        self.geom_cbs = {}  # ROI: [copy checkbox, deform checkbox]
        for i, geom in enumerate(tpct_geoms):
            color = colors[i % 2]  # Alternate the two colors
            # Label
            geom_name_lbl = Label()
            geom_name_lbl.AutoSize = True
            geom_name_lbl.MinimumSize = Size(lbl_width, 0)
            geom_name_lbl.BackColor = color
            geom_name_lbl.Location = Point(15, y)
            geom_name_lbl.Text = geom.OfRoi.Name
            tpct_name_lbl.UseMnemonic = False  # Display ampersands
            self.Controls.Add(geom_name_lbl)

            # "Copy" checkbox
            copy_cb = CheckBox()
            copy_cb.AutoSize = True
            copy_cb.Checked = True
            copy_cb.MinimumSize = Size(deform_x - copy_x, 0)
            copy_cb.BackColor = color
            copy_cb.Click += self._copy_cb_Click
            copy_cb.Location = Point(copy_x, y)
            copy_cb.Name = 'Copy' + geom.OfRoi.Name  # e.g., 'CopyPTV'
            self.Controls.Add(copy_cb)

            # "Deform" checkbox
            if geom.OfRoi.Type in ['Fixation', 'Support']:  # Cannot deform a fixation or support geometry
                self.geom_cbs[geom.OfRoi.Name] = [copy_cb, None]
                placeholder = Label()
                placeholder.BackColor = color
                placeholder.Location = Point(deform_x, y)
                placeholder.Size = Size(deform_x - copy_x, copy_cb.Height)  # Placeholder is same size as a copy checkbox
                self.Controls.Add(placeholder)
                y += int(placeholder.Height * 1.25)  # Point() cannot take a float
            else:
                deform_cb = CheckBox()
                deform_cb.AutoSize = True
                deform_cb.Checked = False
                deform_cb.MinimumSize = Size(deform_x - copy_x, 0)
                deform_cb.BackColor = color
                deform_cb.Click += self._deform_cb_Click
                deform_cb.Location = Point(deform_x, y)
                deform_cb.Name = 'Deform' + geom.OfRoi.Name  # e.g., 'DeformPTV'
                self.Controls.Add(deform_cb)
                self.geom_cbs[geom.OfRoi.Name] = [copy_cb, deform_cb]
                y += int(deform_cb.Height * 1.25)  # Point() cannot take a float
        y += 15

        # Form is at least as wide as title text plus room for 'X' button, etc.
        min_width = TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 150
        self.MinimumSize = Size(min_width, self.Height)

        # 'Start' button
        self._start_btn = Button()
        self._start_btn.AutoSize = True
        self._start_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self._start_btn.Click += self._start_btn_Click
        self._start_btn.Enabled = False
        self._start_btn.Text = 'Start'
        b_x = self.ClientSize.Width - self._start_btn.Width  # Right align
        self._start_btn.Location = Point(b_x, y)
        self.Controls.Add(self._start_btn)
        self.AcceptButton = self._start_btn
    
    def _get_lbl_width(self, lbl_txt: List[str]) -> str:
        """Computes the largest Label width necessary for any of the strings in list

        Argument
        --------
        lbl_txt: List of strings that will be displayed in a Label
        """
        lbl_width = 0
        for txt in lbl_txt:
            l = Label()
            l.Text = txt
            l.UseMnemonic = False  # Treat ampersand as a literal characater
            lbl_width = max(lbl_width, l.Width)
        return lbl_width
    
    def _MouseClick(self, sender: Form, event: EventArgs) -> None:
        """Handles the event of clicking the Form
        
        Removes focus when user clicks the mouse
        """
        self.ActiveControl = None
    
    def _set_copy_checkstate(self) -> None:
        """Sets checkstate for copy "All" checkbox based on checkstates of copy checkboxes"""
        copy_cbs_cked = [cbs[0].Checked for cbs in self.geom_cbs.values()]
        if all(copy_cbs_cked):
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Checked
        elif any(copy_cbs_cked):
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Indeterminate
        else:
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Unchecked

    def _set_deform_checkstate(self) -> None:
        """Sets checkstate for deform "All" checkbox based on checkstates of deform checkboxes"""
        deform_cbs = [cbs[1] for cbs in self.geom_cbs.values() if cbs[1] is not None]  # Deform checkboxes that exist
        deform_cbs_cked = [cb.Checked for cb in deform_cbs]
        if all(deform_cbs_cked):
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Checked
        elif any(deform_cbs_cked):
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Indeterminate
        else:
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Unchecked

    def _copy_cb_Click(self, sender: CheckBox, event: EventArgs) -> None:
        """Handles the event of clicking a copy checkbox
        
        Extracts ROI name from checkbox name and unchecks that ROI's deform checkbox if it exists
        Sets "All" checkstates
        """
        # Set checkstate of 'select all' checkboxes
        geom = sender.Name[4:]  # ROI name is the part of the checkbox name after "Copy"
        deform_cb = self.geom_cbs[geom][1]
        if deform_cb is not None:  # Ignore if the geometry cannot be deformed (deform checkbox does not exist)
            deform_cb.Checked = False
        self._set_deform_checkstate()
        self._set_copy_checkstate()

    def _deform_cb_Click(self, sender: CheckBox, event: EventArgs) -> None:
        """Handles the event of clicking a deform checkbox
        
        Extracts ROI name from checkbox name and unchecks that ROI's copy checkbox
        Sets "All" checkstates
        """
        geom = sender.Name[6:]  # ROI name is the part of the checkbox name after "Deform"
        copy_cb = self.geom_cbs[geom][0]
        copy_cb.Checked = False
        self._set_copy_checkstate()
        self._set_deform_checkstate()

    def _copy_all_cb_Click(self, sender: CheckBox, event: EventArgs) -> None:
        """Handles the event of clicking the copy "All" checkbox
        
        Sets appropriate checkstate based on previous checkstate, stored in Tag attribute
        Sets item checkbox states based on new state
        """
        if self.copy_all_cb.Tag == CheckState.Checked:  # If checked, uncheck
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Unchecked
        else:  # If unchecked or indeterminate, check
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Checked
        # Set item checkstates
        for cbs in self.geom_cbs.values():
            cbs[0].Checked = self.copy_all_cb.CheckState  # Check or uncheck to match "All" checkbox
            # Uncheck the deform checkbox if it exists
            if self.copy_all_cb.CheckState and cbs[1] is not None and cbs[1].Checked:
                cbs[1].Checked = False
        self._set_deform_checkstate()

    def _deform_all_cb_Click(self, sender: CheckBox, event: EventArgs) -> None:
        """Handles the event of clicking the deform "All" checkbox
        
        Sets appropriate checkstate based on previous checkstate, stored in Tag attribute
        Sets item checkbox states based on new state
        """
        if self.deform_all_cb.Tag == CheckState.Checked:  # If checked, uncheck
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Unchecked
        else:  # If unchecked or indeterminate, check
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Checked
        # Set item checkstates
        for cbs in self.geom_cbs.values():
            if cbs[1] is not None:  # Ignore the row if the deform checkbox does not exist
                cbs[1].Checked = self.deform_all_cb.CheckState  # Check or uncheck to match "All" checkbox
                if self.deform_all_cb.CheckState and cbs[0].Checked:  # If deform checkbox is now checked and copy is also checked, uncheck copy
                    cbs[0].Checked = False
        self._set_copy_checkstate()

    def _start_btn_Click(self, sender: Button, event: EventArgs) -> None:
        """Handles the event that is clicking the "Start" button"""
        self.DialogResult = DialogResult.OK


def get_reg_from_reg_list(ui: PyScriptObject, exam: PyScriptObject) -> Optional[PyScriptObject]:
    """Gets the rigid image registration that has the given exam as the target

    Arguments
    ---------
    ui: The ui RaySearch object whose GUI elements will be searched for the registration
    exam: The exam to find the rehistration to

    Returns
    -------
    The registration UI element if it exists, None otherwise
    """
    for reg in ui.ToolPanel.RegistrationList.TreeItem:
        m = re.match(r'<.+\'(.+)\'>', str(reg))  # A registration of the correct type has string representation e.g., "<ScriptObject id=RayStationMainWindow._ws.ToolPanel.RegistrationList.TreeItem@0, 'CT 2'>""
        reg_exam_name = m.group(1)  # Extract target exam name from registration string representation
        if reg_exam_name == exam.Name:  # Correct exam name
            return reg


def qact_adaptive_analysis():
    """Performs an analysis on a TPCT and a QACT to determine whether an adaptive plan is needed

    Using a GUI, user chooses QACT (if there are multiple possible options) and geometries to copy or deform
    Note: Approved geometries on the QACT are not copied/deformed

    1. User chooses QACT (if there are options), ROIs to copy, and ROIs to deform, from a GUI.
    2. Rigidly registers TPCT to QACT, if necessary. User can adjust registration if desired.
    3. If necessary, deforms TPCT to QACT (with a focus on bone).
    4. Deforms POI geometries from TPCT to QACT.
    5. Copies the specified ROI geometries from TPCT to QACT. Crops to QACT field-of-view.
    6. Deforms the specified ROI geometries from TPCT to QACT.
    7. For each beam set that has dose in the current plan, computes dose on the QACT. Updates dose grid structures for the new evaluation dose.
    8. Adds date and " (TPCT)"/" (QACT)" to TPCT/QACT exam names, if necessary.
    """
    # Get current objects
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()  # Exit script
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
        sys.exit()

    # Ensure multiple examinations, not just the TPCT, exist
    if case.Examinations.Count == 1:  # Only the TPCT exists
        MessageBox.Show('Case contains only one exam. Click OK to abort the script.', 'No Possible QACTs')
        sys.exit()

    ss = plan.GetTotalDoseStructureSet()
    tpct_geoms = [geom for geom in ss.RoiGeometries if geom.HasContours()]
    tpct = ss.OnExamination  # TPCT is planning exam

    # Ensure external ROI exists
    try:
        ext_geom = next(geom for geom in ss.RoiGeometries if geom.OfRoi.Type == 'External')
    except StopIteration:
        MessageBox.Show('There is no external ROI. Click OK to abort the script.', 'No External ROI')
        sys.exit()

    # Ensure external is contoured on TPCT
    if not ext_geom.HasContours():  # No external on TPCT
        MessageBox.Show('There is no external geometry on the TPCT. Click OK to abort the script.', 'No External Geometry')
        sys.exit()

    # Possible QACTs are not the TPCT, not registered to TPCT in opposite direction, and not in same frame of reference as TPCT
    reg_to_tpct = [reg for reg in case.Registrations if reg.StructureRegistrations['Source registration'].ToExamination.Equals(tpct)]
    qact_names = []
    for exam in case.Examinations:
        # Exam is the TPCT
        if exam.Equals(tpct):
            continue
        # Exam is in same FoR as TPCT
        if exam.EquipmentInfo.FrameOfReference == tpct.EquipmentInfo.FrameOfReference:
            continue
        # Exam is registered to TPCT
        try:
            next(reg for reg in reg_to_tpct if reg.StructureRegistrations['Source registration'].FromExamination.Equals(exam))
            continue
        except StopIteration:
            pass
        # External is not empty and approved on the exam
        exam_ss = case.PatientModel.StructureSets[exam.Name]
        try:
            next(geom for approved_ss in ss.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures if geom.OfRoi.Type == 'External' and not geom.HasContours())
            continue
        except StopIteration:
            pass
        qact_names.append(exam.Name)
    if not qact_names:  # No other exams w/ an external
        criteria = ['It is not the TPCT (planning exam for current plan).', 
                    'It is not registered with the TPCT in the opposite direction (QACT -> TPCT).',
                    'It is not in the same frame of reference as the TPCT.',
                    'It either has an external geometry, or its external ROI is unapproved (so a geometry can be added).']
        msg = f'There are no possible QACTs. All of the following must be true of an examination used as a QACT:\n-  ' + '\n-  ' .join(criteria) + '\nClick OK to abort the script.'
        MessageBox.Show(msg, 'No Possible QACTs')
        sys.exit()

    form = QACTAdaptiveAnalysisForm(tpct.Name, qact_names, tpct_geoms)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:  # User cancelled GUI window
        sys.exit()  

    # Extract attributes from Form
    if hasattr(form, 'qact_combo'):
        qact_name = form.qact_combo.SelectedItem
    else:
        qact_name = form.qact
    qact = case.Examinations[re.sub(r'&&', r'&', qact_name)]
    qact_ss = case.PatientModel.StructureSets[qact.Name]
    if not qact_ss.RoiGeometries[ext_geom.OfRoi.Name].HasContours():
        case.PatientModel.RegionsOfInterest[ext_geom.OfRoi.Name].CreateExternalGeometry(Examination=qact)
    copy_geoms = [geom for geom, cbs in form.geom_cbs.items() if cbs[0].Checked]
    deform_geoms = [geom for geom, cbs in form.geom_cbs.items() if cbs[1] is not None and cbs[1].Checked]

    # Compute rigid registration, if necessary
    try:
        reg = next(reg for reg in case.Registrations if reg.StructureRegistrations['Source registration'].FromExamination.Equals(tpct) and reg.StructureRegistrations['Source registration'].ToExamination.Equals(qact))
    except StopIteration:
        case.ComputeRigidImageRegistration(FloatingExaminationName=qact.Name, ReferenceExaminationName=tpct.Name, HighWeightOnBones=True, FocusRoisNames=[])

        # Navigate to registration and allow user to make changes
        ui = get_current('ui')
        ui.TitleBar.Navigation.MenuItem['Patient modeling'].Button_Patient_modeling.Click()
        ui = get_current('ui')  # New UI so that "Image Registration" tab is available
        ui.TabControl_Modules.TabItem['Image registration'].Select()
        ui.ToolPanel.TabItem['Registrations'].Select()
        ui = get_current('ui')  # New UI so that list of registrations is available
        tree_item = get_reg_from_reg_list(ui, qact)
        tree_item.Select()
        ui.ToolPanel.TabItem['Scripting'].Select()
        await_user_input('Review the rigid registration and make any necessary changes. Then click the play button to resume the script.')

    # Deformable registration
    try:
        grp_name = next(srg.Name for srg in case.PatientModel.StructureRegistrationGroups for dsr in srg.DeformableStructureRegistrations if dsr.FromExamination.Equals(tpct) and dsr.ToExamination.Equals(qact))  # Find deformable reg in structure reg group for TPCT -> QACT
    except StopIteration:
        grp_name = tpct.Name + ' to ' + qact.Name
        case.PatientModel.CreateHybridDeformableRegistrationGroup(RegistrationGroupName=grp_name, ReferenceExaminationName=tpct.Name, TargetExaminationNames=[qact.Name], AlgorithmSettings={ 'NumberOfResolutionLevels': 3, 'InitialResolution': { 'x': 0.5, 'y': 0.5, 'z': 0.5 },'FinalResolution': { 'x': 0.25, 'y': 0.25, 'z': 0.25 }, 'InitialGaussianSmoothingSigma': 2.0, 'FinalGaussianSmoothingSigma': 0.333, 'InitialGridRegularizationWeight': 1500.0, 'FinalGridRegularizationWeight': 400.0, 'ControllingRoiWeight': 0.5, 'ControllingPoiWeight': 0.1, 'MaxNumberOfIterationsPerResolutionLevel': 1000, 'ImageSimilarityMeasure': 'CorrelationCoefficient', 'DeformationStrategy': 'Default', 'ConvergenceTolerance': 1e-5})

    # Deform POI geometries
    #poi_names = [geom.OfPoi.Name for geom in case.PatientModel.StructureSets[tpct.Name].PoiGeometries if abs(geom.Point.x) < 1000]  # POI w/ no geometry has infinite coordinates
    #case.MapPoiGeometriesDeformably(PoiGeometryNames=poi_names, StructureRegistrationGroupNames=[grp_name], ReferenceExaminationNames=[tpct.Name], TargetExaminationNames=[qact.Name])

    # Copy ROI geometries
    if copy_geoms: 
        case.PatientModel.CopyRoiGeometries(SourceExamination=tpct, TargetExaminationNames=[qact.Name], RoiNames=copy_geoms)
        # Crop to FOV
        fov_name = case.PatientModel.GetUniqueRoiName(DesiredName='FOV')
        fov = case.PatientModel.CreateRoi(Name=fov_name, Type='FieldOfView')
        fov.CreateFieldOfViewROI(ExaminationName=qact.Name)
        for roi_name in copy_geoms:  # Intersect FOV and each geometry on exam, if possible
            case.PatientModel.RegionsOfInterest[roi_name].CreateAlgebraGeometry(Examination=qact, ExpressionA={ 'Operation': 'Union', 'SourceRoiNames': [roi_name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [fov.Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation='Intersection', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
        fov.DeleteRoi()  # Delete unnecessary ROI

    # Deform ROI geometries
    if deform_geoms:
        case.MapRoiGeometriesDeformably(RoiGeometryNames=deform_geoms, StructureRegistrationGroupNames=[grp_name], ReferenceExaminationNames=[tpct.Name], TargetExaminationNames=[qact.Name])

    # Create eval dose for each beam set
    bs_no_dose, bs_w_dose = [], []
    for bs in plan.BeamSets:
        if bs.FractionDose.DoseValues is None:
            bs_no_dose.append(bs.DicomPlanLabel)
        else:
            bs_w_dose.append(bs.DicomPlanLabel)
            bs.ComputeDoseOnAdditionalSets(ExaminationNames=[qact.Name], FractionNumbers=[0])
    # Update dose grid structures for the new eval doses
    for fe in case.TreatmentDelivery.FractionEvaluations:
        for doe in fe.DoseOnExaminations:
            if doe.OnExamination.Equals(qact):
                des = list(doe.DoseEvaluations)
                for i in range(len(bs_w_dose)):
                    des[-i].UpdateDoseGridStructures()

    # Map dose (alternative to 'compute dose')
    #deformable = case.Registrations[0].StructureRegistrations['Deformable Registration1']
    #case.MapDose(DoseDistribution=dose_dist, StructureRegistration=deformable)

    # Rename TPCT and QACT if needed
    add_date_to_exam_name(tpct)
    add_date_to_exam_name(qact)
    if 'TPCT' not in tpct.Name.upper():
        tpct.Name += ' (TPCT)'
    if 'QACT' not in qact.Name.upper():
        qact.Name += ' (QACT)'

    # Alert user if any doses on additional set could not be computed
    if bs_no_dose:
        MessageBox.Show('The following beam set(s) have no dose, so they were not computed on the QACT: ' + ', '.join('"' + bs + '"' for bs in bs_no_dose), 'Warnings')
