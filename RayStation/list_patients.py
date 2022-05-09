import clr
from datetime import datetime
#from os import system

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

from openpyxl import Workbook
from openpyxl.styles import numbers

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


def tx_technique(beam_set: PyScriptObject) -> str:
    """Determines the treatment technique of the beam set

    Code modified from RS support
    
    Arguments
    ---------
    beam_set: The beam set whose treatment technique to return

    Returns
    -------
    The treatment technique ("SMLC", "SRS", "SBRT", "SABR", "VMAT", "DMLC", "Conformal", "Conformal Arc", or "ApplicatorAndCutout"), or "[Unknown]" if the treatment technique could not be determined
    """

    if beam_set.Modality == 'Photons':
        if beam_set.PlanGenerationTechnique == 'Imrt':
            if beam_set.DeliveryTechnique == 'SMLC':
                return 'SMLC'
            if beam_set.DeliveryTechnique == 'DynamicArc':
                if beam_set.Prescription is not None and beam_set.FractionationPattern is not None:
                    fx = beam_set.FractionationPattern.NumberOfFractions
                    rx = beam_set.Prescription.PrimaryPrescriptionDoseReference.DoseValue / fx
                    if rx >= 600:
                        if fx in [1, 3]:
                            return 'SRS'
                        if fx == 5:
                            return 'SBRT'
                        if fx in [6] + list(range(8, 16)):
                            return 'SABR'
                return 'VMAT'
            if beam_set.DeliveryTechnique == 'DMLC':
                return 'DMLC'
        elif beam_set.PlanGenerationTechnique == 'Conformal':
            if beam_set.DeliveryTechnique == 'SMLC':
                # return 'SMLC' # Changed from 'Conformal'. Failing with forward plans.
                return 'Conformal'
                # return '3D-CRT'
            if beam_set.DeliveryTechnique == 'Arc':
                return 'Conformal Arc'
    elif beam_set.Modality == 'Electrons' and beam_set.PlanGenerationTechnique == 'Conformal' and beam_set.DeliveryTechnique == 'SMLC':
        return 'ApplicatorAndCutout'
    return "[Unknown]"


class ListPatientsForm(Form):
    """Form that allows the user to specify filters for patient selection"""
    # Absolute path of directory to save output file in
    # This directory does not have to already exist
    OUTPUT_DIR = os.path.join('T:', os.sep, 'Physics', 'Scripts', 'Output Files', 'List Patients')

    def __init__(self):
        # Public variables (set in "OK" button click event handler)
        self.max_num_pts = float('inf')

        self._set_up_form()
        
        self._y = 15  # Vertical coordinate of next control

        # Script description/instructions
        text = 'MRNs of patients meeting all of the following filter criteria will be written to a file in "' + self.OUTPUT_DIR + '".\n' \
             + 'The filename includes a timestamp.\n' \
             + 'Check the filters that you want to apply.'
        l = self._add_lbl(text)
        self._y += l.Height + 15

        # Number of patients to output
        self._add_num_pts_input()

        # 'Select all filters' checkbox
        self._select_all_filters_cb = self._add_checkbox(select_all=True, txt='Select all filters')
        self._y += self._select_all_filters_cb.Height + 10

        self._filter_cbs, self._filter_gbs = [], []  # Lists of checkboxes and their corresponding groupboxes, for each filter

        # Keywords filter
        instrs = 'The following are searched for the keywords (case insensitive):\n    -  '
        instrs += '\n    -  '.join(['Case names, body sites, comments, and diagnoses', 'Exam names', 'Plan names and comments', 'Beam set names and comments', 'Beam names and descriptions', 'DSP names', 'Rx descriptions and structure names'])
        instrs += '\nEnter one keyword per line:'
        gb, gb_y = self._add_filter('Keywords', instrs)

        self._keyword_tb = TextBox()
        self._set_up_keyword_tb()
        gb.Controls.Add(self._keyword_tb)
        self._y += gb.Height + 10

        # Sex filter
        gb = self._add_checkboxes_filter('Sex', ['Male', 'Female', 'Other'])
        self._y += gb.Height + 10

        # Patient Position filter
        gb = self._add_checkboxes_filter('Patient Position', ['FFS', 'HFP', 'HFS'])
        self._y += gb.Height + 10

        # Tx technique filter
        gb = self._add_checkboxes_filter('Treatment Technique', ['Applicator and cutout', 'Conformal', 'Conformal Arc', 'DMLC', 'SABR', 'SBRT', 'SRS', 'SMLC', 'VMAT', '[Unknown]'])
        self._y += gb.Height + 15

        # Warning label
        # Visible if no filters will be applied
        self._warning_lbl = self._add_lbl(txt='No filters are applied.\nAll MRNs will be written to the file.')
        self._y += self._warning_lbl.Height

        # 'OK' button
        self._ok_btn = Button()
        self._set_up_ok_btn()

    # ------------------------------ Event handlers ------------------------------ #

    def _num_pts_tb_TextChanged(self, sender, event):
        self._set_ok_enabled()

    def _keyword_tb_TextChanged(self, *args):
        self._set_warning_label()
        self._set_ok_enabled()

    def _select_all_Click(self, sender, event):
        # Helper method for clicking a 'Select all' checkbox
        # Set checkstate of options checkboxes
        # If the checkbox is 'Select all filters', enable or disable the corresponding groupboxes

        checked = sender.Tag == CheckState.Checked  # Checkbox tag tracks previous checkstate
        if checked:  # 'Select all' is now checked -> uncheck
            sender.CheckState = sender.Tag = CheckState.Unchecked
        else:  # 'Select all' is now unchecked or indeterminate -> check
            sender.CheckState = sender.Tag = CheckState.Checked

        # Set options checkstates
        parent = sender.Parent
        if isinstance(parent, Form):  # It's a filter checkbox
            for i, cb in enumerate(self._filter_cbs):
                gb = self._filter_gbs[i]
                cb.Checked = gb.Enabled = not checked  # Enable groupbox only if checkbox is checked
        else:  # It's a checkbox within a filter
            item_cbs = [ctrl for ctrl in parent.Controls if isinstance(ctrl, CheckBox) and ctrl.Tag is None]  # All checkboxes in the groupbox that are not 'select all'
            for cb in item_cbs:
                cb.Checked = not checked  # Toggle checkstate

    def _item_checkbox_Click(self, sender, event):
        # Helper method for clicking a checkbox that is not 'Select all'
        # Set corresponding 'Select all' checkstate

        parent = sender.Parent
        select_all = next(ctrl for ctrl in parent.Controls if isinstance(ctrl, CheckBox) and ctrl.Tag is not None)  # "Select all" checkbox in this group
        if isinstance(parent, Form):  # A filter checkbox was clicked
            gb = self._filter_gbs[self._filter_cbs.index(sender)]  # Corresponding groupbox
            gb.Enabled = sender.Checked  # Enable groupbox only if checkbox is now checked

        item_cbs_checked = [ctrl.Checked for ctrl in parent.Controls if isinstance(ctrl, CheckBox) and ctrl.Tag is None]  # All checkboxes in the groupbox that are not "Select all"
        if all(item_cbs_checked):  # All checkboxes are checked
            select_all.CheckState = select_all.Tag = CheckState.Checked  # Check "Select all"
        elif any(item_cbs_checked):  # Some checkboxes are checked, some unchecked
            select_all.CheckState = select_all.Tag = CheckState.Indeterminate  # "Select all" checkstate is indeterminate
        else:  # All checkboxes are unchecked
            select_all.CheckState = select_all.Tag = CheckState.Unchecked  # Uncheck "Select all"
        self._set_warning_label()  # Warning label visibility depends on whether any filters are applied

    def _ok_btn_Click(self, sender, event):
        # Event handler for clicking the 'OK' button
        # Create `values` attribute to hold user-set filters

        self.max_num_pts = float('Inf') if self._num_pts_tb.Text == '[All]' else int(self._num_pts_tb.Text)
        self.values = {}  # keyword name : values
        for i, cb in enumerate(self._filter_cbs):  # Iterate over all filter checkboxes
            gb = self._filter_gbs[i]  # Corresponding groupbox
            name = gb.Text[:-1]  # Remove colon at end of groupbox text to get keyword name
            if cb.Checked:  # User wants to use the filter
                cbs = [ctrl for ctrl in gb.Controls if isinstance(ctrl, CheckBox) and ctrl.Tag is None]  # Non-"Select all" checkboxes
                if cbs:  # There are checkboxes, so this is not the keywords checkbox
                    self.values[name] = [cb.Text for cb in cbs if cb.Checked]  # Text of all checked checkboxes
                else:  # It's the keywords filter
                    self.values[name] = [keyword.strip().lower() for keyword in self._keyword_tb.Text.split('\n')]  # Each non-blank line in the keywords textboc
            else:  # User does not want to use the filter
                self.values[name] = None
        self.DialogResult = DialogResult.OK

    # ---------------------------------- Styling --------------------------------- #

    def _set_up_form(self):
        """Styles this Form"""
        self.Text = 'List Patients'  # Form title
        
        # Adapt form size to controls
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)  # At least as wide as title plus room for "X" button, etc.

        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

    def _set_up_keyword_tb(self):
        self._keyword_tb.AutoSize = True
        self._keyword_tb.Location = Point(15, gb_y)
        self._keyword_tb.Width = 50
        self._keyword_tb.MinimumSize = Size(100, 75)
        self._keyword_tb.Multiline = True  # Accepts more than one line
        self._keyword_tb.TextChanged += self._keyword_tb_TextChanged  # Empty vs. filled textbox influences whether any filters are actually applied

    def _set_up_ok_btn(self):
        self._ok_btn.Location = Point(self.ClientSize.Width - 50, self._y)  # Right align
        self._ok_btn.Text = 'OK'
        self._ok_btn.Click += self._ok_btn_Click
        self.Controls.Add(self._ok_btn)

    # ------------------------------- Add basic controls ------------------------------- #

    def _add_checkbox(self, **kwargs):
        parent = kwargs.get('parent', self)
        x = kwargs.get('x', 15)
        y = kwargs.get('y', parent._y)
        is_select_all = kwargs.get('is_select_all', False)
        txt = kwargs.get('txt', 'Select all' if is_select_all else '')

        cb = CheckBox()
        cb.AutoSize = True
        cb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        if is_select_all:
            cb.Click += self._select_all_checkbox_Click
            cb.IsThreeState = True
            cb.CheckState = cb.Tag = CheckState.Unchecked
        else:
            cb.Click += self._item_checkbox_Click
            cb.Checked = False
        cb.Location = Point(x, y)
        cb.Text = txt
        parent.Controls.Add(cb)
        return cb

    def _add_groupbox(self, **kwargs):
        x = kwargs.get('x', 15)
        txt = kwargs.get('txt', '')
        if txt != '':
            txt += ':'

        gb = GroupBox()
        gb.Enabled = False
        gb.AutoSize = True
        gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        gb.MinimumSize = Size(TextRenderer.MeasureText(txt, gb.Font).Width + 50, 0)
        gb.Location = Point(50, self._y)
        gb.Text = txt
        self.Controls.Add(gb)
        return gb

    def _add_lbl(self, **kwargs) -> Label:
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
        parent = kwargs.get('parent', self)
        x = kwargs.get('x', 15)
        y = kwargs.get('y', parent._y)
        bold = kwargs.get('bold', False)
        txt = kwargs.get('txt', '')

        l = Label()
        l.AutoSize = True
        l.AutoSizeMode = AutoSizeMode.GrowAndShrink
        if bold:
            l.Font = Font(l.Font, FontStyle.Bold)
        l.Location = Point(x, y)
        l.Text = txt
        parent.Controls.Add(l)
        return l

    # -------------------------------- Add user inputs -------------------------------- #

    def _add_num_pts_input(self):
        self._add_lbl('Max # of MRN(s) to write:')

        self._num_pts_tb = TextBox()
        self._num_pts_tb.Location = Point(20 + l.Width, self._y)
        self._num_pts_tb.Text = '[All]'
        self._num_pts_tb.Width = 25
        self._num_pts_tb.TextChanged += self._num_pts_tb_TextChanged
        self.Controls.Add(self._num_pts_tb)
        self._y += 50

    def _add_filter(self, name, instrs=None):
        # Helper method that adds a checkbox and groupbox for a filter
        # `name`: Title of the filter
        # `instrs`: Instructions to display to the user
        # Return a 2-tuple: the groupbox and the y-coordinate for the next control in the groupbox

        # Checkbox
        cb = self._add_checkbox(x=30)

        ## Groupbox
        gb = self._add_groupbox(x=50, txt=name)
        gb_y = 15  # Vertical coordinate of next control in groupbox

        # Filter description/instructions
        if instrs is not None:
            l = self._add_lbl(y=gb_y, txt=instrs)
            gb_y += l.Height

        # Add new checkbox and groupbox to lists
        self._filter_cbs.append(cb)
        self._filter_gbs.append(gb)

        return gb, gb_y

    def _add_checkboxes_filter(self, name, items, instrs=None):
        # Helper method that adds a filter that is checkboxes
        # `name`: Title of the filter
        # `items`: List of checkbox options
        # `instrs`: Instructions to display to the user

        gb, gb_y = self._add_filter(name, instrs)  # Add checkbox and groupbox
        
        # 'Select all' checkbox
        cb = self._add_checkbox(parent=gb, y=gb_y, select_all=True)
        gb_y += cb.Height

        # Item checkboxes
        for item in items:
            cb = self._add_checkbox(parent=gb, y=gb_y, txt=item)
            gb_y += cb.Height
        return gb

    # ------------- Warning label and "OK" button based on selections ------------ #

    def _set_warning_label(self):
        # Helper method that makes the warning label visible if no filters are applied, invisible otherwise
        # Effectively, a filter is not applied if the checkbox is unchecked (obviously), or if no keywords are provided (for keywords filter) or all checkboxes are checked (for other filters)
        # *args instead of `sender` and `event` because this method can also be called not an event handler

        vis = True  # Assume warning label should be visible
        for i, cb in enumerate(self._filter_cbs):
            if cb.CheckState == CheckState.Checked:  # Filter is checked, so if there's any meaningful input, it is applied
                gb = self._filter_gbs[i]  # Corresponding groupbox
                select_all = [ctrl for ctrl in gb.Controls if isinstance(ctrl, CheckBox) and ctrl.Tag is not None]  # 'Select all' checkboxes in that groupbox (not present if it's the keywords groupbox)
                if select_all:  # It's a checkboxes groupbox
                    select_all = select_all[0]  # The 'Select all' checkbox in the group
                    if select_all.CheckState == CheckState.Indeterminate:  # Indeterminate checkstate means some checkboxes arechecked and some unchecked, so there is meaningful input
                        vis = False  # No warning necessary since a filter is applied
                        break
                elif self._keyword_tb.Text.strip() != '':  # It's the keywords groupbox. If there's nothing but whitespace in the keywords textbox, the keyword filter is effectively not applied
                    vis = False  # There is text, so no warning necessary
                    break
        self._warning_lbl.Visible = vis  # Show/hide

    def _set_ok_enabled(self):
        # Helper method that enables/disables the 'OK' button
        # Enable 'OK' button only if the number of MRN(s) is set to '[All]' or a positive integer

        num_pts = self._num_pts_tb.Text
        try:
            num_pts = int(num_pts)
            self._ok_btn.Enabled = num_pts > 0
        except:
            self._ok_btn.Enabled = num_pts == '[All]'
        
        if self._ok_btn.Enabled:
            self._num_pts_tb.BackColor = Color.White
        else:
            self._num_pts_tb.BackColor = Color.Red


def list_patients():
    '''Write a CSV file with the MRNs of RayStation patients matching all the user-selected filters

    Filters:
    - Keywords: Any of the provided keywords is part of any of the following:
        * A case name, body site, comment, or dx
        * A plan name or comment
        * An exam name
        * A beam set name or comment
        * A beam name or description
        * A DSP name
        * An Rx description or structure name
    - Sex
    - Patient position: The patient position of an exam is as specified. If keywords are also provided, this exam must belong to the plan matching a keyword.
    - Treatment technique: The treatment delivery technique of a beam set. If keywords are also provided, this beam set must belong to the plan matching a keyword.

    Filename is the datetime and the number of matching patients
    Note that patients that are currently open will not be checked!
    If no patients are found, file contains one line: 'No matching patients.'
    '''

    # Get filters from user
    form = ListPatientsForm()
    if form.DialogResult != DialogResult.OK:  # 'OK' button was never clicked (user probably exited GUI)
        sys.exit()

    max_num_pts = form.max_num_pts
    keywords = form.values['Keywords']
    sex = form.values['Sex']
    pt_pos = form.values['Patient Position']
    tx_techniques = form.values['Treatment Technique']

    patient_db = get_current('PatientDB')

    matching_mrns = []  # MRNs to write to output file
    all_pts = patient_db.QueryPatientInfo(Filter={})  # Get all patient info in RS DB
    for pt in all_pts:  # Iterate over all patients in RS
        if len(matching_mrns) == max_num_pts:
            break

        try:
            pt = patient_db.LoadPatient(PatientInfo=pt)  # Load the patient
        except:  # Someone has the patient open
            continue
 
        # If sex filter applied and patient doesn't match, move on to next patient
        if sex is not None and pt.Gender not in sex:
            continue
        
        if keywords is not None:  # Keywords filter is applied
            names_to_chk = []  # Strings to check whether they contain any of the keywords
            for case in pt.Cases:
                for exam in case.Examinations:
                    if pt_pos is None or exam.PatientPosition in pt_pos:
                        names_to_chk.extend([case.BodySite, case.CaseName, case.Comments, case.Diagnosis])
                        names_to_chk.append(exam.Name)
                        for plan in case.TreatmentPlans:
                            if plan.GetTotalDoseStructureSet().OnExamination.Name.lower() == exam.Name.lower():
                                names_to_chk.extend([plan.Name, plan.Comments])
                                for bs in plan.BeamSets:
                                    if tx_techniques is None or get_tx_technique(bs) in tx_techniques:
                                        names_to_chk.extend([bs.Comment, bs.DicomPlanLabel])
                                        for b in bs.Beams:
                                            names_to_chk.extend([b.Description, b.Name])
                                        for dsp in bs.DoseSpecificationPoints:
                                            names_to_chk.append(dsp.Name)
                                        if bs.Prescription is not None:
                                            names_to_chk.append(bs.Prescription.Description)
                                            for rx in bs.Prescription.PrescriptionDoseReferences:
                                                if hasattr(rx, 'OnStructure'):
                                                    names_to_chk.append(rx.OnStructure.Name)
           
            # Check each name for each keyword (case insensitive)
            keyword_match = False  # Assume no match
            for keyword in keywords:
                for name in set(names_to_chk):  # Remove duplicates from names list
                    if name is not None:  # E.g., beam description can be None
                        name = name.lower()
                        if keyword in name and (not keyword.startswith('pelvi') or 'abd' not in name):  # Very special case: if searching for 'pelvis', we don't want, e.g, 'abdomen/pelvis'
                            keyword_match = True
                            break
            if not keyword_match:  # If no keywords present in any of the names, move on to next patient
                continue
        
        # Keywords filter is not applied, so exam with correct patient position doesn't have to be associated with a matching keyword
        elif pt_pos is not None and not any(exam.PatientPosition == pt_pos for case in pt.Cases for exam in case.Examinations):
            continue

        elif tx_techniques is not None and not any(get_tx_technique(bs) in tx_techniques for case in pt.Cases for plan in case.TreatmentPlans for bs in plan.BeamSets):
            continue
        
        matching_mrns.append(pt.PatientID)

    ## Write output file
    wb = Workbook()  # Create Excel data
    ws = wb.active
    
    # Format column as text, not number
    for row in ws:
        row[1].number_format = numbers.FORMAT_TEXT
    
    # Write data to file
    if matching_mrns:  # There are patients matching all criteria -> write in sorted order, one MRN per line
        for i, mrn in enumerate(sorted(matching_mrns)):
            ws.cell(i + 1, 1).value = mrn
    else:  # No matching patients, so say so in the first cell of the output file.
        ws.cell(1, 1).value = 'No matching patients.'

    # Write Excel data to file
    dt = datetime.now().strftime('%m-%d-%y %H_%M_%S')
    filename = '{} (1 Patient).xlsx'.format(dt) if len(matching_mrns) == 1 else '{} ({} Patients).xlsx'.format(dt, len(matching_mrns))
    wb.save(r'\\vs20filesvr01\groups\CANCER\Physics\Scripts\Output Files\ListPatients\{}'.format(filename))

    # Open Excel file
    # No permissions to do this from RS script
    # Could always just write to a CSV file instead of XLSX, but then would lose ability to format as text (and thus the leading zeros in the MRN)
    # excel_path = r'\\Client\C$\Program Files (x86)\Microsoft Office\Office16\excel.exe'
    # system(r'START /B '{}' '{}''.format(excel_path, filepath))
