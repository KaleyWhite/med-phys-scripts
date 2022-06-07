import clr
import sys

from connect import *
from connect.connect_cpython import PyScriptObject

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System import EventArgs
from System.Drawing import *
from System.Windows.Forms import *


class ChooseBeamForm(Form):
    """Form that allows user to choose which beam's isocenter to localize to"""
    def __init__(self, beam_set: PyScriptObject) -> None:
        """Initializes a ChooseBeamForm object

        Arguments
        ---------
        beam_set: Beam set whose beams to display as radio buttons
        """
        self._beam_set = beam_set
        self.iso: Dict[str, float] = None  # Isocenter of the selected beam

        self._set_up_form()

        self._y = 15  # Vertical coordinate of next control
        
        instrs_lbl = Label()
        self._set_up_instrs_lbl(instrs_lbl)

        self._rbs = []
        self._add_beam_rbs()

        self._ok_btn = Button()
        self._set_up_ok_btn()

    # ------------------------------ Event handlers ------------------------------ #

    def _rb_Click(self, sender: RadioButton, event: EventArgs) -> None:
        """Handles the Click event for a radio button

        If any radio button is checked, set self.iso to the selected beam's isocenter position and enable the "OK" button
        Otherwise, set self.iso to None and disable the "OK" button
        """
        try:
            beam_name = next(rb.Text for rb in self._rbs if rb.Checked)
        except StopIteration:
            self.iso = None
            self._ok_btn.Enabled = False
        else:
            self.iso = self._beam_set.Beams[beam_name].Isocenter.Position
            self._ok_btn.Enabled = True

    def _ok_btn_Click(self, sender: Button, event: EventArgs) -> None:
        """Handles the Click event for the "OK" button"""
        self.DialogResult = DialogResult.OK

    # ------------------------------- Setup/layout ------------------------------- #

    def _set_up_form(self):
        """Styles the Form"""
        self.Text = 'Choose Beam'  # Form title

        # Adapt form size to controls
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)

        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ('X out of') it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen

    def _set_up_instrs_lbl(self, instrs_lbl: Label) -> None:
        """Styles the instructions label and adds it to the form
        
        Arguments
        ---------
        instrs_lbl: The instructions label
        """
        instrs_lbl.AutoSize = True
        instrs_lbl.Location = Point(15, self._y)
        instrs_lbl.Text = 'Not all beams in the current beam set have the same isocenter position. Choose the beam isocenter to localize to.'
        self.Controls.Add(instrs_lbl)
        self._y += instrs_lbl.Height

    def _add_beam_rbs(self) -> None:
        """Adds radio buttons to the form, one for each beam in the beam set"""
        for beam in self._beam_set.Beams:
            rb = RadioButton()
            rb.AutoSize = True
            rb.Click += self._rb_Click
            rb.Location = Point(15, self._y)
            rb.Text = beam.Name
            self.Controls.Add(rb)
            self._rbs.append(rb)
            self._y += rb.Height

    def _set_up_ok_btn(self) -> None:
        """Styles and adds the "OK" button to the form"""
        self._ok_btn.Click += self._ok_btn_Click
        self._ok_btn.Enabled = False  # By default, no radio button is checked
        self._ok_btn.Location = Point(self.ClientSize.Width - 50, self._y)  # Right align
        self._ok_btn.Text = 'OK'
        self.AcceptButton = self._ok_btn
        self.Controls.Add(self._ok_btn)


def tab_item_to_str(tab_item: PyScriptObject) -> str:
    """Returns the text of an RS UI TabItem

    Arguments
    ---------
    tab_item: The Ui.TabControl_ToolBar.TabItem element whose string representation to return
    """
    return str(tab_item).split("'")[1]


def localize_to_beam_iso() -> None:
    """Localizes views to the beam isocenter of the current beam set

    If beams in the current beam set have different iso coordinates, user chooses the beam to use, from a GUI
    
    After performing actions in Patient Modeling > Structure Definition, returns the user to the tab from which they ran the script
    """
    # Get current variables
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        beam_set = get_current('BeamSet')
    except:
        MessageBox.Show('There is no beam set loaded. Click OK to abort the script.', 'No Beam Set Loaded')
        sys.exit()
    
    if beam_set.Beams.Count == 0:  # Beam set has no beams (and therefore no possible isos)
        MessageBox.Show('The current beam set has no beams. Click OK to abort the script.', 'No Beams')
        sys.exit()

    # Get iso coordinates
    iso = [beam.Isocenter.Position for beam in beam_set.Beams]  # Iso coordinates of all beams
    if len(set((pos.x, pos.y, pos.z) for pos in iso)) == 1:  # All beams have same iso coordinates
        iso = iso[0]
    else:  # Not all beams have same iso coordinates, so allow user to choose the beam whose iso to localize to
        form = ChooseBeamForm(beam_set)
        form.ShowDialog()
        if form.DialogResult != DialogResult.OK:
            sys.exit()
        iso = form.iso

    # Determine which module/submodule/tab the user is currently in

    # Ui.TabControl_Modules.TabItem element : Ui.TitleBar.Navigation.MenuItem element
    modules = {'Fallback planning': 'automated Planning', 'Patient information': 'Patient data management', 'Image registration': 'Patient modeling', 'Plan setup': 'Plan design', 'Plan optimization': 'Plan optimization', 'Plan evaluation': 'Plan evaluation', 'QA preparation': 'QA preparation'}
    
    # Ui.TabControl_ToolBar.TabItem element : Ui.TabControl_Modules.TabItem element
    submodules = {'Automatic tools': 'Image registration', 'ROI tools': 'Structure definition', 'Deformation': 'Deformable registration'}
    
    # Ui.TabControl_ToolBar.ToolBarGroup element : Ui.TabControl_ToolBar.TabItem element
    toolbar_tabs = {'EXTRAS': 'ROI tools', 'CURRENT POI': 'POI tools', 'STRUCTURE SET APPROVAL': 'Approval', 'LEVEL / WINDOW': 'Fusion'}

    ui = get_current('ui')
    module = modules[tab_item_to_str(ui.TabControl_Modules.TabItem[0])]
    submodule = submodules[tab_item_to_str(ui.TabControl_ToolBar.TabItem[0])] if module == 'Patient modeling' else None
    toolbar_tab = toolbar_tabs[tab_item_to_str(list(ui.TabControl_ToolBar.ToolBarGroup)[-1])] if submodule == 'Structure definition' else None

    # Switch to Patient Modeling > Structure Definition > ROI Tools
    ui.TitleBar.Navigation.MenuItem['Patient modeling'].Click()
    ui.TitleBar.Navigation.MenuItem['Patient modeling'].Popup.MenuItem['Structure definition'].Click()
    ui = get_current('ui')
    ui.TabControl_ToolBar.TabItem['ROI tools'].Select()
    ui = get_current('ui')

    # Create new ROI at iso
    roi_name = case.PatientModel.GetUniqueRoiName(DesiredName='Iso')
    roi = case.PatientModel.CreateRoi(Name=roi_name, Type='Control')
    roi.CreateSphereGeometry(Radius=1, Examination=beam_set.GetPlanningExamination(), Center=iso)

    # Localize to ROI
    ui.TabControl_ToolBar.ToolBarGroup['CURRENT ROI'].Button_Localize_ROI.Click()

    # Return user to previous module/submodule
    if module == 'Patient modeling':
        ui.TabControl_Modules.TabItem[submodule].Select()
        if submodule == 'Structure definition':
            ui = get_current('ui')
            ui.TabControl_ToolBar.TabItem[toolbar_tab].Select()
    else:
        getattr(ui.TitleBar.Navigation.MenuItem[module], 'Button_{}'.format('_'.join(module.split()))).Click()  # E.g., click 'Button_Plan_Design' button

    # Delete iso ROI
    roi.DeleteRoi()
