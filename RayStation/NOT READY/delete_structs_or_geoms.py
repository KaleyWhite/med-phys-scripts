# For GUI
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import os
import sys

from connect import *  # Interact w/ RS

# For GUI
from System.Drawing import *
from System.Windows.Forms import *


case = None


class DeleteAllROIsOrGeometriesForm(Form):
    """Form that allows user to select ROI(s) (default) or geometry(ies) to delete

    Also sets whether to retain external, and whether to retain ROIs/geometries that are empty on all exams
    If geometries are to be deleted, user also selects exams to delete them from
    """

    def __init__(self):
        # Set up Form
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = "Delete All ROIs/Geometries"
        # Form is at least as wide as title text
        min_width = TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100  # Form is at least as wide as title plus "X" button
        self.MinimumSize = Size(min_width, 0)
        y = 15  # Vertical coordinate of next control

        # Add "ROIs" & "Geometries" radio buttons
        rbs_gb = GroupBox()
        rbs_gb.AutoSize = True
        rbs_gb.AutoSizeMode = AutoSizeMode.GrowAndShrink
        rbs_gb.Location = Point(15, y)
        rbs_gb.Text = "Delete:"

        self.rois_rb = RadioButton()
        self.rois_rb.Checked = True
        self.rois_rb.Click += self.rois_rb_clicked
        self.rois_rb.Location = Point(15, 15)
        self.rois_rb.Text = "ROIs"
        rbs_gb.Controls.Add(self.rois_rb)
        
        self.geoms_rb = RadioButton()
        self.geoms_rb.Click += self.geoms_rb_clicked
        self.geoms_rb.Location = Point(15, 15 + self.rois_rb.Height)
        self.geoms_rb.Text = "Geometries"
        rbs_gb.Controls.Add(self.geoms_rb)

        self.Controls.Add(rbs_gb)
        y += rbs_gb.Height + 15

        # Add "settings" checkboxes
        self.ext_cb = CheckBox()
        self.ext_cb.AutoSize = True
        self.ext_cb.Checked = True
        self.ext_cb.Location = Point(15, y)
        self.ext_cb.Text = "Do not delete External"
        self.Controls.Add(self.ext_cb)
        y += self.ext_cb.Height

        self.empty_cb = CheckBox()
        self.empty_cb.AutoSize = True
        self.empty_cb.Checked = True
        self.empty_cb.Location = Point(15, y)
        self.empty_cb.Text = "Only delete empty ROIs"
        self.Controls.Add(self.empty_cb)
        y += self.empty_cb.Height + 15

        # Add exams checkboxes (only visible if deleting geometries, not ROIs)
        self.exams_gb = GroupBox()
        self.exams_gb.AutoSize = True
        self.exams_gb.Location = Point(15, y)
        self.exams_gb.Text = "Delete geometries from:"
        self.exams_gb.Visible = False  # Since default is delete ROIs, hide exams checkboxes

        exams_y = 15
        cb = CheckBox()
        cb.AutoSize = True
        self.exams_gb.Controls.Add(cb)
        cb.CheckState = cb.Tag = CheckState.Checked  # Default all exams selected; use tag to keep track of previous check state
        cb.Click += self.select_all_clicked
        cb.Location = Point(15, exams_y)
        cb.Text = "Select all"
        cb.ThreeState = True
        exams_y += cb.Height
        for exam in case.Examinations:
            cb = CheckBox()
            cb.AutoSize = True
            self.exams_gb.Controls.Add(cb)
            cb.Checked = True
            cb.CheckedChanged += self.exam_clicked
            cb.Location = Point(30, exams_y)
            cb.Text = exam.Name
            exams_y += cb.Height
        self.Controls.Add(self.exams_gb)

        self.no_exams_y = y + 15  # y-coord of delete button when exam checkboxes are hidden
        self.w_exams_y = y + 15 + self.exams_gb.Height  # y-coord of delete button when exam checkboxes are shown

        # Add buttons
        self.delete_button = Button()
        self.delete_button.AutoSize = True
        self.delete_button.Click += self.delete_clicked
        self.delete_button.Location = Point(self.ClientSize.Width - 15 - self.delete_button.Width, self.no_exams_y)
        self.delete_button.Text = "Delete"
        self.Controls.Add(self.delete_button)
        self.AcceptButton = self.delete_button

        self.ShowDialog()  # Display Form

    def rois_rb_clicked(self, sender, event):
        # When "ROIs" radio button is checked, change message about empty contours, hide exam checkboxes, and reposition delete button

        self.empty_cb.Text = "Only delete ROIs that have no contours on any exam"
        self.exams_gb.Visible = False
        self.delete_button.Location = Point(self.ClientSize.Width - 15 - self.delete_button.Width, self.no_exams_y)

    def geoms_rb_clicked(self, sender, event):
        # When "ROIs" radio button is checked, change message about empty contours, show exam checkboxes, and reposition delete button

        self.empty_cb.Text = "Only delete geometries of ROIs that have no contours on any exam"
        self.exams_gb.Visible = True
        self.delete_button.Location = Point(self.ClientSize.Width - 15 - self.delete_button.Width, self.w_exams_y)

    def select_all_clicked(self, sender, event):
        # When "Select all" is checked, set new state based on previous check state, and check or uncheck exam checkboxes

        select_all = list(self.exams_gb.Controls)[0]
        cbs = list(self.exams_gb.Controls)[1:]
        if select_all.Tag == CheckState.Checked:
            select_all.CheckState = select_all.Tag = CheckState.Unchecked
            for cb in cbs:
                cb.Checked = False
        else:
            select_all.CheckState = select_all.Tag = CheckState.Checked
            for cb in cbs:
                cb.Checked = True

    def exam_clicked(self, sender, event):
        # When an exam checkbox is checked, set proper state of "Select all" checkbox 

        select_all = list(self.exams_gb.Controls)[0]
        cbs_cked = [cb.Checked for cb in list(self.exams_gb.Controls)[1:]]
        if all(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Checked
        elif any(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Indeterminate
        else:
            select_all.CheckState = select_all.Tag = CheckState.Unchecked

    def delete_clicked(self, sender, event):
        self.DialogResult = DialogResult.OK


def delete_all_rois_or_geometries(**kwargs):
    """Delete all ROI or geometries in the current case

    Keyword Arguments
    -----------------
    gui: bool
        True if user should choose settings from a GUI, False otherwise
        If False, the other keyword arguments are used to determine settings
        If True, the other keyword arguments are ignored
        Defaults to False
    del_rois: bool
        True if ROIs should be deleted, False if geometries
        Defaults to True
    retain_ext: bool
        True if the External should not be deleted, False otherwise
        Defaults to True
    only_empty: bool
        True if script should only delete ROIs/geometries that are empty on all exams
        Defaults to True
    exam_names: List[str]
        List of exam names from which geometries should be deleted
        If `del_rois` is True, does not apply
        If `del_rois` is False, defaults to all exam names in current case
    
    Alert user if any ROIs/geometries could not be deleted because they are approved
    """

    global case

    # Get current variables
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)  # Exit script with an error

    gui = kwargs.get("gui", False)
    exam_names = []
    if gui:
        form = DeleteAllROIsOrGeometriesForm()
        if form.DialogResult != DialogResult.OK:
            sys.exit()  # User exited GUI

        del_rois = form.rois_rb.Checked
        retain_ext = form.ext_cb.Checked
        only_empty = form.empty_cb.Checked
        if not del_rois:
            exam_names = [cb.Text for cb in form.exams_gb.Controls if cb.Checked][1:]
    else:
        del_rois = kwargs.get("del_rois", True)
        retain_ext = kwargs.get("retain_ext", True)
        only_empty = kwargs.get("only_empty", True)
        if del_rois:
            exam_names = kwargs.get("exam_names", [exam.Name for exam in case.Examinations])
    
    rois = case.PatientModel.RegionsOfInterest
    with CompositeAction("Delete all ROIs/geometries"):
        if retain_ext:  # Do not delete External ROI/geometry
            rois = [roi for roi in rois if roi.Type != "External"]
        if only_empty:  # Only delete ROIs/geometries that are empty on all exams in current case
            rois = [roi for roi in rois if not any(ss.RoiGeometries[roi.Name].HasContours() for ss in case.PatientModel.StructureSets)]
        if del_rois:  # Delete ROIs, not geometries
            roi_names = [roi.Name for roi in rois]
            approved_rois = []
            for roi_name in roi_names:
                try:
                    case.PatientModel.RegionsOfInterest[roi_name].DeleteRoi()
                except:
                    approved_rois.append(roi_name)
            if approved_rois:
                msg = "The following ROIs could not be deleted because they are part of approved structure set(s): {}.".format(", ".join(approved_rois))
                MessageBox.Show(msg, "Delete All ROIs")
        else:  # Delete geometries, not ROIs
            approved_geoms = {}  # exam name : [approved geoms on that exam]
            for exam_name in exam_names:
                geoms = case.PatientModel.StructureSets[exam_name].RoiGeometries
                for roi in rois:
                    try:
                        geoms[roi.Name].DeleteGeometry()
                    except:
                        if exam_name in approved_geoms:
                            approved_geoms[exam_name].append(roi.Name)
                        else:
                            approved_geoms[exam_name] = [roi.Name]
            if approved_geoms:
                msg = "The following geometries are approved so could not be deleted:\n- {}".format("\n- {}".join(["{}: {}".format(exam_name, ", ".join(roi_names)) for exam_name, roi_names in approved_geoms.items()]))
                MessageBox.Show(msg, "Delete All Geometries")