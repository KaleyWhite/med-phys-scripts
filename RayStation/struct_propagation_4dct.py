import clr
import sys
from typing import List, Optional

from connect import *
from connect.connect_cpython import PyScriptObject

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


def unique_name(desired_name: str, existing_names: List[str], max_len: Optional[int] = None) -> str:
    """Makes the desired name unique among all names in the list

    Name is made unique with a copy number in parentheses
    If `max_len` is provided, new name is truncated to be at most `max_len` characters long

    Arguments
    ---------
    desired_name: The new name to make unique
    existing_names: List of names among which the new name must be unique
    max_len: The maximum possible length of the new name
             Defaults to None (no length constraint)

    Returns
    -------
    The new name, made unique

    Examples
    --------
    unique_name('1234567890abcdef', ['hello']) -> '1234567890abcdef'
    unique_name('1234567890abcdef', ['1234567890ab (1)'], 16) -> '1234567890ab (2)'
    """
    new_name = desired_name[:max_len] if max_len is not None else desired_name # Truncate to at most `max_len` characters
    copy_num = 0  # Assume no copies
    # Increment the copy number until it makes the name unique
    while new_name in existing_names:
        copy_num += 1
        copy_str = ' (' + str(copy_num) + ')'  # Suffix to add the name to make it unique
        if max_len is None:
            new_name = desired_name + copy_str
        else:
            name_len = max_len - len(copy_str)  # Number of characters allowed before the suffix
            new_name = desired_name[:name_len] + copy_str
    return new_name


class StructProp4DCTForm(Form):
    """Propagates ROI geometries from a reference image set to all images in the selected 4DCT group.
    
    Creates ITV on all images in the 4DCT group. ITV is union of ITV on gated images in 4DCT group, and target geometry on the reference image set.
    Displays estimated min, mid, and max phases, as well as maximum excursions of the selected target ROI in the gated images.

    User selects 4DCT group, reference image set, ROIs to propagate, and structure by which to determine excursions, from a GUI
    The target is propagated whether or not it is selected for propagation. If this unselected target is derived, its dependent structures are propagated as well

    External, fixation, and support structures are not options to be copied.
    """
    def __init__(self, case: PyScriptObject) -> None:
        """Initializes a StructProp4DCTForm object

        case: The case in which to propagate the structures
        """
        self.case = case

        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # No minimize/maximize controls, and no form resizing
        self.ClientSize = Size(650, 500)
        self.Text = 'Structure Propagation 4DCT'
        self.TopMost = True
        y = 0  # Vertical coordinate of next control

        # ------------------------------ Image Selection ----------------------------- #

        self.groupbox()
        self.gb.Location = Point(0, y)
        self.gb.MinimumSize = Size(650, 0)
        self.gb.Text = 'Image selection'

        # 4DCT group data
        grp_names = [group.Name for group in self.case.ExaminationGroups if group.Type == 'Collection4dct']
        # Exit script with an error if no 4DCT groups exist
        if not grp_names:
            MessageBox.Show('There are no 4DCT groups in the current case.', 'No 4DCT Groups')
            sys.exit()
        
        # 4DCT group Label
        lbl = Label()
        lbl.Location = Point(15, 15)
        lbl.Text = '4DCT group:'
        self.gb.Controls.Add(lbl)

        # 4DCT group ComboBox
        self.grp_names_cb = ComboBox()
        self.grp_names_cb.DropDownStyle = ComboBoxStyle.DropDownList
        self.grp_names_cb.Location = Point(200, 15)
        self.grp_names_cb.Items.AddRange(grp_names)
        self.gb.Controls.Add(self.grp_names_cb)

        # Reference image Label
        lbl = Label()
        lbl.Location = Point(15, 40)
        lbl.Text = 'Reference image:'
        self.gb.Controls.Add(lbl)

        # Reference image ComboBox
        exam_names = [exam.Name for exam in self.case.Examinations]
        self.ref_img_cb = ComboBox()
        self.ref_img_cb.DropDownStyle = ComboBoxStyle.DropDownList
        self.ref_img_cb.Location = Point(200, 40)
        self.ref_img_cb.Items.AddRange(exam_names)
        self.gb.Controls.Add(self.ref_img_cb)
        y += self.gb.Height

        self.Controls.Add(self.gb)

        # Set unfiorm width for all comboboxes = widest 4DCT group name or exam name
        self.grp_names_cb.DropDownWidth = self.grp_names_cb.Width = self.ref_img_cb.DropDownWidth = self.ref_img_cb.Width = max(TextRenderer.MeasureText(text, self.grp_names_cb.Font).Width for text in grp_names + exam_names) + 20

        # -------------------------- Structures to propagate ------------------------- #
        self.structs = self.data_grid()
        self.gb.Location = Point(0, y)
        self.gb.Text = 'Structures to propagate'

        # Target
        self.targets = self.data_grid()
        self.targets.MultiSelect = False
        self.gb.Location = Point(330, y)
        self.gb.Text = 'Center-of-mass ROI'
        y += 250  # Leave room for data grids

        # Results label
        self.result = Label()
        self.result.Location = Point(0, y)
        self.result.AutoSize = True
        self.result.Visible = False
        self.Controls.Add(self.result)

        # Button
        self.run_btn = Button()
        self.run_btn.AutoSize = True
        self.run_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.run_btn.Location = Point(600, 460)
        self.run_btn.Text = 'Run'
        self.Controls.Add(self.run_btn)

        # Event handlers
        self.ref_img_cb.SelectedIndexChanged += self.ref_img_chged
        self.grp_names_cb.SelectedIndex = self.ref_img_cb.SelectedIndex = 0  # Default: select first group and first ref image in lists
        self.run_btn.Click += self.run_clicked

    def groupbox(self):
        self.gb = GroupBox()
        self.gb.AutoSize = True
        self.gb.AutoSizeMode = AutoSizeMode.GrowAndShrink

    def data_grid(self):
        """Adds a DataGrid to the Form
        
        Grid displays ROI names and types for ROIs to select for propagation
        """
        self.groupbox()
        self.gb.MinimumSize = Size(320, 0)
        dg = DataGridView()
        dg.AllowUserToAddRows = dg.AllowUserToDeleteRows = dg.AllowUserToResizeRows = False  # User cannot change rows
        dg.AutoGenerateColumns = False
        dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dg.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        dg.BackgroundColor = Color.White
        dg.ClientSize = Size(self.gb.Size.Width, dg.ColumnHeadersHeight)
        # 2 columns: "Name" and "Type"
        dg.ColumnCount = 2
        dg.Columns[0].Name = 'Name'
        dg.Columns[1].Name = 'Type'
        dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dg.Location = Point(0, 15)
        dg.ReadOnly = True  # User cannot type in the data grid
        dg.RowHeadersVisible = False  # No row numbers
        dg.SelectionMode = DataGridViewSelectionMode.FullRowSelect  # User selects whole row, not individual cell
        self.gb.Controls.Add(dg)
        self.Controls.Add(self.gb)
        return dg

    '''
    def grp_name_chged(self, sender, event):
        self.ref_img_cb.Items.Clear()
        grp_name = self.grp_names_cb.SelectedItem
        ref_imgs = [item.Examination.Name for item in self.case.ExaminationGroups[grp_name].Items]
        self.ref_img_cb.Items.AddRange(ref_imgs)
        self.ref_img_cb.SelectedIndex = 0
        
        text = list(self.grp_names_cb.Items) + list(self.ref_img_cb.Items)
        width = self.grp_names_cb.DropDownWidth = self.ref_img_cb.DropDownWidth = 10 + max([TextRenderer.MeasureText(t, self.grp_names_cb.Font).Width for t in text])
        self.grp_names_cb.ClientSize = Size(width, self.grp_names_cb.ClientSize.Height)
        self.ref_img_cb.ClientSize = Size(width, self.ref_img_cb.ClientSize.Height)
    '''

    def ref_img_chged(self, sender, event):
        """Handles the event that is selecting a reference image from the ComboBox
        
        Changes the lists of possible structures and targets
        Structures are all non-external, non-support, and non-fixation structures with volume >= 0.001 cc on the selected reference exam
        Targets are all targets with volume >= 0.001 cc on the selected reference exam
        """
        # Clear existing structure and target lists
        self.structs.Rows.Clear()
        self.targets.Rows.Clear()

        for geom in self.case.PatientModel.StructureSets[self.ref_img_cb.SelectedItem].RoiGeometries:
            if geom.OfRoi.Type not in ['External', 'Fixation', 'Support'] and geom.HasContours() and geom.GetRoiVolume() >= 0.001:  # Ignore external, support, fixation, empty, or small geometries
                self.structs.Rows.Add([geom.OfRoi.Name, geom.OfRoi.Type])
                if geom.OfRoi.OrganData.OrganType == 'Target':
                    self.targets.Rows.Add([geom.OfRoi.Name, geom.OfRoi.Type])

        # Autosize DataGridView according to column and row heights
        for dg in [self.structs, self.targets]:
            ht = dg.ColumnHeadersHeight + sum(row.Height for row in dg.Rows)
            dg.ClientSize = Size(320, min(ht, 200))

    def set_run_enabled(self):
        """Enables or disables the "Run" button based on data selection
        
        Enables only if at least one structure is selected and a target is selected
        """
        self.run_btn.Enabled = self.structs.SelectedRows.Count > 0 and self.targets.SelectedRows.Count > 0

    def run_clicked(self, sender, event):
        """Handles the event that is clicking the "Run" button

        This is the "main" method. 
        """
        # Extract selected values
        grp_name = self.grp_names_cb.SelectedItem
        ref_img = self.ref_img_cb.SelectedItem
        structs = [row.Cells[0].Value for row in self.structs.SelectedRows]
        target = self.targets.SelectedRows[0].Cells[0].Value

        if target not in structs:  # User did not select the target to propagate
            # If the target is not derived, we can just add it to the structures list
            if not self.case.PatientModel.RegionsOfInterest[target].DerivedRoiExpression:
                structs.append(target)
            # If the target is derived, all of its dependent ROIs must propagate as well
            else:
                structs.extend([geom.OfRoi.Name for geom in self.case.PatientModel.StructureSets[ref_img].RoiGeometries[target].GetDependentRois() if geom.OfRoi.Name not in structs])

        grp_4d = self.case.ExaminationGroups[grp_name]
        target_imgs = [item.Examination.Name for item in grp_4d.Items if item.Examination.Name != ref_img]  # All images in the group except the reference image

        # Get the same DIR algorithms settings as if DIR were run from UI.
        # if the internal structure of the lung is of specific interest (and you don't care about structures outside),
        # it could be considered to change DeformationStrategy to 'InternalLung'
        default_dir_settings = self.case.PatientModel.GetAlgorithmSettingsForHybridDIR(ReferenceExaminationName=ref_img, TargetExaminationName=target_imgs[0], FinalResolution={'x': 0.25, 'y': 0.25, 'z': 0.25}, DiscardImageInformation=False, UsesControllingROIs=False, DeformationStrategy='Default')
        
        # Create deformable registration
        dir_grp_name = unique_name('DIR for ROI Propagation', [srg.Name for srg in self.case.PatientModel.StructureRegistrationGroups])
        self.case.PatientModel.CreateHybridDeformableRegistrationGroup(RegistrationGroupName=dir_grp_name, ReferenceExaminationName=ref_img, TargetExaminationNames=target_imgs, AlgorithmSettings=default_dir_settings)

        # Map ROI geometries from reference to all images in group
        # Ignore ROIs that already have contours on the given target image
        for struct in structs:
            # Deform geometry to all target images w/o geometry for this ROI
            imgs = [img for img in target_imgs if img != ref_img and not self.case.PatientModel.StructureSets[img].RoiGeometries[struct].HasContours()]  # Exam names that don't have this ROI contoured
            if imgs:
                self.case.MapRoiGeometriesDeformably(RoiGeometryNames=[struct], StructureRegistrationGroupNames=[dir_grp_name] * len(imgs), ReferenceExaminationNames=[ref_img] * len(imgs), TargetExaminationNames=imgs)  # Map the geometry from the refernce to the taregets
            
            # If ROI is a target, create ITV
            roi = self.case.PatientModel.RegionsOfInterest[struct]
            if roi.OrganData.OrganType == 'Target':
                gated_imgs = [img for img in target_imgs if self.case.Examinations[img].GetAcquisitionDataFromDicom()['SeriesModule']['SeriesDescription'] is not None and 'Gated' in self.case.Examinations[img].GetAcquisitionDataFromDicom()['SeriesModule']['SeriesDescription']]  # Only create ITV from gated images
                non_gated_imgs = [img for img in target_imgs if img not in gated_imgs]
                
                # Create 'real' and temporary ITV. Temp ITV is from geometries on all phases in group. We will set the geometry for the 'real' ITV later by copying the temp ITV geometry into it and then underiving it. The ITV cannot depend on itself
                real_itv_name = 'i{}'.format(roi.Type.upper())
                real_itv_name = self.case.PatientModel.GetUniqueRoiName(DesiredName=real_itv_name)
                real_itv = self.case.PatientModel.CreateRoi(Name=real_itv_name, Color=roi.Color, Type=roi.Type)
                
                itv_name = '{}^Temp'.format(real_itv_name)
                itv_name = self.case.PatientModel.GetUniqueRoiName(DesiredName=itv_name)
                itv = self.case.PatientModel.CreateRoi(Name=itv_name, Color=roi.Color, Type=roi.Type)
                itv.CreateITV(SourceRegionOfInterest=roi, ExaminationNames=gated_imgs, MarginSettingsData={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })  # ITV from geometries on gated images
                
                # Copy ITV to each non-gated phase in group
                if non_gated_imgs:
                    self.PatientModel.CopyRoiGeometries(SourceExamination=self.case.Examinations[gated_imgs[0]], TargetExaminationNames=non_gated_imgs, RoiNames=[struct])  # Doesn't matter which gated exam we copy from b/c geometry is the same on all
                
                # Union each ITV geometry with ITV geometry on the reference exam

                # Create copy of target on reference exam
                copied_target_name = 'Copy of {}'.format(struct)
                copied_target_name = self.case.PatientModel.GetUniqueRoiName(DesiredName=copied_target_name)
                copied_target = self.case.PatientModel.CreateRoi(Name=copied_target_name)  # ROI that is copy of the target ROI
                
                # Copied target's geometry is same as target's
                copied_target.SetAlgebraExpression(ExpressionA={ 'Operation': 'Union', 'SourceRoiNames': [struct], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation='None', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                copied_target.UpdateDerivedGeometry(Examination=self.case.Examinations[ref_img])

                # Copy the copied target to other phases in group
                self.case.PatientModel.CopyRoiGeometries(SourceExamination=self.case.Examinations[ref_img], TargetExaminationNames=[img for img in gated_imgs if img != ref_img], RoiNames=[copied_target_name])

                # Union the copied target with the ITV geometry on each phase
                for img in gated_imgs:
                    if img != ref_img:
                        real_itv.SetAlgebraExpression(ExpressionA={ 'Operation': 'Union', 'SourceRoiNames': [itv.Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [copied_target_name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation='Union', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                        real_itv.UpdateDerivedGeometry(Examination=self.case.Examinations[img])

                # Underived 'real' ITV since we're deleting the ROI it depends on
                real_itv.DeleteExpression()

                # Delete the temp ITV and the copied target
                itv.DeleteRoi()
                copied_target.DeleteRoi()

        # Update derived geometries
        for target_img in target_imgs:
            geoms = self.case.PatientModel.StructureSets[target_img].RoiGeometries
            # Add external geometry if necessary
            ext = [geom for geom in geoms if geom.OfRoi.Type == 'External']
            if ext:
                ext = ext[0]
            else:
                ext = self.case.PatientModel.CreateRoi(Name='External', Color='white', Type='External')
                ext.CreateExternalGeometry(Examination=qact)
            derived_rois = [roi.Name for roi in self.case.PatientModel.RegionsOfInterest if roi.DerivedRoiExpression and all(geoms[dep_roi].HasContours() for dep_roi in geoms[roi.Name].GetDependentRois())]  # Derived ROIs with contours for all dependent ROIs
            if derived_rois:
                self.case.PatientModel.UpdateDerivedGeometries(RoiNames=derived_rois, Examination=self.case.Examinations[target_img])

        # Transverse coordinate of target center-of-mass in all target images
        ctrs_of_mass = [self.case.PatientModel.StructureSets[img].RoiGeometries[target].GetCenterOfRoi() for img in list(set([ref_img] + target_imgs))]
        ctrs_of_mass_x = [ctr.x for ctr in ctrs_of_mass]
        ctrs_of_mass_y = [ctr.y for ctr in ctrs_of_mass]
        ctrs_of_mass_z = [ctr.z for ctr in ctrs_of_mass]
        max_idx = ctrs_of_mass_z.index(max(ctrs_of_mass_z))
        min_idx = ctrs_of_mass_z.index(min(ctrs_of_mass_z))
        mid_idx = ctrs_of_mass_z.index(sorted(ctrs_of_mass_z)[len(ctrs_of_mass_z) / 2])
        text = 'Phases:\n'
        text = f'    Max: "{grp_4d.Items[max_idx].Examination.Name}".\n'
        text += f'    Min: "{grp_4d.Items[min_idx].Examination.Name}".\n'
        text += f'    Mid: "{grp_4d.Items[mid_idx].Examination.Name}".\n'
        text += '\nMax excursion:\n'
        text += f'    R-L: {max(abs(ctr - ctr_2) for ctr in ctrs_of_mass_x for ctr_2 in ctrs_of_mass_x):.2f} cm\n'
        text += f'    I-S: {max(abs(ctr - ctr_2) for ctr in ctrs_of_mass_y for ctr_2 in ctrs_of_mass_y):.2f} cm\n'
        text += f'    P-A: {max(abs(ctr - ctr_2) for ctr in ctrs_of_mass_z for ctr_2 in ctrs_of_mass_z):.2f} cm\n'
        self.result.Text = text
        self.result.Visible = True


def struct_propagation_4dct() -> None:
    """Propagates the selected ROI geometries, including a target geometry, across all gated exams in the selected 4DCT group"""
    # Get current objects
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()

    StructProp4DCTForm(case).ShowDialog()
