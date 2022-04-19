import clr
import random
import re
import sys

from connect import *
import pandas as pd

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


TG263_PATH = r'T:\Physics\KW\med-phys-spreadsheets\TG-263 Nomenclature with CRMC Colors.xlsm'

class CreateDerived(object):
    ROI_NAME_REGEX = r'^({})(_[LR])?(\^.+)?( \(\d+\))?$'

    def __init__(self, case):
        self._case = case
        self._tg263 = self._read_tg263()
        self._to_update = []

    def _populate_target_names(self):
        target_names = {}
        for target_type in ('CTV', 'GTV', 'PTV'):
            target_names[target_type] = [roi.Name for roi in self._case.PatientModel.RegionsOfInterest if roi.Type.upper() == target_type]
        return target_names

    def _read_tg263(self):
        tg263 = pd.read_excel(TG263_PATH, sheet_name='Names & Colors', usecols=['TG-263 Primary Name', 'Color'])
        tg263.rename(columns={'TG-263 Primary Name': 'Name'}, inplace=True)
        tg263.set_index('Name', drop=True, inplace=True)
        return tg263

    def _matching_roi_names(self, name):
        rois = case.PatientModel.RegionsOfInterest
        if name in ('CTV', 'GTV', 'PTV'):
            return [roi.Name for roi in rois if roi.Type.upper() == name]
        

    def _unique_roi_name(self, name):
        all_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
        copy_num = 1
        while name in all_names:
            name = fr'{name} ({copy_num})'
            copy_num += 1
        return name

    def _union_or_intersection(self, source_names, **kwargs):
        union = kwargs.get('union', True)
        derived_name = kwargs.get('derived_name')
        derived_type = kwargs.get('derived_type', 'Control')

        source_names.sort(key=lambda x: x.lower())

        if derived_name is None:
            join_char = '_' if union else '&'
            derived_name = join_char.join(source_names)

        color = self._tg263_info[derived_name]
        derived_name = self._unique_roi_name(derived_name)
        derived = self._case.PatientModel.CreateRoi(Name=derived_name, Type=derived_type, Color=color)
        derived.SetAlgebraExpression(ExpressionA={ 'Operation': 'Union' if union else 'Intersection', 'SourceRoiNames': source_names, 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } })
        
        self._to_update.append(derived_name)
        
    def subtraction(self, source_names, **kwargs):
        # Helper function that sets an ROI algebra expression for organs minus the given target type
        # `target_type`: 'CTV', 'GTV', or 'PTV'
        # `organ_names`: List of names of ROIs from which to subtract the target
        
        # Create each derived ROI, including left and right sides
        subtrahend_regex = ROI_NAME_REGEX.format(source_names[0])
        subtrahends = [roi.Name for roi in self._case.PatientModel.RegionsOfInterest if re.match(subtrahend_regex, roi.Name, re.IGNORECASE) is not None]

        for source_name in source_names
        minuend_regex = ROI_NAME_REGEX.format(r'\(' + '|'.join(source_names[]))
        minuends = [roi.Name for roi in self._case.PatientModel.RegionsOfInterest if re.match(subtrahend_regex, roi.Name, re.IGNORECASE) is not None]
        for roi in case.PatientModel.RegionsOfInterest:
            if not re.match(regex, roi.Name, re.IGNORECASE):
                continue
            for target_name in self.target_names[target_type]:
                derived_name = roi.Name + '-' + target_name
                color = self._tg263_info[derived_name]
                derived_name = self._unique_roi_name(derived_name)
                derived = self._case.PatientModel.CreateRoi(Name=derived_name, Type='Organ', Color=color)
                derived.SetAlgebraExpression(ExpressionA={ 'Operation': 'Union', 'SourceRoiNames': [roi.Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [target_name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation='Subtraction', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                
                self._to_update.append(derived_name)

    def union_or_intersection(self, source_names, **kwargs):
        # Helper function that sets an ROI algebra expression for the union or intersection of the given ROIs
        # `source_names`: List of names of ROIs to be unioned or intersected
        # `union`: True if the source ROIs should be unioned, False if they should be intersected

        new_source_names = np.array(self._target_names[source_name] for source_name in source_names if source_name in self._target_names).flatten()
        for source_name in source_names:
            if source_name in self._target_names:
                new_source_names.append(self._target_names[source_name])
            else:
                regex = ROI_NAME_REGEX.format(source_name)
                new_source_names.append([roi.Name for roi in self._case.PatientModel.RegionsOfInterest if re.match(regex, roi.Name, re.IGNORECASE) is not None])

        combos = list(set(itertools.product(new_source_names)))
        for combo in combos:
            self._union_or_intersection(combo, **kwargs)

    def planning_region_volume(exp, source_name):
        # Helper function that sets a margin expression for the given organ
        # `exp`: Uniform expansion amount, in mm
        # `source_name`: Name of the ROI to expand
        # `derived_name`: Name of the PRV ROI. If None, new ROI name is source name + 'PRV' + two-digit expansion. 
        #                 Necessary to specify because 16-character limit may cause name to differ fro standard.

        derived_name =self._unique_roi_name(f'{source_name}_PRV{str(exp).zfill(2)}')
        exp /= 10  # mm -> cm

        derived.SetMarginExpression(SourceRoiName=source.Name, MarginSettings={ 'Type': 'Expand', 'Superior': exp, 'Inferior': exp, 'Anterior': exp, 'Posterior': exp, 'Right': exp, 'Left': exp })
            to_update.append(derived.Name)



def add_derived_rois():
    """Add derived ROI geometries from the 'Clinical Goals' spreadsheet

    Types of derived geometries created:
    - Bilateral organ sums (e.g., 'Lungs' = 'Lung_L' + 'Lung_R')
    - Other, misc. sums (e.g., 'Jejunum_Ileum' = 'Jejunum' + 'Ileum')
    - Overlaps (intersections) (e.g., 'BneMdbl&JntTM&PTV' = 'Bone_Mandible' & 'Joint_TM' & PTV)
    - Target exclusions (subtracting a target volume from an organ) (e.g., 'Brain-PTV' = 'Brain' - PTV)
    - Planning regions volumes (PRVs) (e.g., 'Brainstem_PRV03' = 3mm expansion of 'Brainstem')
    - 'E-PTV_Ev20' (SBRT plans only)

    If source ROI(s) do not exist in the current case, derived ROI(s) that depend on them are not created.
    
    Source ROIs are the ROIs with the 'latest' name.
    Example:
    We are creating 'Lungs' from 'Lung_L' and 'Lung_R'. ROIs 'Lung_L' and 'Lung_L (1)' exist. 'Lung_L (1)' has the highest copy number, so it is used.

    Derived geometries are created for the latest unapproved ROIs with the desired derived name. If none such ROIs exist, a new ROI is created.
    Example:
    We are creating 'Lungs'. ROIs 'Lungs' and 'Lungs (1)' exist, but 'Lungs (1)' is approved, so 'Lungs' is used. If both 'Lungs' and 'Lungs (1)' were approved, we would create a new ROI 'Lungs (2)'.

    Assumptions
    -----------
    All ROI names (perhaps excluding targets) are TG-263 compliant
    """

    # Get current objects
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort script.', 'No Open Case')
        sys.exit(1)
    try:
        exam = get_current('Examination')
    except:
        MessageBox.Show('There are no exams in the current case. Click OK to abort script.', 'No Exams')
        sys.exit(1)

    to_update = []

    

    
    for _, row in tg263.iterrows():
        # Intersection?
        if '&' in row.name:
            sources = name.split('&')
            source_types = []
            targets = []
            for source in sources:
                source_row = tg263.loc[tg263.name == source]
                if source_row.type == 'Target':
                    source_type = source_row.major_cat[0] + source_row.major_cat[1:].lower()
                    if source_type not in source_types:
                        source_types.append(source_type)
                elif source_row.type == 'Anatomic':
                    if 'Organ' not in source_types:
                        source_types.append('Organ')
            derived_type = source_types[0] if len(source_types) == 1 else 'Control'
            name = unique_roi_name(case, name)
            derived = case.PatientModel.CreateRoi(Name=name, Type=derived_type, Color=color)
            derived.SetAlgebraExpression(ExpressionA={ 'Operation': 'Intersection', 'SourceRoiNames': sources, 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } })
            self._to_update.append(derived_name)
            self._all_roi_names.append(derived_name)

    # Sums
    l_r_sums = ['Kidney', 'Lung', 'Parotid']
    for l_r_sum in l_r_sums:
        create_derived.union_or_intersection([l_r_sum + '_L', l_r_sum + '_R'], derived_name=lr_sum + 's')

    other_sums = ['Bone_Mandible', 'Joint_TM'],
        'Jejunum_Ileum': ['Jejunum', 'Ileum']
    }
    for sum_roi_name, base_roi_names in other_sums.items():
        union_or_intersection(sum_roi_name, base_roi_names)

    # Overlaps
    overlaps = {
        'BnMdbl&JntTM&PTV': ['Bone_Mandible', 'Joint_TM', 'PTV']
    }
    for overlap_name, base_roi_names in overlaps.items():
        union_or_intersection(overlap_name, base_roi_names, union=False)

    # OAR expansions
    prvs = [
        (1, 'Brainstem'),
        (1, 'Lobe_Temporal', 'LbeTemporalPRV1'),
        (1, 'OpticNrv'),
        (1, 'Tongue_All'),
        (3, 'Brainstem'),
        (5, 'CaudaEquina'),
        (3, 'OpticChiasm', 'OpticChiasm_PRV3'),
        (5, 'Glnd_LcrimalPRV5'),
        (5, 'Skin'),
        (5, 'SpinalCord'),
        (7, 'Lens')
    ]
    for prv in prvs:
        planning_region_volume(*prv)

    # Target exclusion
    target_exclude = {
        'CTV': ['Lungs'],
        'GTV': ['Liver', 'Lungs'],
        'PTV': ['Brain', 'Cavity_Oral', 'Chestwall', 'Glottis', 'Larynx', 'Liver', 'Lungs', 'Stomach']
    }
    for target, base_names in target_exclude.items():
        organ_minus_target(target, base_names)

    # E-PTV_Ev20
    try:
        plan = get_current('Plan')
        if plan.BeamSets.Count > 0:
            if any(get_tx_technique(bs) == 'SBRT' for bs in plan.BeamSets):
                ext = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type == 'External']
                if ext:
                    set_target('PTV')
                    ptv = targets['PTV']
                    if ptv is not None:
                        nt = create_roi_if_absent('E-PTV_Ev20', 'Control')
                        nt.SetAlgebraExpression(ExpressionA={ 'Operation': 'Union', 'SourceRoiNames': [ext[0].Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [ptv.Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0.2, 'Inferior': 0.2, 'Anterior': 0.2, 'Posterior': 0.2, 'Right': 0.2, 'Left': 0.2 } }, ResultOperation='Subtraction', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                        to_update.append(nt.Name)
    except:
        pass

    # Update derived geometries if any were added
    if to_update:
        case.PatientModel.UpdateDerivedGeometries(RoiNames=to_update, Examination=exam)


def add_derived_rois():
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()