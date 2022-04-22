import clr
import random
import re
import sys
from typing import List, Optional

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

import pandas as pd

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


TG263_PATH = r'T:\Physics\KW\med-phys-spreadsheets\TG-263 Nomenclature with CRMC Colors.xlsm'

class CreateDerived(object):
    """Class that creates derived ROIs based on their TG-263 compliant names from a spreadsheet"""
    
    # Regular expression that matches an ROI name
    # An ROI name matches a given TG-263 name if it is an exact match:
    # -  As is
    # -  With the _L or _R categorization, and/or
    # -  With miscellaneous info. Per TG-263, this is a suffix after a carat (^)
    # -  With a copy number in parentheses
    # Example: The TG-263 name "Lung" matches "Ling", "Lung_L", "Lung_R", "Lung^DJ", "Lung_L^DJ", "Lung (1)", "Lung^DJ (1)", etc.
    ROI_NAME_REGEX = r'^({})(_[LR])?(\^.+)?( \(\d+\))?$'

    TARGET_TYPES = ['CTV', 'GTV', 'PTV']  # ROI types that are targets

    def __init__(self, case: PyScriptObject) -> None:
        """Initializes a CreateDerived object for the given case

        Arguments
        ---------
        case: The case in which to create the derived ROIs
        """
        self._case = case
        self._tg263: pd.DataFrame = self._read_tg263()
        self._to_update: List[str] = []  # List of names of derived ROIs that will be updated after they are created

    def _read_tg263(self) -> pd.DataFrame:
        """Reads in the TG-263 DataFrame from the spreadsheet

        Only columns "TG-263 Primary Name" and "Color" are used
        Column "TG-263 Primary Name" is used as the DataFrame index
        """
        tg263 = pd.read_excel(TG263_PATH, sheet_name='Names & Colors', usecols=['TG-263 Primary Name', 'Color'])
        tg263.rename(columns={'TG-263 Primary Name': 'Name'}, inplace=True)
        tg263.set_index('Name', drop=True, inplace=True)
        return tg263

    def _matching_roi_names(self, name: str) -> List[str]:
        """Selects ROIs whose names match the TG-263 name

        For targets, name is a target type, and all ROIs of that type match
        For non-targets, see constant ROI_NAME_REGEX for details on which ROI names match

        Arguments
        ---------
        name: The TG-263 name to find matching ROIs for

        Returns
        -------
        List of matching ROI names
        """
        rois = case.PatientModel.RegionsOfInterest
        if name in TARGET_TYPES:  # Target
            return [roi.Name for roi in rois if roi.Type.upper() == name]
        elif name == 'ITV':  # RayStation does not have ITV type, so use GTV
            return [roi.Name for roi in rois if roi.Type.upper() == 'GTV']
        elif name.endswith('s') and name[:-1] in self.TARGET_TYPES + ['ITV']:
            targets_to_sum = self._matching_roi_names(name[:-1])
            return self._union_or_intersection(derived_name)
        else:
            return [roi.Name for roi in rois if re.match(self.ROI_NAME_REGEX, roi.Name, re.IGNORECASE) is not None]

    def _create_roi(self, name: str, type_: Optional[str] = 'Control') -> PyScriptObject:
        """Creates an ROI with the given name (made unique) and the color from the spreadsheet

        Arguments
        ---------
        name: Name of the ROI to create
        type_: Type of the ROI to create

        Returns
        -------
        The created ROI
        """
        color = self._tg263_info[name]
        name = self._case.PatientModel.GetUniqueRoiName(name)
        roi = self._case.PatientModel.CreateRoi(Name=name, Type=type_, Color=color)
        return roi

    def _union_or_intersection(self, derived_name, union: Optional[bool] = True):
        """Creates a union or intersection ROI

        Arguments
        ---------
        source_names: List of existing ROI names to union or intersect
        union: True for a union, False for an intersection
        """
        # Get all matching source names
        join_char = '|' if union else '&'
        source_names = derived_name.split(join_char)
        new_source_names = []
        for source_name in source_names:
            all_source_names = self._matching_roi_names(source_name)
            if all_source_names:
                new_source_names.append(all_source_names)
        if not new_source_names:
            return

        # Create union/intersection of all combinations of source names
        combos = itertools.product(new_source_names)
        for combo in combos:
            # Create derived ROI
            derived = self._create_roi(derived_name)
            derived.SetAlgebraExpression(ExpressionA={ 'Operation': 'Union' if union else 'Intersection', 'SourceRoiNames': combo, 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } })
            
            self._to_update.append(derived.Name)  # The new ROI will need updating
        
    def _subtraction(self, derived_name: str):
        """Creates a derived ROI that is one ROI subtracted from another

        Arguments
        ---------
        minuend: Name of the ROI to subtract from
        subtrahend: Name of the ROI to subtract
        """
        # Get all minuends and subtrahends
        minuend, subtrahend = derived_name.split('-')
        minuends = self._matching_roi_names(minuend)
        subtrahends = self._matching_roi_names(subtrahend)
        if not minuends or not subtrahends:
            return

        # Create derived ROI for each combination of minuend and subtrahend
        for minuend in minuends:
            for subtrahend in subtrahends:
                derived_name = minuend + '-' + subtrahend
                derived.SetAlgebraExpression(ExpressionA={ 'Operation': 'Union', 'SourceRoiNames': [minuend], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [subtrahend], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation='Subtraction', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                self._to_update.append(derived.Name)  # The new ROI will need updating

    def _expansion(self, derived_name):
        # Helper function that sets a margin expression for the given organ
        # `exp`: Uniform expansion amount, in mm
        # `source_name`: Name of the ROI to expand
        # `derived_name`: Name of the PRV ROI. If None, new ROI name is source name + 'PRV' + two-digit expansion. 
        #                 Necessary to specify because 16-character limit may cause name to differ fro standard.
        m = re.match(r'(.+)((PRV)|(Ev))(\d{2})', derived_name)
        source_name = m.group(1)
        exp_str = m.group(2)
        exp = int(m.group(3))
        exp_cm = exp / 10
        
        source_names = self._matching_roi_names(source_name)
        if not source_names:
            return
        
        # Create derived ROI for each source name
        for source_name in source_names:
            derived_name = source_name + exp_str + str(exp).zfill(2)
            derived = self._create_roi(derived_name)
            derived.SetMarginExpression(SourceRoiName=source_name, MarginSettings={ 'Type': 'Expand', 'Superior': exp_cm, 'Inferior': exp_cm, 'Anterior': exp_cm, 'Posterior': exp_cm, 'Right': exp_cm, 'Left': exp_cm })
            self._to_update.append(derived.Name)  # The new ROI will need updating

    def create_derived_rois(self):
        for name, color in self._tg263.iterrows():
            if '&' in name:
                self._union_or_intersection(name)
            elif '|' in name:
                self._union_or_intersection(name, False)
            elif 'PRV' in name or 'Ev' in name:
                self._expansion(name)

def add_derived_rois():
    """Add derived ROI geometries from the 'TG-263 Nomenclature with CRMC Colors' spreadsheet

    If source ROI(s) do not exist in the current case, derived ROI(s) that depend on them are not created.

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