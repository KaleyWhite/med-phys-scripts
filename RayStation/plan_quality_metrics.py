"""This script builds on PaddickAndHomogeneity_10 from RaySearch support"""
import clr
from datetime import datetime
import os
import random
import re
import sys
from typing import Optional

from connect import *
from connect.connect_cpython import PyScriptObject

import pandas as pd

import reportlab.lib.colors
from reportlab.lib.colors import black, Color, dimgray, lightgrey, white
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.platypus.flowables import KeepTogether
from reportlab.platypus.tables import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox


class PlanQualConstants(object):
    """Class that defines several useful constants for this remainder of the script"""

    # ----------------- Clinic-specific constants. CHANGE THESE! ----------------- #

    # Absolute path to the directory in which to create the PDF report
    # This directory does not have to already exist
    OUTPUT_DIR = os.path.join('T:', os.sep, 'Physics', 'Scripts', 'Output Files', 'Plan Quality Metrics')

    # DataFrame of colors from our enhanced TG-263 spreadsheet
    # Indexes ("keys") are the TG-263 names
    TG263_COLORS = pd.read_excel(os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-spreadsheets', 'Structure Names & Colors', 'TG-263 Nomenclature with CRMC Colors.xlsm'), sheet_name='Names & Colors', usecols=['TG-263 Primary Name', 'Color'])
    TG263_COLORS.set_index('TG-263 Primary Name', drop=True, inplace=True)
    TG263_COLORS['Color'] = TG263_COLORS['Color'].apply(lambda x: x[1:-1])  # Remove parens
    
    # --------------------------- No changes necessary --------------------------- #

    STYLES = getSampleStyleSheet()  # Base ReportLab styles (e.g., 'Heading1', 'Normal')

    # Paths to Adobe Reader on RS servers
    ADOBE_READER_PATHS = [os.path.join('C:', os.sep, 'Program Files (x86)', 'Adobe', 'Reader 11.0', 'Reader', 'AcroRd32.exe'), os.path.join('C:', os.sep, 'Program Files (x86)', 'Adobe', 'Acrobat Reader DC', 'Reader', 'AcroRd32.exe')]


def create_dlv_roi(patient: PyScriptObject, case: PyScriptObject, exam: PyScriptObject, target_name: str, idl_roi_name: str, exp_sz: Optional[float] = 2) -> PyScriptObject:
    """Creates a dose-limiting-volume (DLV) ROI that is the intersection of the given target expansion and the given isodose line (IDL) ROI

    The DLV ROI is named "z<target name>_DLV<expansion in mm>"

    Arguments
    ---------
    patient: The patient to whom the ROIs belong
    case: The case to which the ROIs belong
    exam: The examination on which to update the derived ROI geometry
    target_name: The name of the target whose expansion to use in the DLV ROI algebra expression
    idl_roi_name: The name of the IDL (Dose Region) ROI to use in the DLV ROI algebra expression
    exp_sz: The uniform target expansion, in cm, to apply to the target in the DLV ROI algebra expression
            Defaults to 2

    Returns
    -------
    The DLV ROI

    Example
    -------
    create_dlv_roi(some_patient, some_case, some_exam, 'PTV^MD', 'zIDL_100%') -> ROI named 'zPTV^MD_DLV20'
    """
    # Create unique DLV ROI name
    exp_sz_mm = str(int(exp_sz * 10)).zfill(2)  # Convert cm to mm for name to mimic TG-263 convention for PRVs
    dlv_roi = create_roi(patient, case, f'z{target_name}_DLV{exp_sz_mm}', 'DoseRegion')
    
    # DLV = intersection of target expansion, and IDL
    dlv_roi.SetAlgebraExpression(ExpressionA={'Operation': 'Union', 'SourceRoiNames': [target_name],
                                              'MarginSettings': {'Type': 'Expand',
                                                                 'Superior': exp_sz, 'Inferior': exp_sz,
                                                                 'Anterior': exp_sz, 'Posterior': exp_sz,
                                                                 'Right': exp_sz, 'Left': exp_sz}},
                                 ExpressionB={'Operation': 'Union', 'SourceRoiNames': [idl_roi_name],
                                              'MarginSettings': {'Type': 'Expand',
                                                                 'Superior': 0, 'Inferior': 0,
                                                                 'Anterior': 0, 'Posterior': 0,
                                                                 'Right': 0, 'Left': 0}},
                                 ResultOperation='Intersection',
                                 ResultMarginSettings={'Type': 'Expand',
                                                       'Superior': 0, 'Inferior': 0,
                                                       'Anterior': 0, 'Posterior': 0,
                                                       'Right': 0, 'Left': 0})
    
    # Create the geometry on the exam, according to the algebra expression
    dlv_roi.UpdateDerivedGeometry(Examination=exam, Algorithm='Auto')
    
    return dlv_roi


def format_name(name: str) -> str:
    """Converts the patient's Name attribute into a better format for display

    Arguments
    ---------
    pt: The patient whose name to format

    Returns
    -------
    The formatted patient name

    Example
    -------
    Given some_pt with Name '^Jones^Bill^^M':
    format_pt_name(some_pt) -> 'Jones, Bill M'
    """
    parts = [part for part in re.split(r'\^+', name) if part != '']
    name = parts[0]
    if len(parts) > 0:
        name += ', ' + ' '.join(parts[1:])
    return name


def unique_color(case: PyScriptObject) -> str:
    """Generates a new (A, R, G, B) color unique among all ROI colors in the case and the TG-263 colors DataFrame

    Argument
    --------
    The case whose ROI colors the new color must be unique among

    Raises
    ------
    ValueError: If there are 255^4 ROI and TG-263 colors, so that no unique color can be generated

    Returns
    -------
    The unique color, as a string 'A, R, G, B', where each component is between 0 and 255, inclusive
    """
    # Get the colors used so far
    used_colors = [f'{roi.Color.A}, {roi.Color.R}, {roi.Color.G}, {roi.Color.B}' for roi in case.PatientModel.RegionsOfInterest]  # Format System.Colors as 'A, R, G, B'
    used_colors += list(PlanQualConstants.TG263_COLORS['Color'].values)
    used_colors = list(set(used_colors))  # Remove duplicates

    # Ensure there are other colors in the color space
    max_num_colors = 255 ** 4
    if len(used_colors) == max_num_colors:
        raise ValueError(f'There are only {max_num_colors} unique colors in the color space.')

    # Generate 'A, R, G, B' colors until a color is not in the list of used colors
    while True:
        color = f'{random.randint(0, 255)}, {random.randint(0, 255)}, {random.randint(0, 255)}, {random.randint(0, 255)}'
        if color not in used_colors:
            return color


def create_roi(patient: PyScriptObject, case: PyScriptObject, roi_name: str, roi_type: Optional[str] = 'Undefined') -> PyScriptObject:
    """Creates an ROI with the given name, made unique, and the given type

    Makes the ROI invisible

    Arguments
    ---------
    patient: The patient in which to craete the ROI
    case: The case in which to create the ROI
    roi_name: The desired name for the ROI
    roi_type: The ROI type
              Defaults to 'Undefined'

    Returns
    -------
    The new ROI
    """
    unique_roi_name = case.PatientModel.GetUniqueRoiName(DesiredName=roi_name)
    
    # Use color from TG-263 spreadsheet if the name is in the spreadsheet
    # Otherwise, use random unique color
    try:
        color = PlanQualConstants.TG263_COLORS[roi_name]
    except KeyError:
        color = unique_color(case)

    roi = case.PatientModel.CreateRoi(Name=unique_roi_name, Color=color, Type=roi_type)
    patient.SetRoiVisibility(RoiName=unique_roi_name, IsVisible=False)
    return roi


def plan_quality_metrics() -> None:
    """Generates and opens a PDF report of selected plan quality metrics for the current plan

    The following metrics are computed for each beam set in the current plan:
    - D0.035
    - Paddick conformity index (CI) (limits the intersection volume in the denominator, to a 2.2 cm expansion of the target)
    - RTOG (traditional) CI (limits the intersection volume to a 2.2 cm expansion of the target)
    - CTV homogeneity index (HI)
    - Gradient index (GI)
    - % coverage
    - V12Gy
    - V4.5Gy
    - Dmean for Brain

    Dmax, V12, and V4.5 can only be computed if the beam set fractionation is set
    CIs, HI, GI, and % coverage require an Rx
    Brain Dmean obviously requires a Brain geometry
    """
    # Get current variables
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
        sys.exit()

    # Report name
    if not os.path.isdir(PlanQualConstants.OUTPUT_DIR):
        os.makedirs(PlanQualConstants.OUTPUT_DIR)
    pt_name = format_name(patient.Name)
    filename = pt_name + ' ' + plan.Name + ' ' + datetime.now().strftime('%Y-%m-%d %H_%M_%S') + '.pdf'
    filepath = os.path.join(PlanQualConstants.OUTPUT_DIR, re.sub(r'[<>:"/\\\|\?\*]', '_', filename))

    pdf = SimpleDocTemplate(filepath, pagesize=landscape(letter), bottomMargin=0.2 * inch, leftMargin=0.25 * inch, rightMargin=0.2 * inch, topMargin=0.2 * inch)  # 8.5 x 11", 0.2" top and bottom margin, 0.25" left and right margin
    hdg_2 = Paragraph(pt_name + ': MRN ' + patient.PatientID, style=PlanQualConstants.STYLES['Heading2'])
    hdg_1 = Paragraph('Plan Quality Metrics for: ' + plan.Name, style=PlanQualConstants.STYLES['Heading1'])

    metrics_data = [[Paragraph(txt, style=PlanQualConstants.STYLES['Heading3']) for txt in ['Beam set', 'Target', 'D<sub>0.035 cc</sub> [cGy]', 'Paddick CI', 'RTOG CI', 'CTV HI', 'GI', 'Coverage', 'V<sub>12 Gy</sub> [cc]', 'V<sub>4.5 Gy</sub> [cc]', 'Brain D<sub>mean</sub> [cGy]']]]
    metrics_style = [  # Center-align, middle-align, and black outline for all cells. Gray background for header row
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), dimgray),
        ('BACKGROUND', (0, 1), (0, -1), lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
    ]

    # RoIs that are not target specific
    twelve_gy_roi = create_roi(patient, case, 'zIDL_1200', 'DoseRegion')
    fourp5_gy_roi = create_roi(patient, case, 'zIDL_450', 'DoseRegion')

    target_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.OrganData.OrganType == 'Target']

    tbl_row = 1  # Current row in the table (row 0 is the column headers)

    for beam_set in plan.BeamSets:
        beam_set.FractionDose.UpdateDoseGridStructures()  # Just in case brain geometry needs updating. Otherwise all stats for beain (incl. Dmean, which we use, will be zero)
        # ROI geometries on beam set exam
        roi_geoms = beam_set.GetStructureSet().RoiGeometries
        target_names_beam_set = [target_name for target_name in target_names if roi_geoms[target_name].HasContours()]

        # Beam set planning exam
        exam = beam_set.GetStructureSet().OnExamination
        exam_name = exam.Name
        
        # Data to add as row in the ReportLab table
        beam_set_data = [beam_set.DicomPlanLabel]

        rx = beam_set.Prescription.PrimaryPrescriptionDoseReference
        if rx is not None:
            rx = rx.DoseValue
        fractions = beam_set.FractionationPattern
        if fractions is not None:
            fractions = fractions.NumberOfFractions

        # Brain Dmean
        # 'N/A' if no Brain ROI or Brain geometry empty on beam set planning exam
        try:
            brain_geom = roi_geoms['Brain']
        except:  # No Brain ROI
            brain_d_mean = 'N/A'
        else:
            if brain_geom.HasContours():
                brain_d_mean = beam_set.FractionDose.GetDoseStatistic(RoiName='Brain', DoseType='Average')
            else:  # Brain geometry is empty on planning exam
                brain_d_mean = 'N/A'
        if brain_d_mean != 'N/A':
            brain_d_mean = f'{brain_d_mean:.3f}'

        # Can only do Brain Dmean if there is no Rx
        # All other stats depend on Rx
        if fractions is None:
            metrics_style.append(('SPAN', (1, tbl_row), (-2, tbl_row)))
            metrics_data.append([Paragraph(txt, style=PlanQualConstants.STYLES['Normal']) for txt in [beam_set.DicomPlanLabel, 'No fractionation for beam set'] + [''] * 8 + [brain_d_mean]])
            
            tbl_row += 1
        else:
            # Total volume (not volume of an ROI) at 12 Gy and 4.5 Gy, respectively

            twelve_gy_roi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=1200 / fractions)
            twelve_gy_roi_vol = roi_geoms[twelve_gy_roi.Name].GetRoiVolume()
            twelve_gy_roi_vol = f'{twelve_gy_roi_vol:.3f}'

            fourp5_gy_roi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=450 / fractions)
            fourp5_gy_roi_vol = roi_geoms[fourp5_gy_roi.Name].GetRoiVolume()
            fourp5_gy_roi_vol = f'{fourp5_gy_roi_vol:.3f}'

            # Can only do Brain Dmean, V12, and V4.5 if no target geometries
            # All other stats depend on target volume
            if target_names_beam_set:
                # The beam set name and non-target-specific stats should span all target rows for this beam set
                if len(target_names_beam_set) > 1:
                    metrics_style.extend([('SPAN', (i, tbl_row), (i, tbl_row + len(target_names_beam_set) - 1)) for i in [0, -3, -2, -1]])  # Span rows: beam set, V12, V4.5, and brain Dmean columns

                # Compute target volume-dependent stats for each target geometry on the beam set's exam
                for i, target_name in enumerate(target_names_beam_set):
                    abs_target_vol = roi_geoms[target_name].GetRoiVolume()  # Total target volume

                    # Dmax
                    d_max = beam_set.FractionDose.GetDoseAtRelativeVolumes(RoiName=target_name, RelativeVolumes=[0.035 / abs_target_vol])[0]
                    d_max *= fractions
                    d_max = f'{d_max:.2f}'

                    if rx is None:
                        metrics_style.append(('SPAN', (3, tbl_row), (-4, tbl_row)))
                        metrics_data.append([Paragraph(txt, style=PlanQualConstants.STYLES['Normal']) for txt in [beam_set.DicomPlanLabel, target_name, d_max, 'No Rx for beam set'] + [''] * 4 + [twelve_gy_roi_vol, fourp5_gy_roi_vol, brain_d_mean]])
                    else:
                        # Set geometries and get necessary volumes and other stats
                        idl_100_pct_roi = create_roi(patient, case, 'zIDL_100%', 'DoseRegion')
                        idl_100_pct_roi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=rx / fractions)
                        idl_100_pct_roi_vol = roi_geoms[idl_100_pct_roi.Name].GetRoiVolume()

                        idl_50_pct_roi = create_roi(patient, case, 'zIDL_50%', 'DoseRegion')
                        idl_50_pct_roi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=rx * 0.5 / fractions)
                        idl_50_pct_roi_vol = roi_geoms[idl_50_pct_roi.Name].GetRoiVolume()

                        dlv_roi = create_dlv_roi(patient, case, exam, target_name, idl_100_pct_roi.Name, 2.2)  # A bit larger than 2 cm
                        dlv_roi_vol = roi_geoms[dlv_roi.Name].GetRoiVolume()

                        idl_100_pct_and_target_roi = create_roi(patient, case, f'z{idl_100_pct_roi.Name}&{target_name}')
                        idl_100_pct_and_target_roi.SetAlgebraExpression(ExpressionA={'Operation': 'Union', 'SourceRoiNames': [idl_100_pct_roi.Name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': 'Union', 'SourceRoiNames': [target_name], 'MarginSettings': { 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation='Intersection', ResultMarginSettings={ 'Type': 'Expand', 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                        idl_100_pct_and_target_roi.UpdateDerivedGeometry(Examination=exam, Algorithm='Auto')
                        
                        idl_100_pct_and_target_geom = roi_geoms[idl_100_pct_and_target_roi.Name]
                        if idl_100_pct_and_target_geom.HasContours():
                            idl_100_pct_and_target_roi_vol = idl_100_pct_and_target_geom.GetRoiVolume()
                        else:
                            idl_100_pct_and_target_roi_vol = 0

                        # Paddick CI
                        paddick_ci = (idl_100_pct_and_target_roi_vol * idl_100_pct_and_target_roi_vol) / (dlv_roi_vol * abs_target_vol)
                        paddick_ci = f'{paddick_ci:.3f}'

                        # RTOG CI
                        rtog_ci = dlv_roi_vol / abs_target_vol
                        rtog_ci = f'{rtog_ci:.3f}'

                        # CTV HI
                        d_5_pct_vol, d_95_pct_vol = beam_set.FractionDose.GetDoseAtRelativeVolumes(RoiName=target_name, RelativeVolumes=[0.05, 0.95])
                        ctv_hi = (d_5_pct_vol - d_95_pct_vol) / rx
                        ctv_hi = f'{ctv_hi:.3f}'

                        # GI
                        gi = idl_50_pct_roi_vol / idl_100_pct_roi_vol
                        gi = f'{gi:.3f}'

                        # Coverage
                        coverage = beam_set.FractionDose.GetRelativeVolumeAtDoseValues(RoiName=target_name, DoseValues=[rx / fractions])[0]
                        coverage *= 100
                        coverage = f'{coverage:.2f}% at {rx:.0f} cGy'

                        # Delete unnecessary target volume-dependent ROIs
                        for roi in [idl_100_pct_roi, idl_50_pct_roi, dlv_roi, idl_100_pct_and_target_roi]:
                            roi.DeleteRoi()

                        # Add stats to table row data for this beam set
                        beam_set_data = [beam_set.DicomPlanLabel if i == 0 else '', target_name, paddick_ci, rtog_ci, ctv_hi, gi, d_max, coverage]
                        if i == 0:
                            beam_set_data.extend([twelve_gy_roi_vol, fourp5_gy_roi_vol, brain_d_mean]) 
                        else:
                            beam_set_data.extend([''] * 3)
                        metrics_data.append([Paragraph(txt, style=PlanQualConstants.STYLES['Normal']) for txt in beam_set_data])

                    tbl_row += 1

            else:  # No target geometries on beam set exam
                # Display 'No target geometries' message across all target volume-dependent columns in the beam set's row
                metrics_style.append(('SPAN', (1, tbl_row), (-4, tbl_row)))
                msg = 'No target geometries on exam' if target_names else 'No target ROIs in case'
                metrics_data.append([Paragraph(txt, style=PlanQualConstants.STYLES['Normal']) for txt in [beam_set.DicomPlanLabel, msg] + [''] * 6 + [twelve_gy_roi_vol, fourp5_gy_roi_vol, brain_d_mean]])
                
                tbl_row += 1

    # Delete unnecessary Rx-dependent ROIs
    for roi in [twelve_gy_roi, fourp5_gy_roi]:
        roi.DeleteRoi()

    tbl = Table(metrics_data, style=TableStyle(metrics_style))
    elems = [KeepTogether([hdg_2, hdg_1]), tbl]
    pdf.build(elems)

    # Open report
    for reader_path in PlanQualConstants.ADOBE_READER_PATHS:
        try:
            os.system(f'START /B "{reader_path}" "{filepath}"')
            break
        except:
            continue
