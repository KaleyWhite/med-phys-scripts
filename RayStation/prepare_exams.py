import clr
import datetime
import re
import sys
from typing import List, Optional

from connect import *  # Interact w/ RS
from connect.connect_cpython import PyScriptObject

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # For displaying errors and warnings


IMG_SYS = 'HOST-7307'  # Imaging system name for all exams
IMG_FOR_TEMPLATES_REGEX = r'000\d{6} IMAGE FOR TEMPLATES'  # Regular expression to match an exam name that only exists for structure templates
PREPARE_EXAMS_DATE_REGEX = r'(\d{1,2}[/\-\. ]\d{1,2}[/\-\. ](\d{2}|\d{4}))|(\d{6}|\d{8})'


def unique_name(desired_name: str, existing_names: List[str]) -> str:
    """Makes the desired name unique among all names in the list

    Name is made unique with a copy number in parentheses

    Arguments
    ---------
    desired_name: The new name to make unique
    existing_names: List of names among which the new name must be unique

    Returns
    -------
    The new name, made unique

    Example
    -------
    unique_name('hello', ['hello', 'hello (2)']) -> 'hello (1)'
    unique_name('hello', []) -> 'hello'
    """
    new_name = desired_name  # "Base" name to which a copy number may be added
    copy_num = 0  # Assume no copies
    # Increment the copy number until it makes the name unique
    while new_name in existing_names:
        copy_num += 1
        new_name = f'{desired_name} ({copy_num})'
    return new_name


def prepare_exams(study_id: Optional[str] = None) -> Optional[PyScriptObject]:
    """Prepares exams in the given study or with the latest acquisition date

    Sets imaging system, creates 4DCT group (if necessary), and creates AVG/MIP (if necessary)
    Ignores exams that have an approved plan or whose names indicate that they only exist to be included in a structure template
    
    Renames exams:
    - For SBRT exams:
        * Gated exams (exam description includes "Gated"): "Gated __% YYYY-MM-DD", "Gated 0% (Max Inhale) YYYY-MM-DD", or "Gated 50% (Max Exhale) YYYY-MM-DD"
        * Average (exam description includes "AVG"): "AVG (Tx Planning) YYYY-MM-DD"
        * MIP (exam description includes "MIP"): "MIP YYYY-MM-DD"
        * Non-gated (exam description includes "Non-Gated"): "3D YYYY-MM-DD"
    - For other exams, if name does not include a plan or case name:
        * If there are plans on the exam, prepend name with name of first plan on the exam
        * If there are no plans on the exam, prepend name with case name
    Any exam name may include a copy number to prevent duplicates

    A 4DCT group is created from any gated examinations in the given study, that are not already part of a gated group

    If a non-gated exam exists and an average exists or was created, deform from the non-gated to the average and deform POIs

    Arguments
    ----------
    study_id: StudyInstanceUID of the study whose exams to prepare
              If None, use all exams with the latest date
              Defaults to None

    Returns
    -------
    The average exam, if an average exists (or was created), None otherwise
    """
    warnings = ''  # We will display any warnings at the end of the script

    # Get current variables
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()  # Exit script

    # We don't touch exams that have approved plans, because setting imaging system invalidates density values
    approved_planning_exam_names = list(set(plan.GetTotalDoseStructureSet().OnExamination.Name for plan in case.TreatmentPlans if plan.Review is not None and plan.Review.ApprovalStatus == 'Approved'))

    # Select exams in the given study, or latest exams
    if study_id is not None:
        exams = [exam for exam in case.Examinations if exam.Name not in approved_planning_exam_names and re.match(IMG_FOR_TEMPLATES_REGEX, exam.Name) is None and exam.GetAcquisitionDataFromDicom()['StudyModule']['StudyInstanceUID'] == study_id]  # All exam names without approved plans, that are in the study
    # Get all exams with latest date across all exams
    # To start, assume no exams have a date
    else:  
        max_date = None  # Assume no exams have a date
        exams = []  # List of exams with latest date
        for exam in case.Examinations:
            if exam.Name not in approved_planning_exam_names and re.match(IMG_FOR_TEMPLATES_REGEX, exam.Name) is None:  # Ignore exams with approved plans, and structure template exams
                date = exam.GetAcquisitionDataFromDicom()['StudyModule']['StudyDateTime']
                if date is None:  # Exam has no date, so it matches max_date only if max_date is still None
                    if max_date is None:
                        exams.append(exam)
                else:  # Exam has a date
                    date = datetime.date(date.Year, date.Month, date.Day)  # Convert to Python datetime.date to compare to max date
                    if max_date is None or date > max_date:  # We found the first exam with a date, or the exam is later than the current max date
                        max_date = date
                        exams = [exam]  # Clear old list of latest exams and start fresh with current exam
                    elif date == max_date:
                        exams.append(exam)

    # Get gated exam names that are already part of a 4DCT group
    gated_already_in_grp = []
    for grp in case.ExaminationGroups:
        if grp.Type == 'Collection4dct':
            for item in grp.Items:
                if item.Examination.Name not in gated_already_in_grp:
                    gated_already_in_grp.append(item.Examination.Name)

    # Fix imaging system name and add dates to exam names
    gated = []  # List of gated exams to include in new 4DCT group
    date = None
    avg = mip = non_gated = None  # Assume there is no AVG, MIP, or non-gated exam
    for exam in exams:
        old_exam_name = exam.Name
        exam.EquipmentInfo.SetImagingSystemReference(ImagingSystemName=IMG_SYS)  # Correct the imaging system

        dcm = exam.GetAcquisitionDataFromDicom()
        desc = dcm['SeriesModule']['SeriesDescription'] if dcm['SeriesModule']['SeriesDescription'] is not None else dcm['StudyModule']['StudyDescription']  # Exam description is either series description (preferred) or study description
        
        # Format exam date as string to possibly include in exam name
        date = dcm['StudyModule']['StudyDateTime']  
        date = f'{date.Year}-{date.Month:0>2}-{date.Day:0>2}'
        
        series_id = dcm['SeriesModule']['SeriesInstanceUID']
        other_exam_names = [exam_.Name for exam_ in case.Examinations if exam_.GetAcquisitionDataFromDicom()['SeriesModule']['SeriesInstanceUID'] != series_id]  # NAme should be unique among all exams outside the current exam's series

        # Rename exam
        if 'Non-Gated' in desc:
            non_gated = exam
            exam.Name = unique_name('3D ' + date, other_exam_names)
        elif 'AVG' in desc:
            avg = exam
            exam.Name = unique_name('AVG (Tx Planning) ' + date, other_exam_names)
        elif 'MIP' in desc:
            mip = exam
            exam.Name = unique_name('MIP ' + date, other_exam_names)
        elif 'Gated' in desc:  # Description is "Gated, x.0%"
            pct = int(desc[(desc.index(',') + 2):-3])  # e.g., 10
            if pct == 0:
                exam.Name = unique_name(f'Gated {pct}% (Max Inhale) {date}', other_exam_names)
            elif pct == 50:
                exam.Name = unique_name(f'Gated {pct}% (Max Exhale) {date}', other_exam_names)
            else:
                exam.Name = unique_name(f'Gated {pct}% {date}', other_exam_names)
            if old_exam_name not in gated_already_in_grp:
                gated.append(exam.Name)  # Only include exam if it had to be renamed (is new)
        # For non-SBRT exams, add plan or case name to exam name, if necessary
        else:
            names_to_chk = [fr'{plan_.Name}' for plan_ in case.TreatmentPlans if plan_.GetTotalDoseStructureSet().OnExamination.Equals(exam)] + [fr'{case.CaseName}']  # Plan names on the exam, plus case name just in case exam has no plans
            names_to_chk_regex = r'|'.join(names_to_chk)
            if re.search(names_to_chk_regex, exam.Name, re.IGNORECASE) is None:  # No plan or case name in exam name
                exam.Name = unique_name(names_to_chk[0] + exam.Name, other_exam_names)  # Prpend exam name with first plan name, or case name if exam has no plans
        # Add date to exam name, if necessary
        if re.search(PREPARE_EXAMS_DATE_REGEX, exam.Name) is None:  # E.g., 1/26/1999, 03/04/2020, 3/04/2020, 4/5/20, 7-8-21, 8-09-2021, 20200613, 210313
            exam.Name = unique_name(exam.Name + date, other_exam_names)  # Append date to exam name

    # Create gated group from found gated exams
    if gated:  # There are gated exams that are not already part of a gated group
        grp_name = unique_name('4D Phases ' + date, [grp.Name for grp in case.ExaminationGroups])
        case.CreateExaminationGroup(ExaminationGroupName=grp_name, ExaminationGroupType='Collection4dct', ExaminationNames=sorted(gated))  # Sort the exam names to ensure phases are in order
        if avg is None:  # Create AVG, if necessary
            avg_name = unique_name('AVG (Tx Planning) ' + date, [exam_.Name for exam_ in case.Examinations])  # Unique name among all exams
            case.Create4DCTProjection(ExaminationName=avg_name, ExaminationGroupName=grp_name, ProjectionMethod='AverageIntensity')
            avg = case.Examinations[avg_name]
        if mip is None:  # Create MIP, if necessary
            mip_name = unique_name('MIP ' + date, [exam_.Name for exam_ in case.Examinations])  # Unique name among all exams
            case.Create4DCTProjection(ExaminationName=mip_name, ExaminationGroupName=grp_name, ProjectionMethod='MaximumIntensity')

    # Deform from 3D to AVG
    if non_gated is not None and avg is not None:
        # Create external geometry on 3D if it doesn't exist
        try:
            ext = next(roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == 'External')
            ext_geom = case.PatientModel.StructureSets[non_gated.Name].RoiGeometries[ext.Name]
            if not ext_geom.HasContours():
                ext.CreateExternalGeometry(Examination=non_gated)
        except StopIteration:
            warnings += 'There is no external ROI in the current case, so it was not registered with the average.'
        
        # Create deformation in deformation group
        # Create deformable registration group if it doesn't already exist
        deform_grp_name = None
        for grp in case.PatientModel.StructureRegistrationGroups:
            if deform_grp_name is not None:
                break
            for reg in grp.DeformableStructureRegistrations:
                if reg.FromExamination.Equals(non_gated) and reg.ToExamination.Equals(avg):
                    deform_grp_name = grp.Name
                    deform_grp.ComputeHybridDeformableRegistrations(ReferenceExaminationName=non_gated.Name, TargetExaminationNames=[avg.Name], Recompute=True)
                    break
        if deform_grp_name is None:
            deform_grp_name = unique_name(non_gated.Name + ' to ' + avg.Name, [srg.Name for srg in case.PatientModel.StructureRegistrationGroups])
            case.PatientModel.CreateHybridDeformableRegistrationGroup(RegistrationGroupName=deform_grp_name, ReferenceExaminationName=non_gated.Name, TargetExaminationNames=[avg.Name], AlgorithmSettings={ 'NumberOfResolutionLevels': 3, 'InitialResolution': { 'x': 0.5, 'y': 0.5, 'z': 0.5 },'FinalResolution': { 'x': 0.25, 'y': 0.25, 'z': 0.25 }, 'InitialGaussianSmoothingSigma': 2.0, 'FinalGaussianSmoothingSigma': 0.333, 'InitialGridRegularizationWeight': 1500.0, 'FinalGridRegularizationWeight': 400.0, 'ControllingRoiWeight': 0.5, 'ControllingPoiWeight': 0.1, 'MaxNumberOfIterationsPerResolutionLevel': 1000, 'ImageSimilarityMeasure': 'CorrelationCoefficient', 'DeformationStrategy': 'Default', 'ConvergenceTolerance': 1e-5})  # AlgorithmSettings from example in API

        # Map non-empty POI geometries
        poi_names = [poi_geom.OfPoi.Name for poi_geom in case.PatientModel.StructureSets[non_gated.Name].PoiGeometries if poi_geom.Point is not None and abs(poi_geom.Point.x) != float('inf')]
        if poi_names:
            case.MapPoiGeometriesDeformably(PoiGeometryNames=poi_names, CreateNewPois=False, StructureRegistrationGroupNames=[deform_grp_name], ReferenceExaminationNames=[non_gated.Name], TargetExaminationNames=[avg.Name])

    # Display any warnings
    if warnings:
        MessageBox.Show(warnings, 'Warnings')
    
    return avg
