import clr
from datetime import datetime
import os
import re
import shutil
import sys
from typing import Dict, List, Optional, Tuple

from connect import *
from connect.connect_cpython import PyScriptObject

import pydicom

clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
from System import InvalidOperationException
from System.Windows.Forms import MessageBox  # For displaying error messages

sys.path.append(os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-scripts' 'RayStation'))
from copy_beam_set import copy_rxs, copy_setup_beams
from copy_opt_stuff import copy_clinical_goals, copy_objectives_and_constraints, copy_opt_params


# Absolute path to the directory where the temporary directory to hold the exported DICOM files will be created
EXPORT_PATH_PARENT = os.path.join('T:', os.sep, 'Physics', 'Temp', 'Copy Exam')


def new_uid(old_uid: str) -> str:
    """Computes a new DICOM UID based on the given UID

    The new UID is the same as the old UID except that the last "part" is incremented by 1

    Arguments
    ---------
    old_uid: The UID to base the new UID off of

    Returns
    -------
    The new UID

    Example
    -------
    new_uid('2.4.6.8') -> '2.4.6.9'
    """
    # Extract last portion of ID, and everything else
    dot_idx = old_uid.rfind('.')
    after_dot = old_uid[(dot_idx + 1):]
    before_dot = old_uid[:dot_idx]
    
    # Construct new ID
    # First part is same. Last part is incremented by 1
    return before_dot + '.' + str(int(after_dot) + 1)


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


def tx_technique(beam_set: PyScriptObject) -> Optional[str]:
    """Determines the treatment technique of the beam set

    Code modified from RS support
    
    Arguments
    ---------
    beam_set: The beam set whose treatment technique to return

    Returns
    -------
    The treatment technique ("SMLC", "VMAT", "DMLC", "Conformal", "Conformal Arc", or "ApplicatorAndCutout"), or None if the treatment technique could not be determined
    """

    if beam_set.Modality == 'Photons':
        if beam_set.PlanGenerationTechnique == 'Imrt':
            if beam_set.DeliveryTechnique == 'SMLC':
                return 'SMLC'
            if beam_set.DeliveryTechnique == 'DynamicArc':
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


def copy_exam(with_plans: Optional[bool] = False) -> None:
    """Copies the current exam, including all structure sets, and possibly all non-imported beam sets with dose

    Note that new exam "Used for" is "Evaluation" regardless of copied exam "Used for".

    New exam name is old exam name plus " - Copy", made unique with a copy number (e.g., "Breast 1/1/21 - Copy (1)")

    Also copies level/window presets

    Arguments
    ---------
    with_plans: True to also copy any beam sets on the old exam to the new exam, False to only copy structure sets
                Defaults to False

    Raises
    ------
    ValueError: If the modality of an exported DICOM file is not CT, RTSTRUCT, RTPLAN, or RTDOSE
    """
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('No patient is open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('No case is open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        old_exam = get_current('Examination')
    except:
        MessageBox.Show('There are no exams in the open case. Click OK to abort the script.', 'No Exams')
        sys.exit()
    
    # Ensure exam is CT
    if old_exam.EquipmentInfo.Modality != 'CT':
        MessageBox.Show('The script can only copy CT exams, not ' + old_exam.EquipmentInfo.Modality + '. Click OK to abort the script.', 'Unsupported Modality')
        sys.exit()
    
    patient_db = get_current('PatientDB')

    old_struct_set = case.PatientModel.StructureSets[old_exam.Name]

    # Existing names lists for calls to `unique_name`
    existing_plan_names = [plan.Name for plan in case.TreatmentPlans]
    existing_exam_names = [exam.Name for exam in case.Examinations]

    # Create export directory
    export_path = os.path.join(EXPORT_PATH_PARENT, datetime.now().strftime('%Y-%m-%d %H_%M_%S'))
    os.makedirs(export_path)

    # Get beam sets to export
    beam_set_ids_to_export = []  # If not copying plans, this list stays empty
    if with_plans:
        for plan in case.TreatmentPlans:
            for beam_set in plan.BeamSets:
                # Only export beam sets that are not imported and that have dose (RayStation throws error otherwise)
                if not beam_set.HasImportedDose() and beam_set.GetPlanningExamination().Equals(old_exam) and beam_set.FractionDose.DoseValues is not None:# and beam_set.Review is not None and beam_set.Review.ApprovalStatus == 'Approved':
                    beam_set_ids_to_export.append(beam_set.BeamSetIdentifier())

    # Export exam and beam sets
    patient.Save()  # Error if you attempt to export when there are unsaved modifications 'RtStructureSetsReferencedFromBeamSets': beam_sets, 'RtStructureSetsForExaminations': [old_exam.Name],
    export_args = {'ExportFolderPath': export_path, 'Examinations': [old_exam.Name], 'BeamSets': beam_set_ids_to_export, 'PhysicalBeamSetDoseForBeamSets': beam_set_ids_to_export, 'PhysicalBeamDosesForBeamSets': beam_set_ids_to_export, 'RtStructureSetsForExaminations': [old_exam.Name], 'RtStructureSetsReferencedFromBeamSets': beam_set_ids_to_export, 'IgnorePreConditionWarnings': False}
    try:
        case.ScriptableDicomExport(**export_args)
    except InvalidOperationException:
        export_args['IgnorePreConditionWarnings'] = True
        case.ScriptableDicomExport(**export_args)  # Retry the export, ignoring warnings

    # Change UIDs in exported DICOM files
    study_uid = None  # All DICOM files have same Study Instance UID (belong to same study)
    exam_series_uid = None  # All CT files have same Series Instance UID (belong to same series)
    plan_names = {}  # Associate old plan names with new plan names (new plan names are old plan names with a copy number to make unique)
    dcm_files = os.listdir(export_path)
    for f in dcm_files:
        abs_f = os.path.join(export_path, f)  # Absolute path to file
        dcm = pydicom.dcmread(abs_f)  # Read DICOM data

        # SOP Instance UID = unique identifier for each DICOM file
        new_sop_uid = new_uid(dcm.SOPInstanceUID)
        new_abs_f = abs_f.replace(dcm.SOPInstanceUID, new_sop_uid)  # We will rename the file because the filename includes the SOP Instance UID
        dcm.SOPInstanceUID = new_sop_uid
        dcm.file_meta.MediaStorageSOPInstanceUID = new_sop_uid  # For some reason, Media Storage SOP Instance UID isn't always changed when SOP Instance UID is

        # Filetype-specific changes
        if dcm.Modality == 'CT':
            # Get new Study Instance UID and CTs' Series Instance UID if they are not yet set
            # First file in export directory is a CT file, so no need to worry about using these in other files while they are still None
            if study_uid is None:
                study_uid = new_uid(dcm.StudyInstanceUID)
                exam_series_uid = new_uid(dcm.SeriesInstanceUID)
            dcm.SeriesInstanceUID = exam_series_uid
        
        elif dcm.Modality == 'RTSTRUCT':
            # Unapprove
            dcm.SeriesDescription = 'RS: Unapproved Structure Set'
            dcm.StructureSetLabel = 'RS: Unapproved'
            dcm.ApprovalStatus = 'UNAPPROVED'
            # Series Instance UID is SOP Instance UID with a short suffix (keep the same suffix)
            dcm.SeriesInstanceUID = dcm.SOPInstanceUID + '.' + dcm.SeriesInstanceUID.split('.')[-1]
            # Fix all references to study and series
            # Assume there is only one referenced FoR, study, and series
            ref_study = dcm.ReferencedFrameOfReferenceSequence[0].RTReferencedStudySequence[0]
            ref_study.ReferencedSOPInstanceUID = study_uid
            ref_series = ref_study.RTReferencedSeriesSequence[0]
            ref_series.SeriesInstanceUID = exam_series_uid
            for c_img in ref_series.ContourImageSequence:
                c_img.ReferencedSOPInstanceUID = new_uid(c_img.ReferencedSOPInstanceUID)
            for contour in dcm.ROIContourSequence:
                for c in contour.ContourSequence:
                    if hasattr(c, 'ContourImageSequence'):  # Contour Geometric Type of POINT does not have this attribute
                        for c_img in c.ContourImageSequence:
                            c_img.ReferencedSOPInstanceUID = new_uid(c_img.ReferencedSOPInstanceUID)
       
        elif dcm.Modality == 'RTPLAN':
            # Unapprove
            dcm.ApprovalStatus = 'UNAPPROVED'
            # Series Instance UID is SOP Instance UID with a short suffix (keep the same suffix)
            dcm.SeriesInstanceUID = dcm.SOPInstanceUID + '.' + dcm.SeriesInstanceUID.split('.')[-1]
            # Set unique plan name
            if dcm.RTPlanName in plan_names:
                dcm.RTPlanName = plan_names[dcm.RTPlanName]
            else:
                # Associate old plan name with new plan name
                plan_names[dcm.RTPlanName] = dcm.RTPlanName = unique_name(dcm.RTPlanName, existing_plan_names, 16)
            # Fix referenced RTSTRUCT
            dcm.ReferencedStructureSetSequence[0].ReferencedSOPInstanceUID = new_uid(dcm.ReferencedStructureSetSequence[0].ReferencedSOPInstanceUID)
        
        elif dcm.Modality == 'RTDOSE':  # Physical dose
            dcm.ReferencedRTPlanSequence[0].ReferencedSOPInstanceUID = new_uid(dcm.ReferencedRTPlanSequence[0].ReferencedSOPInstanceUID)  # Fix referenced RTPLAN
            if dcm.DoseSummationType == 'PLAN':
                # PLAN dose's Series Instance UID is SOP Instance UID with a short suffix (keep the same suffix)
                dcm.SeriesInstanceUID = dcm.SOPInstanceUID + '.' + dcm.SeriesInstanceUID.split('.')[-1]
            else:  # BEAM dose's SEries Instance UID is referenced RTPLAN's SOP Instance UID with a short suffix (keep same suffix)
                dcm.SeriesInstanceUID = dcm.ReferencedRTPlanSequence[0].ReferencedSOPInstanceUID + '.' + dcm.SeriesInstanceUID.split('.')[-1]
        else:  # Perhaps an exam that isn't a CT
            raise ValueError('Unsupported DICOM file modality "' + dcm.Modality + '".')
        
        dcm.StudyInstanceUID = study_uid
        
        # Can't do os.rename due to for loop
        dcm.save_as(new_abs_f, write_like_original=False)  # Overwrite original DICOM file
        os.remove(abs_f)  # Delete old DICOM file

    # Find and import the edited DICOM files
    try:
        study = patient_db.QueryStudiesFromPath(Path=export_path, SearchCriterias={'PatientID': patient.PatientID})[0]  # There is only one study in the directory
    except IndexError:
        MessageBox.Show('Could not find an exported exam with the correct StudyInstanceUID. Examine DICOM in "' + export_path + '" to troubleshoot. The new ID should be ' + study_uid + '. Click OK to abort the script.', 'Series Not Found')
        sys.exit()
    series = patient_db.QuerySeriesFromPath(Path=export_path, SearchCriterias=study)  # Series belonging to the study
    patient.ImportDataFromPath(Path=export_path, SeriesOrInstances=series, CaseName=case.CaseName)  # Import into current case

    # The exam that was just imported
    try:
        new_exam = next(e for e in case.Examinations if e.Series[0].ImportedDicomUID == exam_series_uid)
    except StopIteration:
        MessageBox.Show('Could not find an exported exam with the correct SeriesInstanceUID. Examine DICOM in "' + export_path + '" to troubleshoot. The new ID should be ' + exam_series_uid + '. Click OK to abort the script.', 'Series Not Found')
        sys.exit()
    
    # Rename new exam
    new_exam_name = old_exam.Name + ' - Copy'
    new_exam.Name = unique_name(new_exam_name, existing_exam_names)
    
    # Set new exam imaging system
    old_img_sys = old_exam.EquipmentInfo.ImagingSystemReference.ImagingSystemName
    if old_img_sys is not None:
        new_exam.EquipmentInfo.SetImagingSystemReference(ImagingSystemName=old_img_sys)

    # Update derived geometries
    # This shouldn't change any geometries since the new exam is the same as the old
    # Ignore:
    # - Empty geometries
    # - Underived geometries (obviously)
    # - Dirty derived ROI geometries - these are not updated on old exam
    for geom in old_struct_set.RoiGeometries:
        roi = case.PatientModel.RegionsOfInterest[geom.OfRoi.Name]
        if geom.HasContours() and geom.OfRoi.DerivedRoiExpression is not None and geom.PrimaryShape.DerivedRoiStatus is not None and not geom.PrimaryShape.DerivedRoiStatus.IsShapeDirty:
            roi.UpdateDerivedGeometry(Examination=new_exam)

    # Copy level/window presets
    new_exam.Series[0].LevelWindow = old_exam.Series[0].LevelWindow

    # Merge beam sets into a single plan. E.g., Plan A (1)_1 contains Plan A's second beam set
    # RayStation is weird with the concept of a "beam set". Each beam set is exported to its own RTPLAN file
    if with_plans:
        beam_set_plans: Dict[str, List[PyScriptObject]] = {}  # plan name : plans containing its additional beam sets
        regex = fr'({"|".join(re.escape(plan.Name) for plan in case.TreatmentPlans)})_\d+'
        for plan in case.TreatmentPlans:
            m = re.match(regex, plan.Name)  # Is the plan just a beam set that should be part of another plan?
            if m is not None:
                plan_name_keep = m.group(1)  # The plan that the beam set should be a part of
                if plan_name_keep in beam_set_plans:
                    beam_set_plans[plan_name_keep].append(plan)
                else:
                    beam_set_plans[plan_name_keep] = [plan]
    
        for plan_name_keep, plans_del in beam_set_plans.items():
            plan_keep = case.TreatmentPlans[plan_name_keep]
            plans_del.sort(key=lambda plan: int(plan.Name.split('_')[-1]))
            for plan_del in plans_del:
                old_beam_set = plan_del.BeamSets[0]
                # Add empty beam set
                new_beam_set = plan_keep.AddNewBeamSet(Name=old_beam_set.DicomPlanLabel, ExaminationName=new_exam.Name, MachineName=old_beam_set.MachineReference.MachineName, Modality=old_beam_set.Modality, TreatmentTechnique=tx_technique(old_beam_set), PatientPosition=old_beam_set.PatientPosition, NumberOfFractions=old_beam_set.FractionationPattern.NumberOfFractions, CreateSetupBeams=old_beam_set.PatientSetup.UseSetupBeams, Comment=old_beam_set.Comment)
                # Copy beams from beam set plan
                new_beam_set.CopyBeamsFromBeamSet(BeamSetToCopyFrom=old_beam_set, BeamsToCopy=[beam.Name for beam in old_beam_set.Beams])
                copy_rxs(old_beam_set, new_beam_set)

        # Copy stuff from original plan to copied plan
        for old_plan_name, new_plan_name in plan_names.items():
            old_plan = case.TreatmentPlans[old_plan_name]
            new_plan = case.TreatmentPlans[new_plan_name]
            
            # Comments
            new_comments = '"' + old_plan.Name + '" on "' + new_exam.Name + '"'
            if old_plan.Comments:
                new_comments = old_plan.Comments + '\n' + new_comments
            new_plan.Comments = new_comments
            
            # Clinical goals
            copy_clinical_goals(old_plan, new_plan)
            
            for old_beam_set in old_plan.BeamSets:
                beam_set_name = old_beam_set.DicomPlanLabel
                try:
                    new_beam_set = new_plan.BeamSets[beam_set_name]
                except:  # That beam set was not copied to new plan
                    continue

                # Beam iso names
                for i, old_beam in enumerate(old_beam_set.Beams):
                    new_beam_set.Beams[i].Isocenter.EditIsocenter(Name=old_beam.Isocenter.Annotation.Name)
                
                # Setup beams
                copy_setup_beams(old_beam_set, new_beam_set)

                # Objectives/constraints and optimization parameters
                old_plan_opt = next(opt for opt in old_plan.PlanOptimizations if opt.OptimizedBeamSets[0].DicomPlanLabel == beam_set_name)
                new_plan_opt = next(opt for opt in new_plan.PlanOptimizations if opt.OptimizedBeamSets[0].DicomPlanLabel == beam_set_name)
                copy_objectives_and_constraints(old_plan_opt, new_plan_opt)
                copy_opt_params(old_plan_opt, new_plan_opt)

        # Tell user to delete the unneeded plans
        if beam_set_plans:
            plan_names_del = [plan_del.Name for plans_del in beam_set_plans.values() for plan_del in plans_del]
            MessageBox.Show('Please delete the temp plans: ' + ', '.join('"' + name + '"' for name in plan_names_del) + '.', 'Delete Temporary Plans!')

    # Delete the temporary directory and all its contents
    shutil.rmtree(export_path)