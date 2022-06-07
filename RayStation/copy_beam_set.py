import clr
import re
import sys
from typing import List, Optional, Tuple

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons

sys.path.append(r'T:\Physics\KW\med-phys-scripts\RayStation')
from copy_opt_stuff import copy_objectives_and_constraints, copy_opt_params
from copy_plan_without_changes import copy_plan_without_changes


def copy_rxs(old_beam_set: PyScriptObject, new_beam_set: PyScriptObject) -> None:
    """Copies prescriptions from one beam set to another

    Arguments
    ---------
    old_beam_set: The beam set to copy prescriptions from
    new_beam_Set: The unapproved beam set to copy prescriptions to

    Raises
    ------
    ValueError: If the new beam set is approved
    """
    if new_beam_set.Review is not None and new_beam_set.Review.ApprovalStatus == 'Approved':
        raise ValueError('Cannot copy Rx\'s to an approved beam set')
    if old_beam_set.Prescription is None:
        return
    with CompositeAction('Copy Prescriptions'):
        for old_rx in old_beam_set.Prescription.PrescriptionDoseReferences:  # Copy all Rx's, not just the primary Rx
            if hasattr(old_rx, 'OnStructure'):
                if old_rx.PrescriptionType == 'DoseAtPoint':  # Rx to POI
                    new_beam_set.AddPoiPrescriptionDoseReference(PoiName=old_rx.OnStructure.Name, DoseValue=old_rx.DoseValue, RelativePrescriptionLevel=old_rx.RelativePrescriptionLevel)
                else:  # Rx to ROI
                    new_beam_set.AddRoiPrescriptionDoseReference(RoiName=old_rx.OnStructure.Name, DoseVolume=old_rx.DoseVolume, PrescriptionType=old_rx.PrescriptionType, DoseValue=old_rx.DoseValue, RelativePrescriptionLevel=old_rx.RelativePrescriptionLevel)
            elif old_rx.OnDoseSpecificationPoint is None:  # Rx to DSP
                new_beam_set.AddSitePrescriptionDoseReference(Description=old_rx.Description, DoseValue=old_rx.DoseValue, RelativePrescriptionLevel=old_rx.RelativePrescriptionLevel)
            else:  # Rx to site that is not a DSP
                new_beam_set.AddSitePrescriptionDoseReference(Description=old_rx.Description, NameOfDoseSpecificationPoint=old_rx.OnDoseSpecificationPoint.Name, DoseValue=old_rx.DoseValue, RelativePrescriptionLevel=old_rx.RelativePrescriptionLevel)


def copy_setup_beams(old_beam_set: PyScriptObject, new_beam_set: PyScriptObject, beam_num: Optional[int] = None, existing_beam_names: Optional[List[str]] = []) -> Optional[int]:
    """Copies setup beams from one beam set to another

    If old beam set is not set to use setup beams, does nothing

    Clears any existing setup beams in the new beam set

    Arguments
    ---------
    old_beam_set: The beam set to copy setup beams from
    new_beam_set: The unapproved beam set to copy setup beam to
    beam_num: The beam number to start numbering the new setup beams with
              Necessary is new beam names should be unique (e.g., in the current case). Otherwise, new setup beams have same numbers as old
              Defaults to None (new setup beam numbers need not be unique)
    existing_beam_names: List of beam names that the new setup beam names must be unique among
                         Necessary is new setup beam names must be unique (e.g., in the current case). Otherwise, new setup beam names are same as old
                         If provided, is modified in-place with the new setup beam names
                         Defaults to the empty list (setup beam names need not be unique)

    Raises
    ------
    ValueError: If the new beam set is approved

    Returns
    -------
    A unique beam number. The beam number after all setup beams have been consecutively renumbered
    If beam_num is None, returns None
    """
    if new_beam_set.Review is not None and new_beam_set.Review.ApprovalStatus == 'Approved':
        raise ValueError('Cannot copy setup beams to an approved beam set')
    use_setup_beams = old_beam_set.PatientSetup.UseSetupBeams
    new_beam_set.PatientSetup.UseSetupBeams = use_setup_beams
    if not use_setup_beams or old_beam_set.PatientSetup.SetupBeams.Count == 0:
        return
    old_setup_beams = old_beam_set.PatientSetup.SetupBeams
    new_beam_set.UpdateSetupBeams(ResetSetupBeams=True, SetupBeamsGantryAngles=[setup_beam.GantryAngle for setup_beam in old_setup_beams])
    for i, old_setup_beam in enumerate(old_setup_beams):
        new_setup_beam = new_beam_set.PatientSetup.SetupBeams[i]
        new_setup_beam.Description = old_setup_beam.Description
        # New setup beam's number
        if beam_num is None:
            new_setup_beam.Number = old_setup_beam.Number
        else:
            new_setup_beam.Number = beam_num
            beam_num += 1
        # New setup beam's name
        if existing_beam_names:  # Same name as old but made unique
            new_setup_beam.Name = unique_name(old_setup_beam.Name, existing_beam_names)
            existing_beam_names.append(new_setup_beam.Name)
        else:  # Same name as old
            new_setup_beam.Name = old_setup_beam.Name
    return beam_num


def is_sabr(beam_set: PyScriptObject) -> bool:
    """Returns whether or not a VMAT beam set is SABR.

    SABR includes SBRT, SRT, and SRS. A SABR beam set is a VMAT beam set with no more than 15 fractions and at least 600 cGy fractional dose.

    Arguments
    ---------
    beam_set: The beam set to check whether it is SABR
              Assumes that the beam set is VMAT!
    """
    fx = rx_val = None
    rx = beam_set.Prescription
    if rx is None:
        return False
    rx = rx.PrimaryPrescriptionDoseReference
    if rx is None:
        return False
    fx = beam_set.FractionationPattern
    if fx is None:
        return False
    fx = fx.NumberOfFractions
    rx = rx.DoseValue
    return fx <= 15 and rx / fx >= 600


def get_tx_technique(beam_set: PyScriptObject) -> Optional[str]:
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


def unique_name(desired_name: str, existing_names: List[str]) -> str:
    """Makes the desired name unique among all names in the list

    Name is made unique with a copy number in parentheses
    New name is truncated to be at most 16 characters long

    Arguments
    ---------
    desired_name: The new name to make unique
    existing_names: List of names among which the new name must be unique

    Returns
    -------
    The new name, made unique

    Example
    -------
    unique_name('1234567890abcdef', ['hello']) -> '1234567890abcdef'
    unique_name('1234567890abcdef', ['1234567890ab (1)']) -> '1234567890ab (2)'
    """
    new_name = desired_name[:16]  # Truncate to at most 16 characters
    copy_num = 0  # Assume no copies
    # Increment the copy number until it makes the name unique
    while new_name in existing_names:
        copy_num += 1
        copy_str = ' (' + str(copy_num) + ')'  # Suffix to add the name to make it unique
        name_len = 16 - len(copy_str)  # Number of characters allowed before the suffix
        new_name = desired_name[:name_len] + copy_str
    return new_name


def unique_iso_name(desired_name, beam_set, plan):
    iso_names = []
    for bs in plan.BeamSets:
        if not bs.Equals(beam_set):
            for beam in bs.Beams:
                iso_name = beam.Isocenter.Annotation.Name
                if iso_name not in iso_names:
                    iso_names.append(iso_name)
    return unique_name(desired_name, iso_names)


def names_nums(patient: PyScriptObject) -> Tuple[List[str], List[str], int]:
    """Returns a list of beam set names, a list of beam names (including setup beams), and the next consecutive unique beam number in the patient

    Arguments
    ---------
    patient: The patient containing the beam sets

    Example
    -------
    names_nums(some_patient) -> (['Beam set 1', 'Another beam set'], ['1', '2', 'SB_1', 'SB_2', 'AP', Rt Lat'], 3)
    """
    beam_set_names, beam_names = [], []
    beam_num = 1
    for case in patient.Cases:
        for plan in case.TreatmentPlans:
            for beam_set in plan.BeamSets:
                beam_set_names.append(beam_set.DicomPlanLabel)
                for beam in beam_set.Beams:
                    beam_names.append(beam.Name)
                    beam_num = max(beam_num, beam.Number + 1)
                for setup_beam in beam_set.PatientSetup.SetupBeams:
                    beam_names.append(setup_beam.Name)
                    beam_num = max(beam_num, setup_beam.Number + 1)
    return beam_set_names, beam_names, beam_num


def copy_beam_set() -> None:
    """Copies the current beam set to a new beam set in the same plan
    
    Copy electron or photon (including VMAT) beam sets:
        - Treatment and setup beams
            * Unique beam numbers across all cases in current patient
            * Beam names are same as old, made unique with a copy number
        - Select optimization settings
        - Prescriptions

    Unfortunately, dose cannot accurately be copied for VMAT
    Does not optimize or compute dose

    Code is modified from RS support's CopyBeamSet script
    """
    # Get current variables
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('No patient is open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There is no plan open. Click OK to abort the script.', 'No Open Plan')
        sys.exit()
    try:
        old_beam_set = get_current('BeamSet')
    except:
        MessageBox.Show('There are no beam sets in the current plan. Click OK to abort the script.', 'No Beam Sets')
        sys.exit()

    # Exit script if modality is neither electrons nor photons
    if old_beam_set.Modality not in ['Electrons', 'Photons']:
        MessageBox.Show(old_beam_set.Modality + ' is/are not supported. Click OK to abort the script.', 'Unsupported Modality')
        sys.exit()

    # Offer to copy plan if it is approved
    if plan.Review is not None and plan.Review.ApprovalStatus == 'Approved':
        res = MessageBox.Show('The plan is approved, so a beam set cannot be added. Would you like to create a copy of the plan, and then copy the beam set in the new plan?\nClick "No" to abort the script.', MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit()
        plan = copy_plan_without_changes()

    # We will display any warnings at the end of the script
    warnings = ''

    # Determine machine for new beam set
    # For imported doses, cannot use the old beam set's machine
    imported = old_beam_set.HasImportedDose()
    machine_name = old_beam_set.MachineReference.MachineName
    if imported:
        if machine_name.endswith('_imported'):
            machine_name = machine_name[:-9]
        elif is_sabr(old_beam_set):
            machine_name = 'SBRT 6MV'
        else:
            machine_name = 'ELEKTA'
        warnings += 'Machine "' + old_beam_set.MachineReference.MachineName + '" is not commissioned, so new beam set uses machine "' + machine_name + '".'

    planning_exam = plan.GetTotalDoseStructureSet().OnExamination
    
    tx_technique = get_tx_technique(old_beam_set)
    if tx_technique is None:
        MessageBox.Show('Treatment technique could not be determined. Click OK to abort the script.', 'Unrecognized Treatment Technique')
        sys.exit()

    # Existing names so new names can be made unique
    existing_beam_set_names, existing_beam_names, beam_num = names_nums(patient)

    # Get unique beam set name
    new_beam_set_name = unique_name(old_beam_set.DicomPlanLabel, existing_beam_set_names)
    new_beam_set = plan.AddNewBeamSet(Name=new_beam_set_name, ExaminationName=planning_exam.Name, MachineName=machine_name, Modality=old_beam_set.Modality, TreatmentTechnique=tx_technique, PatientPosition=old_beam_set.PatientPosition, NumberOfFractions=old_beam_set.FractionationPattern.NumberOfFractions, CreateSetupBeams=old_beam_set.PatientSetup.UseSetupBeams, Comment='Copy of "' + old_beam_set.DicomPlanLabel + '"')
    existing_beam_set_names.append(new_beam_set_name)

    # Copy the beams
    if not imported:  # Super simple for non-imported doses!
        new_beam_set.CopyBeamsFromBeamSet(BeamSetToCopyFrom=old_beam_set, BeamsToCopy=[beam.Name for beam in old_beam_set.Beams])

        # Rename and -number the new beams
        for beam in new_beam_set.Beams:
            beam.Number = beam_num
            beam.Name = unique_name(beam.Name, existing_beam_names)
            existing_beam_names.append(beam.Name)
            beam_num += 1
    else:  # CopyBeamsFromBeamSet does not work w/ imported dose
        for i, old_beam in enumerate(old_beam_set.Beams):
            iso_data = new_beam_set.CreateDefaultIsocenterData(Position=old_beam.Isocenter.Position)
            iso_data['Name'] = iso_data['NameOfIsocenterToRef'] = unique_iso_name(old_beam.Isocenter.Annotation.Name, new_beam_set, plan)
            
            qual = old_beam.BeamQualityId
            name = unique_name(old_beam.Name, existing_beam_names)

            if old_beam_set.Modality == 'Electrons':
                new_beam = new_beam_set.CreateElectronBeam(BeamQualityId=qual, Name=name, GantryAngle=old_beam.GantryAngle, CouchAngle=old_beam.CouchAngle, ApplicatorName=old_beam.Applicator.ElectronApplicatorName, InsertName=old_beam.Applicator.Insert.Name, IsAddCutoutChecked=True, IsocenterData=iso_data)
                existing_beam_names.append(name)
                new_beam.Applicator.Insert.Contour = old_beam.Applicator.Insert.Contour
                new_beam.BeamMU = old_beam.BeamMU
                new_beam.Description = old_beam.Description

            elif tx_technique != 'VMAT':
                existing_beam_names.append(name)
                # Create new beam for each segment
                seg_beam_names = []
                for s in old_beam.Segments:
                    seg_beam_name = unique_name(name, existing_beam_names + seg_beam_names)
                    new_beam = new_beam_set.CreatePhotonBeam(BeamQualityId=qual, Name=seg_beam_name, GantryAngle=old_beam.GantryAngle, CouchRotationAngle=old_beam.CouchRotationAngle, CouchPitchAngle=old_beam.CouchPitchAngle, CouchRollAngle=old_beam.CouchRollAngle, CollimatorAngle=s.CollimatorAngle, IsocenterData=iso_data)  
                    seg_beam_names.append(seg_beam_name)
                    new_beam.BeamMU = round(old_beam.BeamMU * s.RelativeWeight, 2)
                    new_beam.CreateRectangularField()
                    new_beam.Segments[0].JawPositions = s.JawPositions
                    new_beam.Segments[0].LeafPositions = s.LeafPositions
                # Create one beam from all segments
                if old_beam.Segments.Count > 1:
                    new_beam_set.MergeBeamSegments(TargetBeamName=name, MergeBeamNames=[beam.Name for beam in new_beam_set.Beams][(i + 1):])
                new_beam.Description = old_beam.Description
            
            else:  # VMAT
                new_beam = new_beam_set.CreateArcBeam(ArcStopGantryAngle=old_beam.ArcStopGantryAngle, ArcRotationDirection=old_beam.ArcRotationDirection, BeamQualityId=qual, Name=name, GantryAngle=old_beam.GantryAngle, CouchRotationAngle=old_beam.CouchRotationAngle, CouchPitchAngle=old_beam.CouchPitchAngle, CouchRollAngle=old_beam.CouchRollAngle, CollimatorAngle=old_beam.InitialCollimatorAngle, IsocenterData=iso_data)
                existing_beam_names.append(name)
                new_beam.BeamMU = old_beam.BeamMU
                new_beam.Description = old_beam.Description
            
            new_beam.Number = beam_num
            beam_num += 1

            # Does not work with uncommissioned machine
            #old_beam_set.ComputeDoseOnAdditionalSets(ExaminationNames=[planning_exam.Name], FractionNumbers=[0])

    copy_setup_beams(old_beam_set, new_beam_set, beam_num, existing_beam_names)
    
    # Copy objectives and constraints, and optimization parameters
    old_plan_opt = next(opt for opt in plan.PlanOptimizations if opt.OptimizedBeamSets.Count == 1 and opt.OptimizedBeamSets[0].Equals(old_beam_set))
    new_plan_opt = next(opt for opt in plan.PlanOptimizations if opt.OptimizedBeamSets.Count == 1 and opt.OptimizedBeamSets[0].Equals(new_beam_set))
    copy_objectives_and_constraints(old_plan_opt, new_plan_opt)
    copy_opt_params(old_plan_opt, new_plan_opt)

    # Copy Rx's
    copy_rxs(old_beam_set, new_beam_set)
