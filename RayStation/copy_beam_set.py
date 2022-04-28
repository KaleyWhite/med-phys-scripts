import clr
import re
import sys

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons

sys.path.append(r'T:\Physics\KW\med-phys-scripts\RayStation')
from copy_opt_stuff import copy_objectives_and_constraints, copy_opt_params
from copy_plan_without_changes import copy_plan_without_changes


def is_sabr(beam_set):
    """Returns whether or not a beam set is SABR.

    SABR includes SBRT, SRT, and SRS. A SABR beam set is a VMAT beam set with no more than 15 fractions and at least 600 cGy fractional dose.

    Arguments
    ---------
    beam_set: The beam set to check whether it is SABR
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


def get_tx_technique(bs):
    # Helper function that returns the treatment technique for the given beam set.
    # SMLC, VMAT, DMLC, 3D-CRT, conformal arc, or applicator and cutout
    # Return '?' if treatment technique cannot be determined
    # Code modified from RS support

    if bs.Modality == 'Photons':
        if bs.PlanGenerationTechnique == 'Imrt':
            if bs.DeliveryTechnique == 'SMLC':
                return 'SMLC'
            if bs.DeliveryTechnique == 'DynamicArc':
                return 'VMAT'
            if bs.DeliveryTechnique == 'DMLC':
                return 'DMLC'
        elif bs.PlanGenerationTechnique == 'Conformal':
            if bs.DeliveryTechnique == 'SMLC':
                # return 'SMLC' # Changed from 'Conformal'. Failing with forward plans.
                return 'Conformal'
                # return '3D-CRT'
            if bs.DeliveryTechnique == 'Arc':
                return 'Conformal Arc'
    elif bs.Modality == 'Electrons' and bs.PlanGenerationTechnique == 'Conformal' and bs.DeliveryTechnique == 'SMLC':
        return 'ApplicatorAndCutout'


def get_unique_beam_num(patient):
    """Returns a unique beam number among all beams in the current patient

    Arguments
    ---------
    patient: The patient whose beam numbers to check
    """
    beam_num = 1
    for case in patient.Cases:
        for plan in case.TreatmentPlans:
            for beam_set in plan.BeamSets:
                for beam in beam_set.Beams:
                    beam_num = max(beam_num, beam.Number)
                for setup_beam in beam_set.PatientSetup.SetupBeams:
                    beam_num = max(beam_num, setup_beam.Number)
    return beam_num + 1


def unique_name(desired_name, existing_names, max_len=sys.maxsize):
    new_name = desired_name[:max_len]
    copy_num = 0
    while new_name in existing_names:
        copy_num += 1
        copy_str = ' (' + str(copy_num) + ')'
        name_len = 16 - len(copy_str)
        new_name = desired_name[:name_len] + copy_str
    return new_name


def names_nums(patient):
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


def copy_beam_set():
    """Copies the current beam set to a new beam set in the same plan
    
    Copy electron or photon (including VMAT) beam sets:
        - Treatment and setup beams
            * Unique beam numbers across all cases in current patient
            * Beam names are same as numbers
        - AutoScaleToPrescription
        - For VMAT beam sets, other select optimization settings:
            * Maximum number of iterations
            * Optimality tolerance
            * Calculate intermediate and final doses
            * Iterations in preparation phase
            * Max leaf travel distance per degree (and enabled/disabled)
            * Dual arcs
            * Max gantry spacing
            * Max delivery time
            * Objectives and constraints
        - Prescriptions

    Unfortunately, dose cannot accurately be copied for VMAT, so we do the next-best thing and provide the dose on additional set if the beam set is copied to a plan with a different planning exam.

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
            iso_data = beam_set.CreateDefaultIsocenterData(Position=beam.Isocenter.Position)
            iso_data['Name'] = iso_data['NameOfIsocenterToRef'] = beam.Isocenter.Annotation.Name
            
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
                    new_beam = new_beam_set.CreatePhotonBeam(Energy=qual, Name=seg_beam_name, GantryAngle=old_beam.GantryAngle, CouchAngle=old_beam.CouchAngle, CollimatorAngle=s.CollimatorAngle, IsocenterData=iso_data)  
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
                new_beam = new_beam_set.CreateArcBeam(ArcStopGantryAngle=old_beam.ArcStopGantryAngle, ArcRotationDirection=old_beam.ArcRotationDirection, BeamQualityId=qual, Name=name, GantryAngle=beam.GantryAngle, CouchRotationAngle=beam.CouchRotationAngle, CollimatorAngle=beam.InitialCollimatorAngle, IsocenterData=iso_data)
                existing_beam_names.append(name)
                new_beam.BeamMU = old_beam.BeamMU
                new_beam.Description = old_beam.Description
            
            new_beam.Number = beam_num
            beam_num += 1

            old_beam_set.ComputeDoseOnAdditionalSets(ExaminationNames=[planning_exam.Name], FractionNumbers=[0])

    # Manually copy setup beams from old beam set
    if old_beam_set.PatientSetup.UseSetupBeams and old_beam_set.PatientSetup.SetupBeams.Count > 0:
        old_setup_beams = old_beam_set.PatientSetup.SetupBeams
        new_beam_set.UpdateSetupBeams(ResetSetupBeams=True, SetupBeamsGantryAngles=[setup_beam.GantryAngle for setup_beam in old_setup_beams])  # Clear the setup beams created when the beam set was added, to ensure no extraneous setup beams in new beam set
        for i, old_setup_beam in enumerate(old_setup_beams):
            new_setup_beam = new_beam_set.PatientSetup.SetupBeams[i]
            new_setup_beam.Number = beam_num
            new_setup_beam.Name = unique_name(old_setup_beam.Name, existing_beam_names)
            new_setup_beam.Description = old_setup_beam.Description
            beam_num += 1
    
    # Copy objectives and constraints, and optimization parameters
    old_plan_opt = next(opt for opt in plan.PlanOptimizations if opt.OptimizedBeamSets.Count == 1 and opt.OptimizedBeamSets[0].Equals(old_beam_set))
    new_plan_opt = next(opt for opt in plan.PlanOptimizations if opt.OptimizedBeamSets.Count == 1 and opt.OptimizedBeamSets[0].Equals(new_beam_set))
    copy_objectives_and_constraints(old_plan_opt, new_plan_opt)
    copy_opt_params(old_plan_opt, new_plan_opt)

    # Copy Rx's
    if old_beam_set.Prescription is not None:
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
