import clr
import sys

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox


def is_4d_exam(exam):
    dcm = exam.GetAcquisitionDataFromDicom()
    desc = dcm['SeriesModule']['SeriesDescription']
    if desc is None:
        return False
    return any(word in desc for word in ('AVG', 'MIP', 'Gated'))


def convert_virtual_jaw_to_mlc():
    """Converts virtual jaw fields from an external simulation (no MLCs) into jaw- and MLC-defined fields

    Machine is set to 'SBRT 6MV' if there are any AVG, MIP, or gated exams.
    This function is modified from SimPlanConvert_v4 from RayStation support.
    """
    try:
        patient = get_current('Patient')
        try:
            case = get_current('Case')
            try:
                plan = get_current('Plan')
            except:
                MessageBox.Show('There are no plans in the current case. Click OK to abort the script.', 'No Plans')
                sys.exit()
        except:
            MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
            sys.exit()
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()

    clinical_machine = 'SBRT 6MV' if any(is_4d_exam(exam) for exam in case.Examinations) else 'ELEKTA'

    # Delete any beam sets created by a previous run of this script
    # These beam set names end w/ a hyphen
    to_del = [beam_set.DicomPlanLabel for beam_set in plan.BeamSets if beam_set.DicomPlanLabel.endswith('-')]
    for name in to_del:
        plan.BeamSets[name].DeleteBeamSet()

    i = 0
    sim_beam_sets = [beam_set for beam_set in plan.BeamSets if beam_set.MachineReference.MachineName not in ('SBRT 6MV', 'ELEKTA')]  # If machine is correct, beam set is not a sim beam set
    while i < len(sim_beam_sets):
        sim_beam_set = sim_beam_sets[i]

        exam = sim_beam_set.GetPlanningExamination()
        position = sim_beam_set.PatientPosition

        # Create clinical beamset
        clin_beam_set = plan.AddNewBeamSet(Name=sim_beam_set.DicomPlanLabel[:(16 - i)] + '-' * (i + 1), ExaminationName=exam.Name, MachineName=clinical_machine, Modality='Photons', TreatmentTechnique='Conformal', PatientPosition=position, CreateSetupBeams=True, NumberOfFractions=1, UseLocalizationPointAsSetupIsocenter=True, Comment='')

        # Iterate over each beam in the simulation beamset 
        for beam in sim_beam_set.Beams:
            iso_data = beam_set.CreateDefaultIsocenterData(Position=beam.Isocenter.Position)
            iso_data['Name'] = iso_data['NameOfIsocenterToRef'] = beam.Isocenter.Annotation.Name
            
            qual_id = beam.BeamQualityId if beam.BeamQualityId != '' else '6'

            # Create the new beam
            new_beam = clin_beam_set.CreatePhotonBeam(BeamQualityId=qual_id, IsocenterData=iso_data, Name=beam.Name, Description=beam.Description, GantryAngle=beam.GantryAngle, CouchRotationAngle=beam.CouchRotationAngle, CouchPitchAngle=beam.CouchPitchAngle, CouchRollAngle=beam.CouchRollAngle, CollimatorAngle=beam.InitialCollimatorAngle)

            # Create the field
            x1, x2, y1, y2 = beam.Segments[0].JawPositions
            x_width = x2 - x1
            y_width = y2 - y1
            x_ctr = (x1 + x2) / 2
            y_ctr = (y1 + y2) / 2
            new_beam.CreateRectangularField(Width=x_width, Height=y_width, CenterCoordinate={'x': x_ctr, 'y': y_ctr}, MoveMLC=True, MoveAllMLCLeaves=False, MoveJaw=True, JawMargins={'x': 0, 'y': 0}, DeleteWedge=False, PreventExtraLeafPairFromOpening=False)

        # Change setup beam names to match descriptions, as in v8B
        for setup_beam in sim_beam_set.PatientSetup.SetupBeams:
            setup_beam.Name = setup_beam.Description

        sim_beam_set.DeleteBeamSet()
        clin_beam_set.DicomPlanLabel = clin_beam_set.DicomPlanLabel[:(-i - 1)]
        i += 1

    plan.Comments = 'Converted from sim'
    patient.Save()
