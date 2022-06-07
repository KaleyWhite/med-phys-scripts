# -*- coding: utf-8 -*-
#!/usr/bin/env python3
from setup_crap import *

from raystation.raystation_utilities import exit_with_error, get_current_or_abort, tx_technique, unique_beam_num


def temporary():
    patient = get_current_or_abort('Patient')
    case = get_current_or_abort('Case')
    old_plan = get_current_or_abort('Plan')
    
    for beam_set in old_plan.BeamSets:
        if tx_technique(beam_set) != 'VMAT':
            exit_with_error(f'Plan "{old_plan.Name}" is not a VMAT plan.')
        elif beam_set.FractionationPattern is None:
            exit_with_error(f'Beam set "{beam_set.DicomPlanLabel}" has no specified fractionation.')

    plan_exam = plan.GetTotalDoseStructureSet().OnExamination

    beam_num = unique_beam_num(patient)
    for num_new_beams_per_beam in range(1, 11):
        for wting in ['wted', 'unwted']:
            new_plan_name = f'{num_new_beams_per_beam}, {wting}'
            new_plan = case.AddNewPlan(PlanName=new_plan_name, PlannedBy='KEW', ExaminationName=plan_exam.Name)
    
            for beam_set in old_plan.BeamSets:
                beam_set_exam = beam_set.GetPlanningExamination()
                beam_set_pos = beam_set.PatientPosition
                num_fx = beam_set.FractionationPattern.NumberOfFractions
                new_beam_set = new_plan.AddNewBeamSet(MachineName='ELEKTA', Name='temp', ExaminationName=beam_set_exam.Name, Modality='Photons', TreatmentTechnique='Conformal', PatientPosition=beam_set_pos, NumberOfFractions=num_fx)
        
                for beam in beam_set.Beams:
                    ini_gantry_angle = beam.GantryAngle
    
                    segs = list(beam.Segments)
                    num_segs = len(segs)
                    num_segs_per_new_beam = int(round(num_segs / num_new_beams_per_beam))
            
                    for seg_num in range(0, num_segs, num_segs_per_new_beam):
                        segs_for_new_beam = segs[seg_num:(seg_num + num_segs_per_new_beam)]
                        num_segs_for_new_beam = len(segs_for_new_beam)

                        if wting == 'wted':
                            avg_delta_gantry_angle = sum(seg.RelativeWeight * seg.DeltaGantryAngle for seg in segs_for_new_beam)
                            avg_coll_angle = sum(seg.RelativeWeight * seg.CollimatorAngle for seg in segs_for_new_beam)
                        else:
                            avg_delta_gantry_angle = sum(seg.DeltaGantryAngle for seg in segs_for_new_beam) / num_segs_for_new_beam
                            avg_coll_angle = sum(seg.CollimatorAngle for seg in segs_for_new_beam) / num_segs_for_new_beam
                
                        gantry_angle = ini_gantry_angle + avg_delta_gantry_angle
                        while gantry_angle < 0:
                            gantry_angle += 360
                        while gantry_angle > 360:
                            gantry_angle -= 360
                
                        qual_id = beam.BeamQualityId
                        iso_data = beam_set.GetIsocenterData(Name=beam.Isocenter.Annotation.Name)
                        new_beam = new_beam_set.CreatePhotonBeam(BeamQualityId=qual_id, IsocenterData=iso_data, Name=str(beam_num), GantryAngle=gantry_angle, CollimatorAngle=avg_coll_angle)
                        new_beam.ConformMlc()
                
                        new_beam.Number = beam_num
                        beam_num += 1

                        for i, jaw_pos in enumerate(new_beam.Segments[0].JawPositions):
                            if wting == 'wted':
                                jaw_pos = sum(seg.JawPositions[i] for seg in segs_for_new_beam) / num_segs_for_new_beam
                            else:
                                jaw_pos = sum(seg.RelativeWeight * seg.JawPositions[i] for seg in segs_for_new_beam)
                
                        for i, leaf_pos in enumerate(new_beam.Segments[0].LeafPositions):
                            for j, leaf_pos_ in enumerate(leaf_pos):
                                if wting == 'wted':
                                    leaf_pos_ = sum(seg.LeafPositions[i][j] for seg in segs_for_new_beam) / num_segs_for_new_beam
                                else:
                                    leaf_pos_ = sum(seg.RelativeWeight * seg.LeafPositions[i][j] for seg in segs_for_new_beam)

if __name__ == '__main__':
    temporary()
