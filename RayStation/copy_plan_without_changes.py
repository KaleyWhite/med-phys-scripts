import re
import sys

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox  # FOr displaying errors


def copy_plan_without_changes():
    """Copies the current plan, retaining beam set and isocenter names.
    
    The name of the copy is the name of the old plan plus ' (1)'.
    The comments in the new plan are 'Copy of ___'
    The beam set and beam isocenter names match those of the old plan.
    The beam and setup beam numbers are unique across all cases in the current patient.
    The beam and setup beam names are the same as their numbers. If the old (setup) beam is different from its number, the old name is appended to the new (setup) beam's description.

    Does not switch to the new plan because this would require saving the patient.
    """

    # Ensure a patient is open
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Patient Open')
        sys.exit()

    # Ensure a case is open
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit() 
    
    # Ensure a plan is open
    try:
        plan = get_current('Plan')
    except:
        MessageBox.Show('There is no plan lopen. Click OK to abort the script.', 'No Plan Open')
        sys.exit()

    # Copy plan
    new_plan_name = plan.Name + ' (1)'
    case.CopyPlan(PlanName=plan.Name, NewPlanName=new_plan_name)
    new_plan = case.TreatmentPlans[new_plan_name]
    
    # Get unique beam number
    beam_num = 1
    for c in patient.Cases:
        for p in c.TreatmentPlans:
            for bs in p.BeamSets:
                if bs.Beams.Count > 0:
                    beam_num = max(beam_num, max(b.Number for b in bs.Beams))
                if bs.PatientSetup.SetupBeams.Count > 0:
                    beam_num = max(beam_num, max(b.Number for b in bs.PatientSetup.SetupBeams))
    
    # Rename beam sets, beams, and isos
    for i, bs in enumerate(new_plan.BeamSets):
        old_bs = plan.BeamSets[i]
        bs.DicomPlanLabel = old_bs.DicomPlanLabel
        # Beams
        for j, b in enumerate(bs.Beams):
            old_b = old_bs.Beams[j]
            b.Number = beam_num
            b.Name = str(beam_num)
            if old_b.Name == str(old_b.Number):
                b.Description = old_b.Description
            elif old_b.Description == '':
                b.Description = old_b.Name
            else:
                old_b.Description + '; ' + old_b.Name
            b.Isocenter.EditIsocenter(Name=old_b.Isocenter.Annotation.Name)
            beam_num += 1
        # Setup beams
        for j, sb in enumerate(bs.PatientSetup.SetupBeams):
            old_sb = old_bs.PatientSetup.SetupBeams[j]
            sb.Number = beam_num
            sb.Name = str(beam_num)
            if old_sb.Name == str(old_sb.Number) or re.match(fr'SB1_{j + 1}', old_sb.Name):
                sb.Description = old_sb.Description
            elif old_sb.Description == '':
                sb.Description = old_sb.Name
            else:
                sb.Description + '; ' + old_sb.Name
            sb.Isocenter.EditIsocenter(Name=old_sb.Isocenter.Annotation.Name)
            beam_num += 1


if __name__ == '__main__':
    copy_plan_without_changes()
