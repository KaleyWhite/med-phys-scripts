"""
Before running script:
1. Import patient into RayStation Doctor from Storage SCP
2. Lock shift and save patient
"""

# For GUI
import clr
clr.AddReference("System.Windows.Forms")

import os
import sys
sys.path.append(os.path.join("T:", "Physics - T", "Scripts", "RayStation"))

from connect import *  # Interact w/ RS
from System.Windows.Forms import *  # For GUI

from ApplyTemplatesForm import *
from ConvertVirtualJawToMLCScript import *
from PrepareExamsScript import *


def is_sbrt_exam(exam):
  # Helper function that returns True if exam description contains "AVG", "MIP", or "Gated", False otherwise.

  desc = exam.GetAcquisitionDataFromDicom()["SeriesModule"]["SeriesDescription"]
  if desc is None:
    return False
  return any([word in desc for word in ["AVG", "MIP", "Gated"]])


def name_item(item, l, max_len=sys.maxint):
    # Helper function that generates a unique name for `item` in list `l`
    # Limit name to 16 characters
    # E.g., name_item("Isocenter Name A", ["Isocenter Name A", "Isocenter Na (1)", "Isocenter N (10)"]) -> "Isocenter N (11)"

    copy_num = 0
    for l_item in l:
        m = re.search(" \((\d+)\)$", l_item, re.IGNORECASE)
        if m is None and l_item == item[:max_len]:
            copy_num = max(copy_num, 1)
        elif m is not None and l_item[:m.start()] == item[:(max_len - len(m.group()))]:
            copy_num = max(copy_num, int(m.group(1)) + 1)
    if copy_num == 0:
        return item[:max_len]
    return "{} ({})".format(item[:(max_len - 3 - len(str(copy_num)))], copy_num)


def sbrt_lung_prep():
    # Get current variables
    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort the script.", "No Patient Loaded")
        sys.exit(1)
    try:
        plan = get_current("Plan")
    except:
        MessageBox.Show("There is no plan loaded. Click OK to abort the script.", "No Plan Loaded")
        sys.exit(1)
    patient_db = get_current("PatientDB")
    case = get_current("Case")

    # Convert virtual jaw to MLC
    convert_virtual_jaw_to_mlc()

    # Rename "Trial_1" plan to "Initial Sim"
    plan.Name = name_item("Initial Sim", [plan_.Name for plan_ in case.TreatmentPlans])

    # Rename exams, set imaging system, and create AVG/MIP if necessary
    planning_exam = prepare_exams()
    is_sbrt = any(is_sbrt_exam(exam) for exam in case.Examinations)

    # Apply structure template
    apply_templates()

    # Add external geometry to all exams
    for exam in case.PatientModel.RegionsOfInterest:
        try:
            case.PatientModel.RegionsOfInterest["External"].CreateExternalGeometry(Examination=exam.Name)
        except:  # External is approved on this exam
            continue

    if is_sbrt:
        # L or R lung?
        # If loc point is left of horizontal center of planning exam, it's left lung. Otherwise, right.
        loc_x = case.PatientModel.StructureSets[planning_exam.Name].PoiGeometries["Localization point"].Point.x
        lung = "L" if loc_x > 0 else "R"

        # Create SBRT plan
        plan_name = name_item("SBRT Lung_{}".format(lung), [plan_.Name for plan_ in case.TreatmentPlans])
        with CompositeAction("Add Treatment plan"):
            plan = case.AddNewPlan(PlanName=plan_name, PlannedBy="CRMC", ExaminationName=planning_exam.Name)
            beam_set = plan.AddNewBeamSet(Name=plan_name, ExaminationName=planning_exam.Name, MachineName="SBRT 6MV", Modality="Photons", TreatmentTechnique="VMAT", PatientPosition="HeadFirstSupine", NumberOfFractions=5)
            beam_set.AddDosePrescriptionToRoi(RoiName="PTV", DoseVolume=95, PrescriptionType="DoseAtVolume", DoseValue=5000)

        # Set level to Lung presets
        for series in planning_exam.Series:
            series.LevelWindow = {"x": -600, "y": 1600}

        plan.SetDefaultDoseGrid(VoxelSize={ 'x': 0.2, 'y': 0.2, 'z': 0.2 })

        # ADD BOX ROI

        ss = plan.GetStructureSet()
        inner_couch = [geom for geom in ss.RoiGeometries if geom.OfRoi.Type == "Support" and geom.OfRoi.RoiMaterial.OfMaterial.Name == "Cork"][0]
        outer_couch = [geom for geom in ss.RoiGeometries if geom.OfRoi.Type == "Support" and geom.OfRoi.RoiMaterial.OfMaterial.Name == "PMI foam"][0]
        ext = ss.RoiGeometries["External"]

        # RL height
        couch_min_x = inner_couch.GetBoundingBox()[0].x
        couch_max_x = inner_couch.GetBoundingBox()[1].x
        x = couch_max_x - couch_min_x

        # PA height
        loc_y = ss.PoiGeometries["Localization point"].Point.y
        couch_y = outer_couch.GetBoundingBox()[0].y
        y = couch_y - loc_y

        # IS height = height of External
        ext_min_z = ext.GetBoundingBox()[0].z
        ext_max_z = ext.GetBoundingBox()[1].z
        z = ext_max_z - ext_min_z

        z_ctr = (ext_max_z + ext_min_z) / 2
        y_ctr = (couch_y + loc_y) / 2  

        box = case.PatientModel.RegionsOfInterest["Box"]
        box.CreateBoxGeometry(Size={ 'x': x, 'y': y, 'z': z }, Examination=planning_exam, Center={ 'x': 0, 'y': y_ctr, 'z': z_ctr })
        case.PatientModel.RegionsOfInterest['External'].CreateAlgebraGeometry(Examination=planning_exam, ExpressionA={ 'Operation': "Union", 'SourceRoiNames': ["External"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': ["Box"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Union", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
        box.DeleteRoi()
    
    else:
        # Create new plan here
        #plan = 

        # Update dose grid
        plan.SetDefaultDoseGrid(VoxelSize={ 'x': 0.2, 'y': 0.2, 'z': 0.2 })
