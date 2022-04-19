"""
This attached script is provided as a tool and not as a RaySearch endorsed script for
clinical use.  The use of this script in its whole or in part is done so
without any guarantees of accuracy or expected outcomes. Please verify all results. Further,
per the Raystation Instructions for Use, ALL scripts MUST be verified by the user prior to
clinical use.

PaddickAndHomogeneity_01.py
PaddickAndHomogeneity_02.py fixed line 60-61 to use current case
PaddickAndHomogeneity_03.py fixed paddickConformity = (tvPivVolume * tvPivVolume)/(vRiVolume * abs_target_vol)
PaddickAndHomogeneity_04.py CPython, RS 10 compatible
PaddickAndHomogeneity_05.py RS 11 compatible
PaddickAndHomogeneity_06.py D0.03cc <= 150%, V100% >= 95%, CI_ring5-6 100% <= 1-2, GI 50% <= 3-5, PCI
PaddickAndHomogeneity_07.py Beamset doses (per Rx), CG fix (no delete)
PaddickAndHomogeneity_08.py iterate thru Rx ROIs in plan
PaddickAndHomogeneity_09.py RTOG Conformity
PaddickAndHomogeneity_10.py repackaging & radial limit on Paddick exp_size

The statistics will appear in a pop-up window and in the Execution Details
"""
"""NOTE: This is a modified versio of PaddickAndHomogeneity_10 from RaySearch support. My changes are as follows:
- Better conform to Python 3 style guidelines.
"""
import clr
import sys

from connect import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox

case = get_current("Case")
patient = get_current("Patient")
plan = get_current("Plan")

ui = get_current('ui')
_version = ui.GetApplicationVersion().split('.')
version = int(_version[0]) if int(_version[1]) < 99 else int(_version[0]) + 1
subversion = int(_version[1])


def dose_limit(exam_name, targetName, idl_roi_name, exp_size=2):
    try:
        dlv_roi_name = case.PatientModel.GetUniqueRoiName(DesiredName="{}_DLV".format(targetName))
    except:
        dlv_roi_name = "_{}_DLV".format(targetName)
    dlv_roi = case.PatientModel.CreateRoi(Name=dlv_roi_name, Color="white", Type="DoseRegion")
    patient.SetRoiVisibility(RoiName=dlv_roi_name, IsVisible=False)
    dlv_roi.SetAlgebraExpression(ExpressionA={'Operation': "Union", 'SourceRoiNames': [targetName],
                                              'MarginSettings': {'Type': "Expand",
                                                                 'Superior': exp_size, 'Inferior': exp_size,
                                                                 'Anterior': exp_size, 'Posterior': exp_size,
                                                                 'Right': exp_size, 'Left': exp_size}},
                                 ExpressionB={'Operation': "Union", 'SourceRoiNames': [idl_roi_name],
                                              'MarginSettings': {'Type': "Expand",
                                                                 'Superior': 0, 'Inferior': 0,
                                                                 'Anterior': 0, 'Posterior': 0,
                                                                 'Right': 0, 'Left': 0}},
                                 ResultOperation="Intersection",
                                 ResultMarginSettings={'Type': "Expand",
                                                       'Superior': 0, 'Inferior': 0,
                                                       'Anterior': 0, 'Posterior': 0,
                                                       'Right': 0, 'Left': 0})
    dlv_roi.UpdateDerivedGeometry(Examination=case.Examinations[exam_name], Algorithm="Auto")
    return dlv_roi_name


finalMessage = ''
target_rois = []
temp_rois = []
# get all Rx ROIs
# for bs in plan.BeamSets:
#     try:
#         target_rois += [https://protect-us.mimecast.com/s/aoZuCW6j8LI5Zk97smYgsH?domain=rx.onstructure.name for rx in bs.Prescription.PrescriptionDoseReferences]
#     except:
#         target_rois += [https://protect-us.mimecast.com/s/aoZuCW6j8LI5Zk97smYgsH?domain=rx.onstructure.name for rx in bs.Prescription.DosePrescriptions]

# get all ROIs
for bs in plan.BeamSets:
    try:
        target_rois += [roi.OfRoi.Name for roi in plan.GetStructureSet().RoiGeometries if roi.HasContours() and roi.OfRoi.OrganData.OrganType == 'Target']
        all_rois = [roi.OfRoi.Name for roi in plan.GetStructureSet().RoiGeometries if roi.HasContours()]
        #target_rois += [https://protect-us.mimecast.com/s/9TJDCXDkMLFXjJ5vT9t5OI?domain=roi.ofroi.name for roi in plan.GetStructureSet().RoiGeometries if roi.HasContours() and roi.OfRoi.OrganData.OrganType == 'Target']
        #all_rois = [https://protect-us.mimecast.com/s/9TJDCXDkMLFXjJ5vT9t5OI?domain=roi.ofroi.name for roi in plan.GetStructureSet().RoiGeometries if roi.HasContours()]
    except:
        target_rois += [roi.OfRoi.Name for roi in bs.GetStructureSet().RoiGeometries if roi.HasContours() and roi.OfRoi.OrganData.OrganType == 'Target']
        all_rois = [roi.OfRoi.Name for roi in bs.GetStructureSet().RoiGeometries if roi.HasContours()]
        #target_rois += [https://protect-us.mimecast.com/s/9TJDCXDkMLFXjJ5vT9t5OI?domain=roi.ofroi.name for roi in bs.GetStructureSet().RoiGeometries if roi.HasContours() and roi.OfRoi.OrganData.OrganType == 'Target']
        #all_rois = [https://protect-us.mimecast.com/s/9TJDCXDkMLFXjJ5vT9t5OI?domain=roi.ofroi.name for roi in bs.GetStructureSet().RoiGeometries if roi.HasContours()]
target_rois = list(set(target_rois))
all_rois = list(set(all_rois))

# MAIN PROGRAM
# for beam_set in plan.BeamSets: # iterate thru beamsets
if True: # just this beamset
    beam_set = get_current('BeamSet')
    try:
        exam_name = plan.GetStructureSet().OnExamination.Name # get_current("Examination")
    except:
        exam_name = beam_set.GetStructureSet().OnExamination.Name

    finalMessage += '===\t' + beam_set.DicomPlanLabel + '\t===\n'
    
    for index, targetName in enumerate(target_rois):
        finalMessage += ' \t' + beam_set.DicomPlanLabel + ' \t' + targetName + '\n'
        ##################################################
        # Paddick Conformity Index  pCl = (TV PIV)^2 / (TV * V RI)
        ##################################################
        # TV   target statistics
        if version <= 10:
            prescription = beam_set.Prescription.PrimaryDosePrescription # up to 9B, 10A, 10B
        else:
            prescription = beam_set.Prescription.PrimaryPrescriptionDoseReference # 11A and forward

        # targetName = https://protect-us.mimecast.com/s/OIjKCYElMWILm81Rs3W63v?domain=prescription.onstructure.name
        abs_target_vol = case.PatientModel.StructureSets[exam_name].RoiGeometries[targetName].GetRoiVolume()

        # V RI   total volume covered by Rx isodose
        try:
            rxDoseVolName = case.PatientModel.GetUniqueRoiName(DesiredName = "Rx_Dose_Volume")
        except:
            rxDoseVolName = "Rx_Dose_Volume_"

        rxDose = prescription.DoseValue
        isodoseRoi = case.PatientModel.CreateRoi(Name = rxDoseVolName, Color = "Yellow", Type = "DoseRegion")
        patient.SetRoiVisibility(RoiName = rxDoseVolName, IsVisible = False)
        temp_rois.append(rxDoseVolName)
        fractions = beam_set.FractionationPattern.NumberOfFractions
        # isodoseRoi.CreateRoiGeometryFromDose(DoseDistribution=plan.TreatmentCourse.TotalDose, ThresholdLevel=rxDose)
        try:
            isodoseRoi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=rxDose/fractions)
        except Exception as e:
            print (e)

        # limit to around this target
        dlv_roi = dose_limit(exam_name, targetName, rxDoseVolName, exp_size=2.2) # a bit larger than 2.0cm
        temp_rois.append(dlv_roi)

        # vRiVolume = case.PatientModel.StructureSets[exam_name].RoiGeometries[rxDoseVolName].GetRoiVolume() # RX volume
        vRiVolume = case.PatientModel.StructureSets[exam_name].RoiGeometries[dlv_roi].GetRoiVolume() # RX volume

        # TV PIV  union algebra between TV and V RI
        try:
            overlapRoi = case.PatientModel.GetUniqueRoiName(DesiredName = "VolumeOverlap")
        except:
            overlapRoi = "VolumeOverlap"

        overlap_roi = case.PatientModel.CreateRoi(Name=overlapRoi, Color="255, 128, 0", Type="Undefined", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)
        patient.SetRoiVisibility(RoiName = overlapRoi, IsVisible = False)
        temp_rois.append(overlapRoi)
        overlap_roi.SetAlgebraExpression(ExpressionA={'Operation': "Union", 'SourceRoiNames': [rxDoseVolName], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [targetName], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Intersection", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
        overlap_roi.UpdateDerivedGeometry(Examination=case.Examinations[exam_name], Algorithm="Auto")
        if case.PatientModel.StructureSets[exam_name].RoiGeometries[overlapRoi].HasContours():
            tvPivVolume = case.PatientModel.StructureSets[exam_name].RoiGeometries[overlapRoi].GetRoiVolume()
        else:
            tvPivVolume = 0

        # calc
        paddickConformity = (tvPivVolume * tvPivVolume)/(vRiVolume * abs_target_vol)
        messagePaddick = str("The Paddick Conformity Index for {} is : {}".format(targetName, round(paddickConformity,3)))
        print (messagePaddick)
        finalMessage += messagePaddick + '\n'

        # cleanup/delete ROis
        # case.PatientModel.RegionsOfInterest[overlapRoi].DeleteRoi()

        ##################################################
        # RTOG Conformity Index  rCl = (prescription isodose volume) / (target volume)
        ##################################################
        rtogConformity = vRiVolume / abs_target_vol
        messagertogConformity = str("The RTOG Conformity Index for {} is : {}".format(targetName, round(rtogConformity,3)))
        print (messagertogConformity)
        finalMessage += messagertogConformity + '\n'

        ##################################################
        # CTV homogeneity (conformity) index (D5 - D95)/Drx
        ##################################################
        # D5
        # fivePercentDoseVolume = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = targetName, RelativeVolumes = [0.05])
        fivePercentDoseVolume = beam_set.FractionDose.GetDoseAtRelativeVolumes(RoiName = targetName, RelativeVolumes = [0.05])
        # D95
        # nintyfivePercentDoseVolume = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = targetName, RelativeVolumes = [0.95])
        nintyfivePercentDoseVolume = beam_set.FractionDose.GetDoseAtRelativeVolumes(RoiName = targetName, RelativeVolumes = [0.95])

        # calc
        ctvConformityIndex = (fivePercentDoseVolume[0] - nintyfivePercentDoseVolume[0]) / rxDose
        ctvMessageConformity = str("The Homogeneity (Conformity) Index for {} is : {}".format(targetName, round(ctvConformityIndex,3)))
        print (ctvMessageConformity)
        finalMessage += ctvMessageConformity + '\n'

        # cleanup/delete ROis
        # case.PatientModel.RegionsOfInterest[rxDoseVolName].DeleteRoi()

        ##################################################
        # Gradient index GI = PIVhalf / PIV
        # where PIVhalf is the rx isodose volume, at half the rx isodose and PIV is the rx isodose volume
        ##################################################
        try:
            rx50DoseVolName = case.PatientModel.GetUniqueRoiName(DesiredName = "Rx_50Dose_Volume")
        except:
            rx50DoseVolName = "Rx_50Dose_Volume"
        fiftyPercentIsodoseRoi = case.PatientModel.CreateRoi(Name = rx50DoseVolName, Color = "Yellow", Type = "DoseRegion")
        # fiftyPercentIsodoseRoi.CreateRoiGeometryFromDose(DoseDistribution=plan.TreatmentCourse.TotalDose, ThresholdLevel=rxDose*(0.50))
        patient.SetRoiVisibility(RoiName = rx50DoseVolName, IsVisible = False)
        temp_rois.append(rx50DoseVolName)
        fiftyPercentIsodoseRoi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=rxDose*(0.50)/fractions)
        fiftyPercentIsodoseVolume = case.PatientModel.StructureSets[exam_name].RoiGeometries[rx50DoseVolName].GetRoiVolume()

        try:
            rx100DoseVolName = case.PatientModel.GetUniqueRoiName(DesiredName = "Rx_100Dose_Volume")
        except:
            rx100DoseVolName = "Rx_100Dose_Volume"

        isodoseRoi = case.PatientModel.CreateRoi(Name = rx100DoseVolName, Color = "Yellow", Type = "DoseRegion")
        # isodoseRoi.CreateRoiGeometryFromDose(DoseDistribution=plan.TreatmentCourse.TotalDose, ThresholdLevel=rxDose)
        patient.SetRoiVisibility(RoiName = rx100DoseVolName, IsVisible = False)
        temp_rois.append(rx100DoseVolName)
        isodoseRoi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=rxDose/fractions)
        rxVolume = case.PatientModel.StructureSets[exam_name].RoiGeometries[rx100DoseVolName].GetRoiVolume()

        gradientIndex = fiftyPercentIsodoseVolume/rxVolume
        gradientIndexMsg = str("The Gradient Index for {} is : {}".format(targetName, round(gradientIndex,3)))
        print (gradientIndexMsg)
        finalMessage += gradientIndexMsg + '\n'

        # cleanup/delete ROis
        # case.PatientModel.RegionsOfInterest[rx50DoseVolName].DeleteRoi()
        # case.PatientModel.RegionsOfInterest[rx100DoseVolName].DeleteRoi()

        ##################################################
        # dose at 0.03 cc
        ##################################################
        maxPixelDose = beam_set.FractionDose.GetDoseAtRelativeVolumes(RoiName = targetName, RelativeVolumes = [0.03/abs_target_vol])[0]
        dMaxMessage = str("The Dose to 0.03cc in {} is : {} cGy".format(targetName,round(fractions * maxPixelDose, 2)))
        print(dMaxMessage)
        finalMessage += dMaxMessage + '\n'

        ##################################################
        # volume at Rx dose (percent coverage)
        ##################################################
        coveragePercentage = beam_set.FractionDose.GetRelativeVolumeAtDoseValues(RoiName = targetName, DoseValues = [rxDose/fractions])[0]
        # print(coveragePercentage)
        coverageMessage = str("{}% coverage on {} at {} cGy".format(round(coveragePercentage * 100, 2), targetName, rxDose))
        print(coverageMessage)
        finalMessage += coverageMessage + '\n'

        ##################################################
        # V12 (Vol recieving 12Gy)
        ##################################################
        if index == 0: # this is the first ROI so we need these general statistics
            try:
                twelve_gray = case.PatientModel.GetUniqueRoiName(DesiredName = "12Gy_Dose_Volume")
            except:
                twelve_gray = "12Gy_Dose_Volume"

            twelve_gray_roi = case.PatientModel.CreateRoi(Name = twelve_gray, Color = "Yellow", Type = "DoseRegion")
            # isodoseRoi.CreateRoiGeometryFromDose(DoseDistribution=plan.TreatmentCourse.TotalDose, ThresholdLevel=rxDose)
            patient.SetRoiVisibility(RoiName = twelve_gray, IsVisible = False)
            temp_rois.append(twelve_gray)
            twelve_gray_roi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=1200.0/fractions)
            twelve_gray_vol = case.PatientModel.StructureSets[exam_name].RoiGeometries[twelve_gray].GetRoiVolume()
            twelveGrayMessage = str("Volume recieving 12Gy is {} cc".format(round(twelve_gray_vol, 3)))
        print(twelveGrayMessage)
        finalMessage += twelveGrayMessage + '\n'


        ##################################################
        # V4.5
        ##################################################
        if index == 0: # this is the first ROI so we need these general statistics
            try:
                four_ana_half_gray = case.PatientModel.GetUniqueRoiName(DesiredName = "450cGy_Dose_Volume")
            except:
                four_ana_half_gray = "450cGy_Dose_Volume"

            twelve_gray_roi = case.PatientModel.CreateRoi(Name = four_ana_half_gray, Color = "Yellow", Type = "DoseRegion")
            # isodoseRoi.CreateRoiGeometryFromDose(DoseDistribution=plan.TreatmentCourse.TotalDose, ThresholdLevel=rxDose)
            patient.SetRoiVisibility(RoiName = four_ana_half_gray, IsVisible = False)
            temp_rois.append(four_ana_half_gray)
            twelve_gray_roi.CreateRoiGeometryFromDose(DoseDistribution=beam_set.FractionDose, ThresholdLevel=450/fractions)
            four_ana_half_gray_vol = case.PatientModel.StructureSets[exam_name].RoiGeometries[four_ana_half_gray].GetRoiVolume()
            fourPointFiveGrayMessage = str("Volume recieving 4.5Gy is {} cc".format(round(four_ana_half_gray_vol, 3)))
        print(fourPointFiveGrayMessage)
        finalMessage += fourPointFiveGrayMessage + '\n'

        ##################################################
        # Mean Brain Dose
        ##################################################
        if index == 0: # this is the first ROI so we need these general statistics
            brain_name = 'Brain'
            if brain_name in all_rois:
                mean_brain_dose = beam_set.FractionDose.GetDoseStatistic(RoiName = brain_name, DoseType = 'Average')
                meanBrainDoseMessage = str(
                    "The mean brain dose is {} cGy".format(round(mean_brain_dose, 3)))
                print(meanBrainDoseMessage)
                finalMessage += meanBrainDoseMessage + '\n'
            else:
                print('No Brain ROI found')


        ##################################################
        # EOF
        ##################################################
        print('\n')

##################################################
# cleanup
##################################################
for r in temp_rois:
    case.PatientModel.RegionsOfInterest[r].DeleteRoi()

##################################################
# print
##################################################
print (finalMessage)
MessageBox.Show(0, finalMessage, "Conformity Indexes", 0)
