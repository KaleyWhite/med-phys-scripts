import clr
from collections import OrderedDict
from datetime import datetime
import os
import re
import sys
from typing import Any, Dict, Optional, Tuple

import numpy as np

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')

from System.Drawing import *
from System.Windows.Forms import *


# Directory to which to write the exported CSV file
OUTPUT_DIR = r'T:\Physics\Scripts\Output Files\export_line_doses'


def get_mach_str(bs):
    # Returns the "MachineName" CSV header row for the beam set
    mach_ref = bs.MachineReference
    dt = mach_ref.CommissioningTime
    return '#TreatmentMachine:\t' + mach_ref.MachineName + ' Commission time: ' + datetime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second).strftime('%m/%d/%Y %H:%M:%S')  + '\n'


def get_poi_geoms(struct_set):
    points = OrderedDict()
    for poi_geom in struct_set.PoiGeometries:
        if poi_geom.Point is not None and abs(poi_geom.Point.x) != float('inf'):
            points['POI: ' + poi_geom.OfPoi.Name] = poi_geom.Point
    return points


class Dose(object):
    def __init__(self, dose, available_points):
        self.dose = dose
        self.available_points = available_points


class Row(object):
    row_id = 0

    def __init__(self, dose_cb, point_cb, endpt_or_ctr_cb, length_tb, direction_cb, del_btn):
        self.dose_cb = dose_cb
        self.point_cb = point_cb
        self.endpt_or_ctr_cb = endpt_or_ctr_cb
        self.length_tb = length_tb
        self.direction_cb = direction_cb
        self.del_btn = del_btn


class ExportLineDosesForm(Form):
    """Form that allows the user to choose parameters for exported line doses centered at POIs, DSPs, and/or beam isocenters.

    User chooses:
    - Dose(s): plan, beam set, beam, eval
    - Direction(s): X, Y, Z
    - Length(s), in cm
    - Center(s): POI(s), DSP(s), and/or beam isocenter(s)
    """

    def __init__(self, case: PyScriptObject) -> None:
        """Initializes an ExportLineDosesForm object.

        Args:
            case: The case from which to export line doses.
        """
        self.case = case

        self.available_doses = {}
        self.create_available_doses()

        self.export_btn = self.create_btn('Export')
        
        self.x = self.y = 15  # Horizontal and vertical coordinates of next control
        self.set_up_layout()

    def export_btn_click(self, sender, event):
        pass

    def create_available_plan_doses(self):
        for plan in self.case.TreatmentPlans:
           points = get_poi_geoms(plan.GetTotalDoseStructureSet())
           if not points:
               continue
           dose_name = 'Plan: ' + plan.Name
           self.available_doses[dose_name] = Dose(plan.TreatmentCourse.TotalDose, points)
           for beam_set in plan.BeamSets:
               dose_name = '\tBeam set: ' + beam_set.DicomPlanLabel
               self.available_doses[dose_name] = Dose(beam_set.FractionDose, points)
               for beam_dose in beam_set.FractionDose.BeamDoses:
                   dose_name = '\t\tBeam: ' + beam_dose.ForBeam.Name
                   self.available_doses[dose_name] = Dose(beam_dose, points)

    def create_available_eval_doses(self):
        # This code is modified from a RayStation Scripting Guidelines example script
        for fe in self.case.TreatmentDelivery.FractionEvaluations:
            for doe in fe.DoseOnExaminations:
                exam = doe.OnExamination
                points = get_poi_geoms(self.case.PatientModel.StructureSets[exam.Name])
                if not points:
                    continue
                for de in doe.DoseEvaluations:
                    #mach_str = get_mach_str(de.OfDoseDistribution.ForBeamSet)
                    if de.PerturbedDoseProperties is not None:
                        rds = de.PerturbedDoseProperties.RelativeDensityShift
                        density = f'{(rds * 100):.1f}%'
                        iso = de.PerturbedDoseProperties.IsoCenterShift
                        isoctr = f'({iso.x:.1f}, {iso.y:.1f}, {iso.z:.1f})'
                        beam_set_id = de.ForBeamSet.DicomPlanLabel
                        txt = f'Perturbed dose of beam set "{beam_set_id}": {density}, {isoctr} (exam "{exam.Name}")'
                    elif de.Name != '':
                        txt = f'"{de.Name}" (exam "{exam.Name}")'
                    elif hasattr(de, 'ByStructureRegistration'):
                        reg_name = de.ByStructureRegistration.Name
                        beam_set_id = de.OfDoseDistribution.ForBeamSet.DicomPlanLabel
                        txt = f'Deformed dose of beam set "{beam_set_id}" by registration "{reg_name}" (exam "{exam.Name}")'
                    else:
                        txt = f'Beam set "{de.ForBeamSet.DicomPlanLabel}" (exam "{exam.Name}")'
                    self.available_doses[txt] = Dose(de, points)
    
    def create_available_doses(self):
        self.create_available_plan_doses()
        self.create_available_eval_doses()

    def create_btn(self, txt):
        btn = Button()
        btn.AutoSize = True
        btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        btn.Text = txt
        return button

    def create_lbl(self, txt):
        lbl = Label()
        lbl.AutoSize = True
        lbl.AutoSizeMode = AutoSizeMode.GrowAndShrink
        lbl.Text = txt
        return lbl
    
    def set_up_layout(self):
        # Style this form
        self.Text = 'Export Line Doses'
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # Disallow user resizing
        self.StartPosition = FormStartPosition.CenterScreen  # Launch window in middle of screen

        # Adapt form size to contents
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)

        self.export_btn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.export_button.Click += self.export_button_Click
        self.Controls.Add(self.export_btn)

        self.add_dose_col()
    
    def add_dose_col(self):
        dose_lbl = self.create_lbl('Dose')
        dose_width = max(TextRenderer.MeasureText(dose_name, dose_lbl.Font).Width for dose_name in self.available_doses)
        dose_lbl.Font = Font(dose_lbl.Font, FontStyle.Bold)
        dose_width = max(dose_width, TextRenderer.MeasureText(dose_lbl.Text, dose_lbl.Font).Width)
        dose_lbl.Location = Point(self.x, self.y)
        

        self.add_dose_cb(self)

    def add_dose_cb(self, x, y):

    def add_row(self, y):
        


    def add_table_hdrs(self):



def export_line_doses():
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click Ok to abort the script.', 'No Open Patient')
        sys.exit()
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click Ok to abort the script.', 'No Open Case')
        sys.exit()

    # Launch user input form
    #form = ExportLineDosesForm(case)
    #form.ShowDialog()
    #if form.DialogResult != DialogResult.OK:
       # sys.exit()
    Application.EnableVisualStyles()
    form = ExportLineDosesForm(case)
    Application.Run(form)


if __name__ == '__main__':
    export_line_doses()  
