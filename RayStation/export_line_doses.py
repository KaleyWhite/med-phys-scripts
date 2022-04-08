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

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')

from System import ArgumentException
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
    def __init__(self, case):
        self.case = case

        self.available_doses = OrderedDict()
        self.create_available_doses()

        self.btn_panel = Panel()
        self.data_grid_view = DataGridView()
        self.add_row_btn = Button()
        self.del_row_btn = Button()

        self.set_up_layout()
        self.set_up_data_grid_view()
        self.populate_data_grid_view()
        
    def add_row_btn_Click(self, sender, e):
        self.data_grid_view.Rows.Add()

    def del_row_btn_Click(self, sender, e):
        if self.data_grid_view.SelectedRows.Count > 0 and self.data_grid_view.SelectedRows[0].Index != self.data_grid_view.Rows.Count - 1:
            self.data_grid_view.Rows.RemoveAt(self.data_grid_view.SelectedRows[0].Index)

    def data_grid_view_CellValueChanged(self, sender, event):
        row = self.data_grid_view.Rows[event.RowIndex]
        cb = row.Cells[0]
        if cb.Value is not None:
            row.Cells[1].Items.Clear()
            points = list(self.available_doses[cb.Value].available_points)
            print(points)
            row.Cells[1].Items.AddRange(points)
            points_col = self.data_grid_view.Columns[1]
            width = TextRenderer.MeasureText(points_col.HeaderText, self.data_grid_view.ColumnHeadersDefaultCellStyle.Font).Width
            for row in self.data_grid_view.Rows:
                print(item for item in row.Cells[1].Items)
                width = max(width, max(TextRenderer.MeasureText(item, self.data_grid_view.Font).Width for item in row.Cells[1].Items))
            points_col.width = width

    def create_available_plan_doses(self):
        for plan in self.case.TreatmentPlans:
           points = get_poi_geoms(plan.GetTotalDoseStructureSet())
           if not points:
               continue
           dose_name = 'Plan: ' + plan.Name
           self.available_doses[dose_name] = Dose(plan.TreatmentCourse.TotalDose, points)
           for beam_set in plan.BeamSets:
               dose_name = ' ' * 4 + 'Beam set: ' + beam_set.DicomPlanLabel
               self.available_doses[dose_name] = Dose(beam_set.FractionDose, points)
               for beam_dose in beam_set.FractionDose.BeamDoses:
                   dose_name = ' ' * 8 + 'Beam: ' + beam_dose.ForBeam.Name
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

    def set_up_layout(self):
        self.Size = Size(600, 500)

        self.add_row_btn.Text = "Add Row"
        self.add_row_btn.Location = Point(10, 10)
        self.add_row_btn.Click += self.add_row_btn_Click

        self.del_row_btn.Text = "Delete Row"
        self.del_row_btn.Location = Point(100, 10)
        self.del_row_btn.Click += self.del_row_btn_Click

        self.btn_panel.Controls.Add(self.add_row_btn)
        self.btn_panel.Controls.Add(self.del_row_btn)
        self.btn_panel.Height = 50
        self.btn_panel.Dock = DockStyle.Bottom

        self.Controls.Add(self.btn_panel)

    def set_up_data_grid_view(self):
        self.Controls.Add(self.data_grid_view)

        self.data_grid_view.AutoGenerateColumns = False

        self.data_grid_view.ColumnHeadersDefaultCellStyle.Font = Font(self.data_grid_view.Font, FontStyle.Bold)

        self.data_grid_view.Name = "data_grid_view"
        self.data_grid_view.Location = Point(8, 8)
        self.data_grid_view.Size = Size(500, 250)

        # Basic formatting settings
        self.data_grid_view.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
        self.data_grid_view.RowHeadersVisible = False

        # Set property values appropriate for read-only display and limited interactivity. 
        self.data_grid_view.AllowUserToOrderColumns = True
        self.data_grid_view.AllowUserToResizeColumns = False
        self.data_grid_view.AllowUserToResizeRows = False
        self.data_grid_view.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self.data_grid_view.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        dose_col = self.create_combobox_col('Dose', list(self.available_doses))
        self.data_grid_view.Columns.Add(dose_col)

        point_col = self.create_combobox_col('Point')
        self.data_grid_view.Columns.Add(point_col)

        self.data_grid_view.CellValueChanged += self.data_grid_view_CellValueChanged
        
        """
        self.data_grid_view.Columns[0].Name = "Dose"
        self.data_grid_view.Columns[1].Name = "Point"
        self.data_grid_view.Columns[2].Name = "Endpoint/Center"
        self.data_grid_view.Columns[3].Name = "Length(s)"
        self.data_grid_view.Columns[4].Name = "Direction"
        """

        self.data_grid_view.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.data_grid_view.MultiSelect = True
        self.data_grid_view.Dock = DockStyle.Fill

    def create_combobox_col(self, hdr, choices=[]):
        col = DataGridViewComboBoxColumn()
        col.HeaderText = hdr
        col.Items.AddRange(choices)
        col.Width = TextRenderer.MeasureText(hdr, self.data_grid_view.ColumnHeadersDefaultCellStyle.Font).Width
        if choices:
            item_width = max(TextRenderer.MeasureText(item, self.data_grid_view.Font).Width for item in col.Items)
            col.Width = max(col.Width, item_width)
        return col

    def populate_data_grid_view(self):

        self.data_grid_view.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders)

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
    form = ExportLineDosesForm(case)
    form.ShowDialog()
    if form.DialogResult != DialogResult.OK:
        sys.exit()


export_line_doses()  
