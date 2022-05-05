import clr
from datetime import datetime
import os
import re
import sys
from typing import Optional, TextIO

import pandas as pd

from connect import *
from connect.connect_cpython import PyScriptObject  # For type hints

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox


class LastModifiedTimes(object):
    """Creates and opens a TXT file listing the last modified dates and times for objects that have this information:
        - Patient
        - Registrations
        - Structure sets
        - Beam sets
        - Plan dose
            - Beam set dose
                - Beam dose
        - Evaluation doses
    """
    # Absolute path of directory to save output file in
    # This directory does not have to already exist
    OUTPUT_DIR = os.path.join('T:', os.sep, 'Physics', 'Scripts', 'Output Files', 'Last Modified Times')

    # Absolute path to the spreadsheet containing the "CRMC Usernames" sheet
    CREDS_PATH = os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-spreadsheets', 'Credentials & Computer Info.xlsx')

    def __init__(self, patient: PyScriptObject, case: Optional[PyScriptObject] = None, plan: Optional[PyScriptObject] = None) -> None:
        """Initializes a LastModifiedTimes object
        
        Calls the "main" method, _last_modified_times
        
        Arguments
        ---------
        patient: The patient to write modification info for
        case: The case to write modification info for
        plan: The plan to write modification info for
        """
        self._patient, self._case, self._plan = patient, case, plan
        self._usernames: pd.DataFrame = self._read_usernames()
        self._last_modified_times()

    def _fmt_pt_name(self) -> str:
        """Converts the patient's Name attribute into a better format for display

        Arguments
        ---------
        pt: The patient whose name to format

        Returns
        -------
        The formatted patient name

        Example
        -------
        Given some_pt with Name "^Jones^Bill^^M":
        format_pt_name(some_pt) -> "Jones, Bill M"
        """
        parts = [part for part in re.split(r'\^+', self._patient.Name) if part != '']
        name = parts[0]
        if len(parts) > 0:
            name += ', ' + ' '.join(parts[1:])
        return name

    def _output_filepath(self) -> str:
        """Constructs a filepath for the output file

        Output filename is in the format "<patient name> Last Modified Times YYYY-MM-DD HH:MM:SS.txt"

        Returns
        -------
        The absolute filepath for the output file
        """
        # Create output directory if its doesn't exist
        if not os.path.isdir(self.OUTPUT_DIR):
            os.makedirs(self.OUTPUT_DIR)

        pt_name = self._fmt_pt_name()  # Nicely formatted patient name
        dt = datetime.now().strftime('%Y-%m-%d %H_%M_%S')  # Timestamp
        filename = pt_name + ' Last Modified Times ' + dt + '.txt'
        return os.path.join(self.OUTPUT_DIR, filename)  # Return absolute path

    def _read_usernames(self) -> pd.DataFrame:
        """Reads in the DataFrame linking usernames to names

        DataFrame index is the username
        There is a single column, "Name"

        Cleans up rows that look like this:
            Name  Username
            ----  --------
            Joe   Old: js123\nNew: kl1d4
        Splits into multiple rows and removes "Old: " and "New: " labels:
            Name  Username
            ----  --------
            Joe   js123
            Joe   kl1d4

        Returns
        -------
        The DataFrame of usernames and names
        """
        unames = pd.read_excel(self.CREDS_PATH, sheet_name='CRMC Usernames')  # Read in table from "CRMC Usernames" sheet
        unames['Username'] = unames['Username'].str.split('\n')
        unames = unames.explode('Username', ignore_index=True)
        unames['Username'] = unames['Username'].apply(lambda s: s.split(': ')[-1])  # Remove "Old: " and "New: " labels
        unames = unames.reset_index(drop=True).set_index('Username')  # Make "Username" column the index
        return unames

    def _write_mod_info(self, lbl: str, obj: PyScriptObject, handle: TextIO) -> None:
        """Writes a PyScriptObject's modification info to the file

        Info is written in the following format: "<label>: <datetime> by <username or name>"
        If the username is in the usernames DataFrame, the actual name is used instead of the username

        Arguments
        ---------
        lbl: The "name" of the object whose modification info to write
        obj: The object whose modification info to write
        handle: The file-like object to write to
        """
        mod_info = obj.ModificationInfo
        if mod_info is None:
            handle.write(lbl + ': Modification info unavailable because updates need saving\n')
            return

        # User
        user = mod_info.UserName.split("\\")[-1]
        try:
            user = self._usernames.loc[user, 'Name']
        except KeyError:
            pass

        # Timestamp
        dt = mod_info.ModificationTime.ToString()
        
        handle.write(lbl + ': ' + dt + ' by ' + user + '\n')

    def _write_pt(self, handle):
        """Writes patient modification info to a file

        Arguments
        ---------
        handle: The file-like object to write to
        """
        self._write_mod_info('Patient', self._patient, handle)
        handle.write('\n')

    def _write_registrations(self, handle):
        """Writes registration modification info to a file

        Arguments
        ---------
        handle: The file-like object to write to
        """
        if self._case is None:
            return
        registrations = self._case.Registrations
        if registrations.Count == 0:
            return
        for r in registrations:
            r_name = r.RegistrationSource.FromExamination.Name + ' to ' + r.RegistrationSource.ToExamination.Name
            self._write_mod_info('Registration "' + r_name + '"', r, handle)
        handle.write('\n')

    def _write_struct_sets(self, handle):
        """Writes structure set modification info to a file

        Arguments
        ---------
        handle: The file-like object to write to
        """
        if self._case is None:
            return
        struct_sets = self._case.PatientModel.StructureSets
        if struct_sets.Count == 0:
            return
        for struct_set in struct_sets:
            self._write_mod_info('Structure set on "' + struct_set.OnExamination.Name + '"', struct_set, handle)
        handle.write('\n')

    def _write_beam_sets(self, handle):
        """Writes beam set modification info to a file

        Arguments
        ---------
        handle: The file-like object to write to
        """
        if self._plan is None:
            return
        beam_sets = self._plan.BeamSets
        if beam_sets.Count == 0:
            return
        for beam_set in beam_sets:
            self._write_mod_info('Beam set "' + beam_set.DicomPlanLabel + '"', beam_set, handle)
        handle.write('\n')

    def _write_doses(self, handle):
        """Writes plan, beam set, and beam dose modification info to a file

        Arguments
        ---------
        handle: The file-like object to write to
        """
        if self._plan is None:
            return
        self._write_mod_info('Plan dose', self._plan.TreatmentCourse.TotalDose, handle)
        beam_sets = self._plan.BeamSets
        if beam_sets.Count == 0:
            return
        for beam_set in beam_sets:
            handle.write('\t')
            self._write_mod_info(f'Beam set dose "' + beam_set.DicomPlanLabel + '"', beam_set.FractionDose, handle)
            beams = beam_set.Beams
            if beams.Count == 0:
                continue
            for i, beam in enumerate(beams):
                handle.write('\t\t')
                self._write_mod_info('Beam dose "' + beam.Name + '"', beam_set.FractionDose.BeamDoses[i], handle)
            handle.write('\n')

    def _write_eval_doses(self, handle):
        """Writes evaluation dose modification info to a file

        Arguments
        ---------
        handle: The file-like object to write to
        """
        if self._case is None:
            return
        fes = self._case.TreatmentDelivery.FractionEvaluations
        if fes.Count == 0:
            return
        for fe in fes:
            for doe in fe.DoseOnExaminations:
                exam_name = doe.OnExamination.Name
                for de in doe.DoseEvaluations:
                    if de.PerturbedDoseProperties is not None:  # perturbed dose
                        rds = de.PerturbedDoseProperties.RelativeDensityShift
                        density = f'{(rds * 100):.1f}%'
                        iso = de.PerturbedDoseProperties.IsoCenterShift
                        isoctr = f'({iso.x:.1f}, {iso.z:.1f}, {(-iso.y):.1f}) cm'
                        beam_set_name = de.ForBeamSet.DicomPlanLabel
                        dose_txt = f'Perturbed dose of ' + beam_set_name + ':' + density + ', ' + isoctr
                    elif de.Name != '':  # not perturbed, but has a name (probably summed)
                        dose_txt = de.Name
                    elif hasattr(de, 'ByStructureRegistration'):  # registered (mapped) dose
                        reg_name = de.ByStructureRegistration.Name
                        name = de.OfDoseDistribution.ForBeamSet.DicomPlanLabel
                        dose_txt = 'Deformed dose of ' + name + ' by registration ' + reg_name
                    else:  # neither perturbed, summed, nor mapped
                        dose_txt = de.ForBeamSet.DicomPlanLabel
                    self._write_mod_info(dose_txt, de, handle)
                    handle.write('\n')

    def _last_modified_times(self):
        """Writes modification info for several objects"""
        filepath = self._output_filepath()
        with open(filepath, 'w') as f:
            self._write_pt(f)
            self._write_registrations(f)
            self._write_struct_sets(f)
            self._write_beam_sets(f)
            self._write_doses(f)
            self._write_eval_doses(f)
        os.system('"' + filepath + '"')


def last_modified_times():
    """Writes modification info to a file"""
    try:
        patient = get_current('Patient')
    except:
        MessageBox.Show('There is no patient open. Click OK to abort the script.', 'No Open Patient')
        sys.exit()
    try:
        case = get_current('Case')
    except:
        MessageBox.Show('There is no case open. Click OK to abort the script.', 'No Open Case')
        sys.exit()
    try:
        plan = get_current('Plan')
    except:
        plan = None

    LastModifiedTimes(patient, case, plan)
