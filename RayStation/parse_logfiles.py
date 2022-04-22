"""Writes information from RayStation logfiles to an Excel spreadsheet.

See github.com/KaleyWhite/med-phys-scripts/Scripts User Manual for the spreadsheet's data dictionary.
Change LOGFILE_PATH and OUTPUT_PATH according to your needs.
"""
from collections import OrderedDict
from datetime import datetime
import os
import re
from typing import Union

import numpy as np
import pandas as pd


class ParseLogfiles(object):
    """Class that parses all '.log' files in the LOGFILE_PATH directory and writes selected information to an Excel spreadsheet

    '.exn.log' files are ignored as all exceptions are logged in the '.log' files.
    """
    LOGFILE_PATH = os.path.join('C:', os.sep, 'ProgramData', 'RaySearch')  # Path to logfiles for your RayStation installation
    OUTPUT_PATH = os.path.join('T:', os.sep, 'Physics', 'Scripts', 'Output Files', 'Parse Logfiles')  # Directory in which to write the spreadsheet

    def __init__(self) -> None:
        """Initializes a ParseLogfiles object"""
        self._output_file: pd.ExcelWriter = self._open_output_file()  # Handle to spreadsheet
        self._log_info = OrderedDict([(col, []) for col in ('Date', 'App', 'App Version', 'User', 'Machine', 'Severity', 'Type', 'Exception', 'Exception Name', 'Message/Description', 'Server', 'Scripting Env', 'Script', 'Pt', 'Filename', 'Full Text')])  # To be converted to a Pandas DataFrame and written to spreadsheet
        self._parse_logfiles()

    def _open_output_file(self) -> None:
        """Creates and returns a handle to the output spreadsheet.

        Creates OUTPUT_PATH directory if it does not exist
        Spreadsheet name is 'Parsed Logfiles YYYY-MM-DD HH:MM:SS.xlsx'
        """
        # Create output directory if it doesn't exist
        if not os.path.isdir(self.OUTPUT_PATH):
            os.makedirs(self.OUTPUT_PATH)

        # Construct filename with timestamp
        dt = datetime.now().strftime('%Y-%m-%d %H_%M_%S')  # E.g., '2022-04-22 11_20_00'
        output_filename = os.path.join(self.OUTPUT_PATH, 'Parsed Logfiles ' + dt + '.xlsx')  # E.g., 'Parsed Logfiles 2022-04-22 11_20_00.xlsx'

        # Create and return ExcelWriter
        # openpyxl engine gives an invalid character error when writing to spreadsheet, but xlsxwriter works
        # Dates formatted in Excel as, e.g., '4/22/2022 11:21 AM'
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter', datetime_format='m/d/yyyy h:mm AM/PM')
        return writer

    def _extract_fld(self, event: str, regex: str) -> Union[str, np.float64]:
        """Helper function that extracts a field from the given text of a log event

        The event text is searched for a match to the regular expression, which must contain a single group representing the field

        Arguments
        ---------
        event: The text to search
        regex: The regular expression containing the desired field

        Returns
        -------
        The field value - the substring matching the first group in the regular expression
        If the text does not contain a matching substring, returns np.nan
        
        Examples
        --------
        self._extract_field('Yay! This is some text!\n', r'(.+)! This is some text') -> 'Yay!'
        self._extract_field('Here be some text.', r'(.+)! This is some text') -> np.nan
        """
        m = re.search(regex, event)
        if m is not None:
            return m.group(1)
        return np.nan

    def _parse_log(self, filename: str) -> None:
        """Adds field values from each event in a logfile, to the _logfile_info dictionary

        Add np.nan for any field that is not found
        Assumes that each event starts with a line containing a 'YYYY-MM-DD' date and that no line inside an event starts like this

        Arguments
        ---------
        filename: Absolute path to the logfile
        """
        with open(filename) as f:
            # Split the text into individual log events
            events = re.split(r'\n(?=\d{4}-\d{2}-\d{2})', f.read())

            # Extract app name
            simple_filename = os.path.split(filename)[1]  # Filename without path (e.g., 'RayStation_StorageTool_20190416222004.log')
            underscore_idx = simple_filename.rfind('_')  # Position of the rightmost underscore character in the filename
            app = simple_filename[:underscore_idx]  # Chop off everything after the rightmost underscore (e.g., 'RayStation_StorageTool')
            if app.endswith('_Errors'):  # Remove '_Errors' at the end of an app name (E.g., 'RayPhysics_Errors' -> 'RayPhysics')
                app = app[:-7]

            # Extract machine name
            machine = self._extract_fld(events[0], '#Machine: (.+)\n')

            # Add fields from each event
            # First event is header info, so ignore
            for event in events[1:]:
                event = event.strip()
                
                # Sometimes this comes at the end of the file, so ignore it
                if event == '':
                    continue

                # First line of an event is tab-separated info that applies to all event types
                flds = event.split('\n')[0]  # First line of event (e.g., '2019-04-16 22:20:04.655Z    13916    Verbose    General    ValidatePassword was called with username = DOMAIN\USER')
                dt, _, severity, type_, msg = re.split(r'\t|\s{2,}', flds)  # E.g., '2019-04-16 22:20:04.655Z', 'Verbose', 'General', 'ValidatePassword was called with username = DOMAIN\USER'
                
                # Convert datetime to Pandas datetime
                dt = pd.to_datetime(dt[:-1], format='%Y-%m-%d %H:%M:%S')

                # Extract exception name and description
                exn_name = self._extract_fld(event, '---\nName: (.+)\n')
                exn_desc = self._extract_fld(event, 'Description: (.+)\n')
                
                # Extract exception-specific info
                if 'Exception' in msg:
                    # Extract only the last part (most important info) from a long exception message
                    if '--->' in msg:
                        msg = msg.split('--->')[-1].strip()
                    
                    # Extract the name of the exception
                    exn = next(word for word in msg.split() if 'Exception' in word and word != 'Exception').strip(':')
                    
                    # If the message format is '<exception>: <other stuff>', extract '<other stuff>'
                    if ':' in msg:
                        colon_idx = msg.index(':')
                        msg = msg[(colon_idx + 1):].strip().strip(':')  # There may be an extra colon at the end, so remove it
                    
                    # Append the exception description to the message, if it is not np.nan
                    if isinstance(exn_desc, str):
                        msg += '\n' + exn_desc
                else:  # Not an exception
                    exn = np.nan
                
                # Extract more fields
                version = self._extract_fld(event, r'ApplicationVersion: (.+)\n')
                user = self._extract_fld(event, r'User: (.+)\n')
                server = self._extract_fld(event, r'Server: (.+)\n')
                env = self._extract_fld(msg, r'Using scripting environment \'(.+)\'')
                script = self._extract_fld(event, r'Running \'(.+)\' from database.\n')
                pt = self._extract_fld(event, r'Current patient is patient with name: (.+)\n')
                full_txt = event.replace('\t', ' ' * 4)  # For some reason, tabs are removed when writing to spreadsheet, so replace tabs with 4 spaces

                # A scripting log message is 'Using scripting environment ____'. Since we extract the env name into another field, no message is necessary for this event type
                if type_ == 'RayStation Scripting Log':
                    msg = np.nan
                else:
                    # Check for any 'Message' field in the event
                    more_msg = self._extract_fld(event, 'Message: (.+)\n')
                    if isinstance(more_msg, str):
                        msg += '\n' + more_msg

                # Add field values to dictionary
                self._log_info['Date'].append(dt)
                self._log_info['App'].append(app)
                self._log_info['App Version'].append(version)
                self._log_info['User'].append(user)
                self._log_info['Machine'].append(machine)
                self._log_info['Severity'].append(severity)
                self._log_info['Type'].append(type_)    
                self._log_info['Exception'].append(exn)
                self._log_info['Exception Name'].append(exn_name)  
                self._log_info['Message/Description'].append(msg)
                self._log_info['Server'].append(server)
                self._log_info['Scripting Env'].append(env)
                self._log_info['Script'].append(script)
                self._log_info['Pt'].append(pt)
                self._log_info['Filename'].append(filename)
                self._log_info['Full Text'].append(full_txt)

    def _write_output(self) -> None:
        """Writes logfile info to a table in an Excel spreadsheet

        Text in the 'Message/Description' column is wrapped
        All other columns except 'Full Text' are "autofit"
        """
        # Create DataFrame from dictionary, sorted by all columns that it makes sense to sort by
        log_sort_cols = [col for col in self._log_info if col != 'Message/Description']
        log_df = pd.DataFrame(self._log_info).sort_values(log_sort_cols, ignore_index=True)

        # Write data to spreadsheet
        log_df.to_excel(self._output_file, sheet_name='Logged Events', startrow=1, header=False, index=False)
        
        # If there are any non-header rows in the spreadsheet, format the data as an Excel table
        if len(log_df) > 0:
            # Add wrapped format to Excel workbook
            wb = self._output_file.book
            desc_fmt = wb.add_format({'text_wrap': True})

            # Create table from all spreadsheet data
            ws = self._output_file.sheets['Logged Events']
            max_row, max_col = log_df.shape
            col_settings = [{'header': col} for col in log_df.columns]
            ws.add_table(0, 0, max_row, max_col - 1, {'columns': col_settings})
            
            # Format columns: text wrapping and "autofit" widths
            for i, col in enumerate(log_df.columns):
                if col == 'Message/Description':
                    ws.set_column(i, i, 80, desc_fmt)  # Width 80 with wrapped text
                else:
                    if col == 'Date':
                        width = len('MM/DD/YYYY HH:MM AM')  # Maximum possible width of a date
                    elif col == 'Full Text':
                        width = 80  # Neither wrap nor "autofit"
                    else:
                        width = max(len(col) + 2, max(len(str(val)) + 1 for val in log_df[col].values))  # Max width of a value in that column (incl. the header), plus some padding (for the header, a little more to account for the filter arrow)
                    ws.set_column(i, i, width)

        self._output_file.save()

    def _parse_logfiles(self) -> None:
        """Parses logfiles in LOGFILE_PATH and writes selected fields to an Excel spreadsheet"""
        # Check all files in LOGFILE_PATH
        for filename in os.listdir(self.LOGFILE_PATH):
            if filename.endswith('.log') and not filename.endswith('.exn.log'):  # Parse .log, but not .exn.log (exception) files
                filename = os.path.join(self.LOGFILE_PATH, filename)  # Create absolute filepath to logfile
                self._parse_log(filename)
        self._write_output()


def parse_logfiles() -> None:
    """Creates a ParseLogFiles object"""
    ParseLogfiles()
