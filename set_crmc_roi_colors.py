from datetime import datetime
from functools import partial
import os
import random
from typing import List, Tuple
import webbrowser

import pandas as pd

from tkinter import *


EXCEL_FILETYPES = (('Excel files', '*.' + ext) for ext in ('csv', 'xls', 'xlsx', 'xlsm'))  # `filetypes` argument for file dialogs

tg263_data = None
TG263_DATA_DIR = os.path.join('T:', os.sep, 'Physics', 'KW', 'med-phys-spreadsheets')  # Absolute path to starting directory for file dialogs


class HyperlinkManager(object):
    """Class that manages clickable links in a Tkinter Text widget
    
    Modified from StackOverflow: https://stackoverflow.com/questions/50327234/adding-link-to-text-in-text-widget-in-tkinter
    """
    def __init__(self, text: str) -> None:
        """Initializes a HyperlinkManager object
        
        Arguments
        text: The text of the hyperlink
        """
        self.text = text
        self.text.tag_config('hyper', foreground='blue', underline=1)  # Blue, underlined link text
        self.text.tag_bind('hyper', '<Enter>', self._enter)  # When mouse hovers over the link
        self.text.tag_bind('hyper', '<Leave>', self._leave)  # When mouse leaves the link
        self.text.tag_bind('hyper', '<Button-1>', self._click)  # When left mouse clicks the link
        self.reset()  # Clear all links

    def reset(self) -> None:
        """Clears all links in this link manager"""
        self.links = {}

    def add(self, action: partial) -> Tuple[str, str]:
        """Adds an action to this link manager 

        Arguments
        ---------
        action:

        Returns
        -------
        Tag to use in an associated text widget
        """
        tag = 'hyper-%d' % len(self.links)  # E.g., 'hyper-2' if there are 2 links
        self.links[tag] = action  # Associate link tag with action
        return 'hyper', tag

    def _enter(self, event: partial) -> None:
        """Handles the mouse enter event for the link manager

        Changes the cursor to a hand
        """
        self.text.config(cursor='hand2')

    def _leave(self, event: partial) -> None:
        """Handles the mouse enter event for the link manager

        Changes the cursor to the default
        """
        self.text.config(cursor='')

    def _click(self, event: partial) -> None:
        """Handles a widget's click event

        If this is a link widget, calls the action associated with the link
        """
        for tag in self.text.tag_names(CURRENT):
            if tag[:6] == 'hyper-':
                self.links[tag]()
                return


def rand_colors(num_colors: int) -> List[str]:
    """Generates a list of unique random colors in the format '(A, R, G, B)', where A is always 255 and R, G, and B are integers between 0 and 255, inclusive

    Arguments
    ---------
    num_colors: The number of random colors to generate
                Must be positive and not greater than the maximum possible number of random colors

    Returns
    -------
    A list of the randomly generated colors

    Raises
    ------
    ValueError: If the number of colors is zero or negative, or if it is greater than the number of possible unique colors in the space
    """
    # Validate `num_colors`
    if num_colors <= 0:
        raise ValueError('Number of colors must be positive')
    max_num_colors = 255 * 255 * 255
    if num_colors > max_num_colors:
        raise ValueError('Can only generate ' + str(max_num_colors) + ' unique colors')
    
    colors = []
    while len(colors) < num_colors:
        color = f'(255, {random.randint(0, 255)}, {random.randint(0, 255)}, {random.randint(0, 255)})'
        if color not in colors:  # Ignore non-unique colors
            colors.append(color)
    return colors


def choose_tg263_filepath(tg263_filepath_lbl: Text, tg263_filepath_btn: Button, colors_filepath_btn: Button, write_btn: Button) -> None:
    """Gets TG-263 nomenclature data from the user-sepcified spreadsheet
    
    Arguments
    ---------
    tg263_filepath_label: The label containing the user-selected filepath to the original TG-263 spreadsheet
    tg263_filepath_btn: The button that the user clicks to choose the original TG-263 spreadsheet
    colors_filepath_btn: The button that the user clicks to choose the spreadsheet of non-random colors
    write_btn: The button that the user clicks to write the new data (with the "Colors" column) to a file
    """
    global tg263_data

    # Get filepath from user
    filepath = ''
    while not filepath:
        filepath = filedialog.askopenfilename(initialdir=TG263_DATA_DIR, title='Choose TG-263 Spreadsheet', filetypes=EXCEL_FILETYPES)
    
    tg263_filepath_lbl.config(text=filepath)  # Show filepath in GUI
    
    tg263_data = pd.read_excel(filepath)
    tg263_data = tg263_data.loc[:, ~tg263_data.columns.str.contains('^Unnamed')]  # Remove extraneous column
    tg263_data.drop('Color', axis=1, errors='ignore', inplace=True)
    if 'TG263-Primary Name'in tg263_data.columns:
        tg263_data.rename(columns={'TG263-Primary Name': 'TG-263 Primary Name'}, inplace=True)  # Make column name easier to work with

    # Standardize inconsistent column names
    tg263_data.loc[tg263_data['Target Type'] == 'Non_Anatomic', 'Target Type'] = 'Non-Anatomic' 
    tg263_data.loc[tg263_data['Anatomic Group'] == 'Limbs', 'Anatomic Group'] = 'Limb'
    # Correct obvious typos
    tg263_data.loc[tg263_data['TG-263 Primary Name'] == 'Colon_PTVxx', ['TG-263 Primary Name', 'TG-263-Reverse Order Name', 'Description']] = ['Colon_PRVxx', 'PRVxx_Colon', 'PRV created with xx mm expansion on the colon']
    tg263_data.loc[tg263_data['TG-263 Primary Name'] == 'LN_lliac_Int_R', 'TG-263 Primary Name'] = 'LN_Iliac_Int_R'
    
    # Set colors
    print('?')
    tg263_filepath_btn.config(state=DISABLED)  # Disable choose filepath button after file has been chosen
    tg263_data['Color'] = rand_colors(len(tg263_data))  # Populate random colors column
    colors_filepath_btn.config(state=NORMAL)  # Enable choose colors filepath button
    write_btn.config(state=NORMAL)  # Enable write colors filepath button


def choose_colors_filepath(colors_filepath_lbl: Text, colors_filepath_btn: Button) -> None:
    """Adds non-random colors to the `tg_263_data` DataFrame based on the user-chosen spreadsheet

    Disables the button that allows choosing a colors filepath

    Arguments
    ---------
    colors_filepath: The label containing the filepath of the user-selected spreadsheet of non-random colors
    colors_filepath_btn: The button that the user clicks to choose the spreadsheet of non-random colors
    """
    global tg263_data

    # Get colors filepath
    colors_filepath_btn.config(state=DISABLED)
    filepath = ''
    while not filepath:
        filepath = filedialog.askopenfilename(initialdir=TG263_DATA_DIR, title='Choose Colors Spreadsheet', filetypes=EXCEL_FILETYPES)#, filetypes=(('Excel files', '*.csv'), ('Excel files', '*.xls'), ('Excel files', '*.xlsx')))
    colors_filepath_lbl.config(text=filepath, state=DISABLED)

    # Read in colors data
    colors_data = pd.read_excel(filepath, usecols=['TG-263 Primary Name', 'Color'])
    color_regex = r'\(' + r', '.join([r'([01]?\d\d?|2[0-4]\d|25[0-5])'] * 4) + r'\)'  # Matches a color in "(A, R, G, B)" format
    colors_data = colors_data.loc[colors_data['Color'].str.match(color_regex)]  # Select rows with a valid color
    
    # Add "Colors" column to `tg263_data`
    # Populate this column with the colors for names in the colors DataFrame
    tg263_data = tg263_data.merge(colors_data, on='TG-263 Primary Name', how='left')
    tg263_data['Color'] = tg263_data['Color_y'].fillna(tg263_data['Color_x'])
    tg263_data.drop(['Color_x', 'Color_y'], axis=1, inplace=True)


def write_to_file(tg263_filepath_lbl: Text, colors_filepath_btn: Button, write_btn: Button, successful_write_txt: Text) -> None:
    """Event handler for clicking the "write to new file" button
    
    Writes the new data (with the "Colors" column) to an XLSX file called "TG-263 Nomenclature with CRMC Conventional Colors YYYY-MM-DD HH:MM:SS.xlsx"
    Disables the buttons to choose a colors filepath and write to file (the TG-263 filepath is already disabled)

    Arguments
    ---------
    tg263_filepath_label: The label containing the user-selected filepath to the original TG-263 spreadsheet
    colors_filepath_btn: The button that the user clicks to choose the spreadsheet of non-random colors
    write_btn: The button that the user clicks to invoke this function
    successful_write_txt: The label that displays the "success" message after the new spreadsheet is created
    """
    # Disable buttons
    colors_filepath_btn.config(state=DISABLED)
    write_btn.config(state=DISABLED)
    
    # Create filepath for new XLSX file
    tg263_filepath = tg263_filepath_lbl.cget('text')
    new_tg263_filename = f'TG-263 Nomenclature with CRMC Conventional Colors {datetime.now().strftime('%Y-%m-%d %H_%M_%S')}.xlsx'
    new_tg263_filepath = os.path.join(os.path.dirname(tg263_filepath), new_tg263_filename)

    # Write to file
    writer = pd.ExcelWriter(new_tg263_filepath)
    tg263_data.to_excel(writer, startrow=1, header=False, index=False)  # Do not write row or column headers
    wb = writer.book
    ws = writer.sheets['Sheet1']
    (max_row, max_col) = tg263_data.shape
    col_settings = [{'header': col} for col in tg263_data.columns]
    ws.add_table(0, 0, max_row, max_col - 1, {'columns': col_settings})
    writer.save()

    # Display "success" message
    successful_write_txt.config(state=NORMAL)  # Unhide the label
    successful_write_txt.insert(END, 'New file written successfully: ')
    successful_write_txt.insert(END, new_tg263_filename, HyperlinkManager(successful_write_txt).add(partial(webbrowser.open_new_tab, new_tg263_filepath)))
    successful_write_txt.config(state=DISABLED)  # Disable the label so it can't have events


def set_crmc_roi_colors() -> None:
    """Sets CRMC conventional colors for a spreadsheet of TG-263 standard ROI names.

    The user chooses the TG-263 spreadsheet form a GUI, which includes the option to download the original xls file from the AAPM website.
    If the original spreadsheet conatins a "Color" column, this column is dropped.
    A color is randomly generated for each ROI (TG-263 Primary Name) in the spreadsheet. The color format is (A, R, G, B), where A is always 255.
    The user also optionally chooses a spreadsheet of names and colors that should not be randomly generated. This spreadsheet should contain two columns: "TG-263 Primary Name" and "Color".
    Until the original spreadsheet is chosen, the user is disabled from choosing the colors spreadsheet or writing the new data to a file.
    The new data, which is the same as the original except with the added "Color" column, is written to a spreadsheet called "TG-263 Names with CRMC Conventional Colors <timestamp>.xlsx" in the same directory as the original spreadsheet.
    """
    root = Tk()  # Main Tk (GUI) window

    # Fonts for GUI
    reg_font = font.Font(family='Helvetica', size=10)
    bold_font = font.Font(family='Helvetica', size=10, weight=font.BOLD)

    # Function call for when the user clicks the hyperlink to download the TG-263 spreadsheet from the AAPM website
    url = 'https://www.aapm.org/pubs/reports/RPT_263_Supplemental/TG263_Nomenclature_Worksheet_20170815.xls'
    download_tg263_spreadsheet = partial(webbrowser.open_new_tab, url)

    row = 0  # Keep track of position for next widget in the GUI

    # Format the GUI
    root.configure(bg='white')
    root.geometry('625x360')
    root.title('Set CRMC Conventional Colors for TG-263 Structure Names')

    # User instructions
    txt = Text(root, background='white', bd=0, font=reg_font, padx=10, pady=5, wrap=WORD)
    link = HyperlinkManager(txt)  # The text of the Text widget will have hyperlinks
    txt.insert(END, 'This program adds randomly generated (A, R, G, B) colors to the ')
    txt.insert(END, 'spreadsheet', link.add(download_tg263_spreadsheet))  # Clickable hyperlink
    txt.insert(END, ' downloaded from the AAPM website. The TG-263 data from the spreadsheet is written to a new spreadsheet in the same directory as the original, with the additional "Colors" column. All A = 255.')
    txt.grid(row=row, ipadx=10, ipady=10, sticky=W)  # Position the Text widget in the GUI
    txt.config(height=txt.index(END), state=DISABLED) # Disable the Text widget so it can't have events
    row += 1

    # Controls to choose TG-263 spreadsheet
    Label(root, bg='white', font=bold_font, padx=10, pady=5, text='TG-263 Spreadsheet').grid(row=row, sticky=W)  # Position bold "TG-263 Spreadsheet" text in GUI
    row += 1
    tg263_filepath_lbl = Label(root, bg='white', font=reg_font, padx=10, pady=5, text='[No file chosen]')  # Before a file is chosen, the filepath has placeholder "[No file chosen]"
    tg263_filepath_lbl.grid(row=row, sticky=W)
    tg263_filepath_btn = Button(root, text='Choose file')
    tg263_filepath_btn.grid(row=row, sticky=E)
    row += 1

    # Controls to choose non-random colors spreadsheet
    
    Label(root, bg='white', font=bold_font, padx=10, pady=5, text='Non-Random Colors Spreadsheet').grid(row=row, sticky=W)   # Position bold "Non-Random Colors Spreadsheet" text in GUI
    row += 1
    # Instructions for choosing the non-random colors spreadsheet
    txt = Text(root, background='white', bd=0, font=reg_font, padx=10, pady=5, wrap=WORD)
    link = HyperlinkManager(txt)  # The text of the Text widget will have hyperlinks
    txt.insert(END, 'In a spreadsheet, specify any colors that you don\'t want to be randomly generated. The spreadsheet should have two columns:\n-  Name: Name from the ')
    txt.insert(END, 'TG-263 spreadsheet', link.add(download_tg263_spreadsheet))  # Clickable hyperlink
    txt.insert(END, '. Names not in the ')
    txt.insert(END, 'TG-263 spreadsheet', link.add(download_tg263_spreadsheet))
    txt.insert(END, ' will be ignored.\n-  Color: The color you want that structure to be, in (A, R, G, B) format. E.g., (255, 255, 255, 255) for white. Colors not in this format will be ignored and randomly generated.')
    txt.grid(row=row, ipadx=10, ipady=20, sticky=W)
    txt.config(height=txt.index(END), state=DISABLED)
    row += 1
    # Label to display filepath of non-random colors spreadsheet
    colors_filepath_lbl = Label(root, bg='white', font=reg_font, padx=10, pady=5, text='[No file chosen]')
    colors_filepath_lbl.grid(row=row, sticky=W)
    colors_filepath_btn = Button(root, state=DISABLED, text='Choose file')
    colors_filepath_btn.grid(row=row, sticky=E)
    row += 1

    # Button to click to write the updated data (with the added "Colors" column) to a new spreadsheet
    write_btn = Button(root, font=reg_font, state=DISABLED, text='Write to new file')
    write_btn.grid(row=row, sticky=E+W)
    row += 1

    # Label to display "success" message after the new spreadsheet is created
    successful_write_txt = Text(root, background='white', bd=0, font=reg_font, padx=10, pady=5, wrap=WORD)
    successful_write_txt.grid(row=row, sticky=W)
    successful_write_txt.config(state=DISABLED)

    # Event handlers for button clicks
    tg263_filepath_btn.config(command=partial(choose_tg263_filepath, tg263_filepath_lbl, tg263_filepath_btn, colors_filepath_btn, write_btn))
    colors_filepath_btn.config(command=partial(choose_colors_filepath, colors_filepath_lbl, colors_filepath_btn))
    write_btn.config(command=partial(write_to_file, tg263_filepath_lbl, colors_filepath_btn, write_btn, successful_write_txt))

    root.mainloop()  # Launch GUI


if __name__ == '__main__':
    set_crmc_roi_colors()
    