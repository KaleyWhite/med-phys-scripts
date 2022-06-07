import os
import tkinter as tk
from tkinter import filedialog
from typing import Callable, Optional

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from mpl_toolkits.axes_grid1 import make_axes_locatable
import numpy as np
from PIL import UnidentifiedImageError
from pydicom import dcmread
from pydicom.errors import InvalidDicomError
from pymedphys import gamma
from skimage.io import imread

from crmc_utils.constants import BG_COLOR, BTN_COLOR, TXT_COLOR, FONT
from crmc_utils.output_formatting import format_num


class FilmDicomGamma(tk.Frame):
    """Tkinter GUI that displays gamma analysis data for the user-selected film and DICOM files.
    
    Intended as a tool for IMRT QA. 
    Dose is measured at in the coronal direction on film inside of a cheese phantom. Film is centered at the center of the phantom.
    The measured dose is compared to the exported coronal dose plane of the center of the phantom in the IMRT QA plan in the TPS.
    
    Via the GUI, the user provides the inf-sup and right-left film dimensions and chooses the film (TIFF) and DICOM (.DCM) files to compare.
    The GUI then displays relative dose heatmaps for dose distributions, a gamma heatmap, and a gamma histogram.

    User inputs are validated, and the 'Analyze' button is disabled unless all input fields have valid values.
    """
    #: Dict[str, List[str]]: File types for tkfilebrowser
    filetypes = {
        'TIFF': ['*.tiff', '*.tif'],
        'DCM': ['*.dcm'],
    }

    #: Dict[str, Any]: Styling for any input field (Tk.Entry)
    input_ui_args = {'font': (FONT, 12)}

    #: Dict[str, Any]: Styling for the GUI in general
    ui_args = input_ui_args.copy()
    ui_args.update({'fg': TXT_COLOR, 'bg': BG_COLOR})

    #: Dict[str, Any]: Styling for a header label
    hdr_ui_args = ui_args.copy()
    hdr_ui_args['font'] = (FONT, 12, 'bold')

    #: Dict[str, Any]: Styling for a button
    btn_ui_args = ui_args.copy()
    btn_ui_args['bg'] = BTN_COLOR

    def __init__(self, parent: tk.Tk) -> None:
        """Inits a FilmDicomGamma object as a Tk Frame with user input fields and event handlers.""" 
        tk.Frame.__init__(self, parent)

        self.config(bg=BG_COLOR)

        #: Dict[str, Dict[str, Any]]: Association between input names and their Entrys, values (perhaps calculated), and Tk variables
        # {input_name : {'entry': tk.Entry, 'val': None if invalid input, otherwise the (calculated if necessary) input, 'var': variable bound to the Entry, other items specific to input type (e.g., 'minimum' for a number)}}
        self.inputs = {}

        # Add film dimension input
        self.add_lbl(10, 10, 'Film Dimensions (cm):', True)
        self.i_s = self.add_input('IS', 10, 40, 'Inf-Sup =', default=18)
        self.r_l = self.add_input('RL', 150, 40, 'Rt-Lt =', default=30)

        # Add file input
        self.add_lbl(10, 85, 'Files:', True)
        self.film = self.add_input('film', 10, 115, 'Film:', input_type='filepath', filetype='TIFF')
        self.dcm = self.add_input('DICOM', 10, 190, 'DICOM:', input_type='filepath', filetype='DCM')

        # Add 'Analyze' button
        self.analyze_btn = self.add_btn(10, 250, 'Analyze', self.analyze)
        self.analyze_btn['state'] = 'disabled'

    def add_lbl(self, x: float, y: float, txt: str, hdr: Optional[bool] = False) -> tk.Label:
        """Adds a label to the GUI.
        
        Args:
            x: The x-coordinate of the label.
            y: The y-coordinate of the label.
            txt: The text of the label.
            hdr: True if the label is a header, False otherwise.
            
        Returns:
            The new label.
        """
        args = self.hdr_ui_args if hdr else self.ui_args
        lbl = tk.Label(self, text=txt, **args)
        lbl.place(x=x, y=y)
        return lbl

    def add_btn(self, x: float, y: float, txt: str, cmd: Callable) -> tk.Button:
        """Adds a button to the GUI.
        
        Args:
            x: The x-coordinate of the button.
            y: The y-coordinate of the button.
            txt: The text of the button.
            cmd: The function to call when the button is clicked.
            
        Returns:
            The new button.
        """
        btn = tk.Button(self, text=txt, command=cmd, **self.btn_ui_args)
        btn.place(x=x, y=y)
        return btn

    def add_input(self, input_name: str, x: float, y: float, prompt: str, default: Optional[str] = '', input_type: Optional[str] = None, **kwargs) -> tk.Entry:
        """Adds a labeled input (tk.Entry) to the GUI.

        You may think of an 'input' as a group of an Entry, a Var, the (perhaps calculated) value from that variable, and any other pertinent info.
        Such 'other pertinent info' are keyword args.
        These associations are kept in the `inputs` attribute.
        
        Args:
            input_name: The name of the input.
                This is only used to associate an entry with its value and bound variable.
            x: The x-coordinate of the input.
            y: The y-coordinate of the input.
            prompt: The text of the label.
            default: The default value for the input.
            input_type: The type of input.
                Must be 'number' or 'filepath'.
                If not provided, is inferred from the type of `default` if it exists.
        
        Keyword Args:
            minimum (float): The minimum valid value for a 'number' input.
            maximum (float): The maximum valid value for a 'number' input. 
            filetype (str): For 'filepath' inputs, the desired type of file.
                Must be a key in self.filetypes.
            
        Returns:
            The new entry.
        """
        # Determine input type if not provided
        if input_type is None:
            if isinstance(default, int) or isinstance(default, float):
                input_type = 'number'
            elif isinstance(default, str):
                input_type = 'filepath'

        # Add prompt
        self.add_lbl(x, y, prompt)
        
        # Create Entry and update self.inputs
        var = tk.StringVar(self, value=str(default), name=input_name)
        args = {'textvariable': var}
        args.update(self.input_ui_args)
        if input_type == 'number':
            args['width'] = 5
        elif input_type == 'filepath':
            args['width'] = 50
            x += 40
        entry = tk.Entry(self, **args)  
        entry.place(x=x + 40, y=y)
        self.inputs[input_name] = {'entry': entry, 'var': var, 'input_type': input_type}
        self.inputs[input_name]['val'] = None if default == '' else var.get() if input_type == 'filepath' else float(var.get())
        self.inputs[input_name].update(kwargs)
        var.trace('w', lambda nm, idx, mode, name=input_name: self.validate_input(name))  # When the value in the Entry is changed, validate it with self.validate_input

        # For filepath input, add 'Browse' button
        if input_type == 'filepath':
            filetype = kwargs.get('filetype')
            self.inputs[input_name]['filetype'] = filetype
            self.add_btn(x + 500, y, 'Browse', cmd=lambda: self.choose_file(input_name))  # When the browse button is clicked, allow user to choose a file

        return var

    def choose_file(self, input_name: str):
        """Allows user to choose a filename from a file dialog.

        Args:
            input_name: The name of the input to choose the file for.
        """
        # Get input's info by its name
        entry_info = self.inputs[input_name]
        filetype = entry_info['filetype']

        # Set up `filetypes` argument for file dialog
        filetypes = []
        if filetype is not None:
            filetypes = [(filetype, ext) for ext in self.filetypes[filetype]]
        filetypes.append(('All Files', '*.*'))
  
        # Get filepath and update the variable
        filepath = filedialog.askopenfilename(title='Choose File', filetypes=filetypes)
        entry_info['var'].set(filepath)

    def validate_input(self, input_name: str) -> None:
        """Validates user input.

        If input is invalid according to restrictions in `self.inputs`, Entry background turns red.
        Enables the 'Analyze' button if any inputs are invalid.

        Args:
            input_name: The name of the input to validate.
        """
        # Get input's info by its name
        entry_info = self.inputs[input_name]
        entry = entry_info['entry']
        var = entry_info['var']
        val = entry_info['val']
        input_type = entry_info['input_type']
        new_val = var.get()

        # Validate input according to `entry_info`
        if input_type == 'number':
            # Number must be a valid number and, if applicable, not outside the 'minimum'/'maximum' bounds
            valid = True
            try:
                val = float(new_val)
                if ('minimum' in entry_info and val < entry_info['minimum']) or ('maximum' in entry_info and val > entry_info['maximum']):
                    valid = False
            except ValueError:
                valid = False
        elif input_type == 'filepath':
            valid = os.path.isfile(new_val)  # Filepath must exist
            if valid:
                filetype = entry_info.get('filetype')
                if filetype == 'TIFF':
                    # Film (TIFF) file must be a valid image file
                    try:
                        val = imread(new_val)
                    except (ValueError, UnidentifiedImageError):
                        valid = False
                elif filetype == 'DCM':
                    # DICOM file must be a valid DICOM file and must be a dose file
                    try:
                        val = dcmread(new_val)
                        valid = val.Modality == 'RTDOSE'  
                    except (InvalidDicomError, KeyError):
                        valid = False
        
        # Set entry background color and val in `entry_info`
        if valid:
            var.set(new_val)
            entry_info['val'] = val
            entry.config(bg='white')
        else:
            entry_info['val'] = None
            entry.config(bg='red')
        
        # Enable or disbale 'Analyze' button
        self.analyze_btn['state'] = 'disabled' if any(entry_info['val'] is None for entry_info in self.inputs.values()) else 'normal'

    def analyze(self) -> None:
        """Performs gamma analysis on the film and DICOM data and displays the results."""
        # Get user input values
        i_s = self.inputs['IS']['val']
        r_l = self.inputs['RL']['val']
        film = self.inputs['film']['val']
        dcm = self.inputs['DICOM']['val']

        # Compute film axes
        film_w, film_h = film.shape
        film_res_x = round(float(i_s) * 10 / film_w, 1)  # Convert cm -> mm
        film_res_y = round(float(r_l) * 10 / film_h, 1)
        film_axes = (np.arange(0, film_w) * film_res_x, np.arange(0, film_h) * film_res_y)

        # Compute DICOM axes
        dcm_res_x, dcm_res_y = map(float, dcm.PixelSpacing)
        dcm = dcm.pixel_array
        dcm_w, dcm_h = dcm.shape
        dcm_axes = (np.arange(0, dcm_w) * dcm_res_x, np.arange(0, dcm_h) * dcm_res_y)

        # Scale film data to DICOM data
        film = np.multiply(film, np.max(dcm) / np.max(film))

        # Convert to relative dose
        film = np.multiply(film, 100 / np.max(film))
        dcm = np.multiply(dcm, 100 / np.max(dcm))

        # Compute gamma
        gamma_options = {
            'dose_percent_threshold': 3,
            'distance_mm_threshold': 2,
            'lower_percent_dose_cutoff': 20,
            'interp_fraction': 10,  # Should be 10 or more for more accurate results
            'max_gamma': 2,
            'random_subset': None,
            'ram_available': 2**29,  # 1/2 GB
            'quiet': True
        }
        
        gam = gamma(dcm_axes, dcm, film_axes, film, **gamma_options)
        valid_gamma = gam[~np.isnan(gam)]

        # Add label with pass information
        pass_ratio = np.sum(valid_gamma <= 1) / len(valid_gamma)
        results_txt = f"{gamma_options['dose_percent_threshold']}%, {gamma_options['distance_mm_threshold']} mm\n{format_num(pass_ratio * 100, 0.01)}% Pass"
        self.add_lbl(10, 315, results_txt)

        # Add plots

        plt.ioff()  # Prevent new figure window from opening
        fig, ((film_ax, dcm_ax), (raw_ax, hist_ax)) = plt.subplots(nrows=2, ncols=2, figsize=(6, 6))
        fig.patch.set_facecolor(BG_COLOR)

        # Add film dose heatmap
        self.im_show(fig, film_ax, film, 'Film', 'Rel. Dose (%)')
        
        # Add DICOM dose heatmap
        self.im_show(fig, dcm_ax, dcm, 'DICOM', 'Rel. Dose (%)')

        # Add gamma heatmap
        self.im_show(fig, raw_ax, gam, 'γ', legend_lim=(0, 2))

        # Add gamma histogram
        # Code modified from PyMedPhys docs
        num_bins = gamma_options['interp_fraction'] * gamma_options['max_gamma']
        bins = np.linspace(0, gamma_options['max_gamma'], num_bins + 1)
        hist_ax.set_aspect('auto')
        hist_ax.hist(valid_gamma, bins, density=True, color='gray')
        hist_ax.set_xlim(0, gamma_options['max_gamma'])
        hist_ax.set_title('γ')
        hist_ax.yaxis.set_visible(False)

        plt.subplots_adjust(left=0, bottom=0.1, right=0.9, top=0.97, wspace=0.2, hspace=0.1)

        # Add plots to GUI
        canvas = FigureCanvasTkAgg(fig, self)
        canvas.get_tk_widget().place(x=10, y=370)

    def im_show(self, fig: plt.Figure, ax: plt.Axes, data: np.ndarray, title: str, legend_lbl: Optional[str] = '', legend_lim: Optional[Tuple(float, float)] = None):
        """Creates a rainbow-colored heatmap of the data.

        A labeled legend is included.

        Args:
            fig: The Figure to add the legend to.
            ax: The axis to add the hetmap to.
            data: The data to plot on the heatmap.
            title: The title for the plot.
            legend_lbl: The label for the legend.
            legend_lim: The min and max values to display on the legend.
        """
        # Plot the heatmap
        im = ax.imshow(data, cmap='rainbow')
        ax.set_title(title)
        ax.tick_params(top=False, bottom=False, left=False, right=False, labelleft=False, labelbottom=False)  # Hide x- and y-axes
        divider = make_axes_locatable(ax)

        # Add legend
        cax = divider.append_axes('right', size='13%', pad=0.1)
        cbar = fig.colorbar(im, cax=cax, orientation='vertical')
        cbar.set_ticks([])
        cax.set_ylabel(legend_lbl, rotation=270)
        cax.get_yaxis().labelpad = 15

        # Legend data labels
        cax_xlim = cax.get_xlim()
        cax_xmid = (cax_xlim[1] + cax_xlim[0]) / 2

        cax_ylim = cax.get_ylim()
        cax_yrange = cax_ylim[1] - cax_ylim[0]

        if legend_lim is None:
            legend_lim = cax.get_ylim()
        for i, pt in enumerate(np.linspace(*legend_lim, 6)):
            cax.text(cax_xmid, cax_ylim[0] + cax_yrange / 5 * i, str(round(pt, 1)), ha='center', va='center')


def film_dicom_gamma():
    """Displays a FilmDicomGamma GUI."""
    root = tk.Tk()
    root.geometry(f'{root.winfo_screenwidth()}x{root.winfo_screenheight()}')  # Take up entire monitor
    root.resizable(False, False)
    root.title('Film & DICOM γ Analysis')
    root.config(bg=BG_COLOR)
    FilmDicomGamma(root).pack(fill='both', expand=True)
    root.mainloop()


if __name__ == '__main__':
    film_dicom_gamma()
