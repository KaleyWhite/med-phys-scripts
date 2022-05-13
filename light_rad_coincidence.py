"""Performs the Winston-Lutz light-radiation coincidence test

This is truly the poor man's QA software. To perform the test:
1. Irradiate square fields on a piece of film.
2. Scan the film into the computer as a PDF.
3. Run this script. The user chooses the PDF file.
4. The script converts the PDF to a 300dpi image, reads in the image, and annotates the square fields. The field is outlined and labeled with its cm dimensions.
5. The script displays the annotated image.
"""
import os
import pip
from tkinter import filedialog, Tk
from typing import Tuple

for pkg in ['cv2', 'matplotlib', 'pdf2image']:
    try:
        __import__(pkg)
    except ImportError:
        pip.main(['install', package])

import matplotlib.pyplot as plt


QA_FOLDER = os.path.join('T:', os.sep, 'Physics', 'QA & Procedures')  # Absolute path to the starting directory for the file dialog


def px_to_cm(*args: float) -> Tuple[float, ...]:
    """Converts each argument, in pixels, to centimeters

    Assumes 300dpi. Definitely works with letter-sized film scanned as a PDF on our printer

    Arguments
    ---------
    *args: Variable number of pixel values

    Returns
    -------
    A tuple of the px values converted to cm
    """
    return (dim / 300 * 2.54 for dim in args)


def main():
    # Get filename using tkinter
    Tk().withdraw()  # Hide main Tk window
    filename = ''
    while not filename:  # Prompt for filename until the user provides it
        filename = filedialog.askopenfilename(filetypes=[('PDF', '*.pdf')], initialdir=QA_FOLDER, title='Select film PDF')
    
    # Convert PDF to image (JPEG)
    img = pdf2image.convert_from_path(filename, dpi=300)[0]
    img_filename = os.path.splitext(filename)[0] + '.jpg'
    img.save(img_filename, 'JPEG')

    # Read in and preprocess image for OpenCV
    img = cv2.imread(img_filename)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)  # Convert to grayscale (GaussianBlur requires)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)  # Reduce nois with a Gaussian filter
    thresh = cv2.threshold(blurred, 60, 255, cv2.THRESH_BINARY_INV)[1]  # Set any pixels >60, to white

    # Find and annotate contours
    cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)[0]  # List of contours (each contour is a list of [x, y] coordinates)
    for c in cnts:  # Process each contour
        x, y, w, h = cv2.boundingRect(c)  # Top-left coordinates, and size (px) of the contour
        w, h = px_to_cm(w, h)  # Convert px to cm
        # Ignore very small contours. Assume we'll never test smaller than a 1x1 field
        if w < 1 or h < 1:
            continue
        cv2.drawContours(img, [c], -1, (255, 0, 0), 5)  # Outline the field
        cv2.putText(img, f'{w:.2f}x{h:.2f}', (x, y - 20), cv2.FONT_HERSHEY_SIMPLEX, 2.5, (255, 0, 0), 5)  # Write dimensions ("wxh"), rounded to 2 decimal places, above the contour

    # Display image
    plt.imshow(img)
    plt.axis('off')
    plt.show()


if __name__ == '__main__':
    main()
