import os
import glob
import subprocess
from tkinter import Tk, filedialog

# List of file formats to convert
input_formats = ['.ppt', '.pptx']
output_format = '.pdf'

# Function to prompt user to select folder
def select_folder():
    # Create Tkinter root window
    root = Tk()
    root.withdraw()

    # Open file dialog to choose folder
    folder_selected = filedialog.askdirectory()

    # Destroy the root window
    root.destroy()

    return folder_selected

# Function to convert files
def convert_files(input_path, output_folder, max_quality):
    # Get the current filename and extension
    current_filename, current_extension = os.path.splitext(input_path)
    input_extension = current_extension.lower()

    # Only convert the file if it's in the list of input formats
    if input_extension in input_formats:
        # Set the output filename and path
        output_filename = f"{current_filename}{output_format}"
        output_path = os.path.join(output_folder, output_filename)

        # Use libreoffice to convert the document file
        cmd = ["libreoffice", "--convert-to", output_format[1:], "--outdir", output_folder, input_path]
        subprocess.call(cmd)

# Get the input and output folder paths from user
print("Please select the folder with the files to convert:")
input_folder = select_folder()
if not input_folder:
    print("No input folder selected. Exiting program!")
    exit()

print("Please select the output folder to save the converted files:")
output_folder = select_folder()
if not output_folder:
    print("No output folder selected. Exiting program!")
    exit()

# Process all files in the input folder
files = glob.glob(f"{input_folder}/*")
for file in files:
    convert_files(file, output_folder)

print("Conversion completed successfully!")

# softy_plug