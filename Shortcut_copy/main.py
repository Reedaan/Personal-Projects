import os
import shutil

# Set the paths to the input and output folders
input_folder = r'C:\Users\MarekM\Desktop\input folder'
output_folder = r'C:\Users\MarekM\Desktop\output folder'

# Iterate through all the files in the input folder
for filename in os.listdir(input_folder):
    # Check if the file is a shortcut
    if filename.endswith('.lnk'):
        shortcut_path = os.path.join(input_folder, filename)
        
        if os.path.exists(shortcut_path):
            # Use the win32com.client library to get the target path
            from win32com.client import Dispatch
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            target_path = shortcut.Targetpath

            # Extract the filename from the target path
            filename = os.path.basename(target_path)

            # Construct the path to the output file
            output_path = os.path.join(output_folder, filename)

            # Create a copy of the target file in the output folder
            shutil.copy2(target_path, output_path)

            print(f'File {filename} has been copied to {output_path}')
        else:
            print(f'{shortcut_path} does not exist.')
