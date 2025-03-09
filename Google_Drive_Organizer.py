"""
Script to organize files in Google Drive based on modification date and client name.

Description:
- Mounts Google Drive.
- Organizes files from a source folder into a folder structure by month and client.
- Extracts the client name from the file name using " - " as a separator.
- Creates the necessary folders and moves the files to their corresponding destination.

Libraries used:
- os: File and directory management.
- shutil: File movement.
- re: Regular expressions for string manipulation.
- google.colab.drive: Connect Google Drive in Google Colab.
- datetime: Date and time management.
"""

import os
import shutil
import re
from google.colab import drive
from datetime import datetime

# Connect Google Drive in Google Colab
drive.mount('/content/drive')

# Define source and destination folders
origin_folder = ''  # Path to the source folder
destination_folder = ''  # Path to the destination folder

# Check if the destination folder exists and create it if necessary
os.makedirs(destination_folder, exist_ok=True)

# Iterate through files in the source folder
for filename in os.listdir(origin_folder):
    file_path = os.path.join(origin_folder, filename)

    # Check if it is a file
    if os.path.isfile(file_path):
        # Get the file's modification date
        mod_time = os.path.getmtime(file_path)
        date_folder = datetime.fromtimestamp(mod_time).strftime('%Y %m %B')

        # Extract the client name by splitting on " - " and removing the file extension
        parts = filename.rsplit(" - ", 1)
        client_name = parts[1].split("(")[0].strip() if len(parts) > 1 else "Unknown"
        client_name = os.path.splitext(client_name)[0].title()

        # Create organization folders
        month_folder = os.path.join(destination_folder, date_folder)
        client_folder = os.path.join(month_folder, client_name)
        print(client_folder)

        os.makedirs(client_folder, exist_ok=True)  # Create folder if it does not exist

        # Move the file to the corresponding folder
        shutil.move(file_path, os.path.join(client_folder, filename))
