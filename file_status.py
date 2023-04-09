import os
from collections import defaultdict
from datetime import datetime
from openpyxl import Workbook
from prettytable import PrettyTable
import subprocess

# Set the path to the directory to be scanned
path = "."

# Get the current folder name
folder_name = os.path.basename(path)

# Get the current date and time
now = datetime.now()
date_string = now.strftime("%Y-%m-%d_%H-%M-%S")

# Initialize dictionaries to store counts of subfolders and file extensions
subfolder_count = 0
file_extension_count = defaultdict(int)

# Loop through all the files and subfolders in the specified directory
for root, dirs, files in os.walk(path):
    # Increment the subfolder count for each subfolder
    subfolder_count += len(dirs)

    # Increment the file extension count for each file
    for name in files:
        file_extension = os.path.splitext(name)[1]
        file_extension_count[file_extension] += 1

# Create a table to display the counts of subfolders and file extensions
table = PrettyTable()
table.field_names = ["Item", "Count"]
table.align["Count"] = "r"

# Add the subfolder count to the table
table.add_row(["Subfolders", subfolder_count])

# Add the file extension counts to the table
for extension, count in sorted(file_extension_count.items(), key=lambda x: x[1], reverse=True):
    table.add_row([f"File extension: {extension}", count])

# Print the table
print(table)

# Export the table as a text file
txt_file_name = f"{folder_name}-{date_string}.txt"
with open(txt_file_name, "w", encoding="utf-8", errors="ignore") as f:
    f.write(str(table))

print(f"\nReport exported to {txt_file_name}")

# Print the table to the console
print(table)

# Open the exported Excel file
subprocess.run(["start", txt_file_name], shell=True)
