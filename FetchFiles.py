import zipfile
import os
import csv

def search_files(dir_path, file_name):
    for root, dirs, files in os.walk(dir_path):
        if file_name in files:
            return os.path.join(root, file_name)
    return None

csv_path = "file.csv"
dir_path = "/Users/arijitghosh/Documents/Codes"
write_path = "/Users/arijitghosh/Downloads"

found_files = []

with open(csv_path, 'r') as csvfile:
    csv_reader = csv.reader(csvfile)
    next(csv_reader)
    
    for row in csv_reader:
        file_name = row[0]
        file_path = search_files(dir_path, file_name)
        
        if file_path:
            found_files.append(file_path)
            print("File found:", file_path)
        else:
            print("File not found:", file_name)

if found_files:
    zip_file_name = "found_files.zip"
    
    with zipfile.ZipFile(zip_file_name, mode='w') as zip_file:
        for file_path in found_files:
            zip_file.write(file_path)
            
    print(f"All found files have been zipped to {zip_file_name}")
else:
    print("No files found.")
