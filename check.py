#prints out the names of people who have not completed the form;

import os
import glob

print("The people who have not completed the form are: ")

# specify the directory to search for text files
directory_path = "emails"

if not os.listdir(directory_path):
    print(f"Make sure that you run the Feedback.py file before checking if everyone has filled out the form. Please run that, then try running this script again.")
    exit(0)

# find all text files in the directory
text_files = glob.glob(os.path.join(directory_path, "*.txt"))

# loop over each text file and check if it starts with 'Good Morning'
for file_path in text_files:
    with open(file_path, "r") as f:
        first_line = f.readline().strip()
        if not first_line.startswith("Good Morning"):
            print(file_path[7:-4])

print("If no names show up than everyone has completed it.\n\nIf you see a None, Na, etc., somebody did not fill out the form correctly and the email program may crash. Either delete that file in the emails directory or change the cell in the spreadsheet.\n\nIf you see someone's name spelled incorrectly, that means that someone filled out the form incorrectly. Please edit the downloaded spreadsheet by finding the mispelled name (ctrl-f), then spelling it correctly. This will ensure that the MIDN has received all feedback that was given to them")
