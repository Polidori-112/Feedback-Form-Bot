import os
import subprocess

# specify the directory to search for text files
directory_path = "emails"

# find all text files in the directory
text_files = [f for f in os.listdir(directory_path) if f.endswith(".txt")]

# loop over each text file and run grep on its filename
for file_name in text_files:
    # remove the .txt extension from the filename
    file_base_name = os.path.splitext(file_name)[0]
    
    # print a message indicating which file is being checked
    print(f"Checking {file_base_name}:")
    
    # run grep on the filename with case-insensitive search
    command = "grep" + " -il " + file_base_name + " " + os.path.join(directory_path, "*")
    grep_output = subprocess.check_output(command, shell=True, universal_newlines=True)
    
    # print the grep output
    print(grep_output)
