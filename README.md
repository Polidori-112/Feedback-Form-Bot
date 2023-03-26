##README

This is a script written by MIDN Polidori and chatgpt to help with the feedback form responses.
This is not a polished, final product, so expect bugs and crashes. Reach out to MIDN Polidori,
your previous S3-LD, or a CS friend to help you fix any bugs or crashes that you experience.

The general steps to operate this bot are as follows:
1. Download the zip file from either github or the shared drive and unzip the files
2. Download the excel spreadsheet (as a .xlsx file) containing responses to this form, place it in the S3LDBot/spreadsheets folder, and rename it to 'Feedback.xlsx'
3. Run the feedback.py script. This will take the information from the spreadsheet and output them to .txt files that are drafts of the email in the emails directory.
4. Run the check.py script. This script will tell you everyone who has not filled out this form. The program will write funky emails to those who have not filled it out.
5. Double check some of the files in the emails directory to make sure its formatted correctly. The filename corresponds to who it will be sent to
6. Run the gmail.py script. This will send out all the emails to MIDN.

###Important Notes:
There are many inputs that will cause this program to work not as intended. Some known ones and fixes are below. Please add any that you find to this file in the shared drive and/or request a commit github

1. To run this bot you must have pip3 and python3 installed along with some dependencies.
	- To install python and pip, use google or ask chatgpt/CS friend to walk you through an install on your computer
	- To check if you have python 3 installed, use the command 'python3 --version'. If it outputs 'Python 3.9.x' or something similar, you're good.
	- Do the same thing with pip3 to ensure that is downloaded
	- To install all required dependencies run: 'pip3 install openpyxl, ezgmail, os, pandas' in your terminal
	- To run the scripts, type 'python3 feedback.py' into your terminal for feedback.py. Replace feedback.py with check.py and gmail.py to run those scripts as well. Note you must be in the S3LD directory to run these scripts. Ask a CS friend for help if you do not know how to do this
	- Note: you must be running the above 2 commands by tying them in an app called Terminal for mac users or CMD for windows users

2. DO NOT EDIT THE QUESTIONS TO THE GOOGLE FORM. Even adding or removing a question will mess with the spreadsheets the results generate to, resulting in this bot either working very incorrectly or not at all
	
3. If two MIDN have the same last name (ie. MIDN 1/C Smith and MIDN 3/C Smith), there will
be one email sent to one of the two MIDN containing the feedback for both of them. I would
recommend just deleting the smith.txt file in the emails directory and then manually grabbing
information from the spreadsheet and sending the email to the two of them (the method the S3LD
used to have to use for every MIDN)

4. If a name is spelled incorrectly, you may see two files in the emails directory (ie.
smith.txt and smiht.txt). If this is the case, the gmail program will have issues sending out
the data and MIDN smith will miss out on some of their feedback. I would recommend just
control-f'ing the spreadsheet (that you already downloaded) and fixing the typo, then rerunning
feedback.py

5. If you are viewing/downloading these files from github, I have removed the token.json and credentials.json files for security reasons. THE GMAIL.PY FILE WILL NOT WORK WITHOUT THESE FILES DOWNLOADED. Please download them from the battalion staff shared drive or the previous S3LD and place them in the .../S3LDBot/ directory. It should be the folder that contains all of the python scripts you will be using
