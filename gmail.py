import ezgmail
import openpyxl
import os

#used to take last name from name string (protect against faulty inputs)
def name(s):
  # Split the string into a list of words
  try:
    words = s.split()
  except:
    return s
  
  # Return the last word in the list in lowercase
  return words[-1].lower()

#sends an email
def send_email(recipient, subject, txt_file_path):
  # Read the contents of the .txt file
  with open(txt_file_path, 'r') as f:
    body = f.read()

  # Send the email
  #ezgmail.send(recipient, subject, body)

#returns a dictionary of filenames as keys and associated emails to send to as values  
def get_emails(file_path):
    # Load the workbook
    wb = openpyxl.load_workbook(file_path)
    
    # Get the first sheet
    sheet = wb.worksheets[0]
    
    # Initialize the dictionary
    data = {}
    
    # Iterate through the rows
    for row in sheet.iter_rows(min_row=1):  # Skip the first row (row 1) as it contains the headers
        # Get the values in the second and third columns
        key = name(row[1].value)
        value = row[2].value
        
        # Add the key and value to the dictionary
        data[key] = value
    
    return data
 
#returns a list of files within a directory for reading the emails/ files
def get_files(directory_path):
    # Get a list of all the files in the directory
    files = os.listdir(directory_path)
    
    # Filter out any files that are not .xlsx files
    txt_files = [f for f in files if f.endswith('.txt')]
    
    return txt_files

dict = get_emails("spreadsheets/Feedback.xlsx")
#get a list of all .txt files
os.chdir("emails/")
files = get_files("./")

#send each of the text files to their associated email
for file in files:
    try:
        print(f"Sending email to: {dict[file[0:len(file) - 4]]}...")
        send_email(dict[file[0:len(file) - 4]],'Peer Feedback Results', os.path.abspath(file))
    except:
        print(f'ERROR Cannot send email to: {file[0:-4]}')

print("\n\nThank for using S3LD Bot. If you saw this run with out any error messages, you are good to go. However, if you saw an error with a MIDN name, check to see if that MIDN has completed the form, see if you can figure out what the issue is, or simply manually send the email with information from the spreadsheet to that MIDN. If you saw another error, refer to the README and if you cannot fix it, just manually send out all the emails using the information from the emails directory")
