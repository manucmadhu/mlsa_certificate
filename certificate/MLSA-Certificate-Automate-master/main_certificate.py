
import os
from certificate import *
from docx import Document
import csv
from docx2pdf import convert


# create output folder if not exist
try:
    os.makedirs("Output/Doc")
    os.makedirs("Output/PDF")
except OSError:
    pass


def get_participants(f):
    data = [] # create empty list
    with open(f, mode="r", encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row) # append all results
    return data

    
def create_docx_files(filename, list_participate, ambassador):
    for participate in list_participate:
        doc = Document(filename)
        name = participate["Name Surname"]
        event = participate["Workshop"]
        date = participate["date"]  # Ensure this matches the column name in your CSV
        print("Date read from CSV:", date)  # Print to debug
        replace_participant_name(doc, name)
        replace_event_name(doc, event)
        replace_ambassador_name(doc, ambassador)
        replace_date(doc, date)  # Ensure this function is called correctly
        doc.save('Output/Doc/{}.docx'.format(name))
        print("Output/{}.pdf Creating".format(name))
        convert('Output/Doc/{}.docx'.format(name), 'Output/Pdf/{}.pdf'.format(name))

# get certificate temple path
certificate_file =""
participate_file = ""
# Enter your name here [Ambassador Name]
ambassador_name = ""

# get participants
list_participate = get_participants(participate_file)
# process data
create_docx_files(certificate_file, list_participate, ambassador_name)



