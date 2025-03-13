'''
 # @ Author: Aaron Shackelford
 # @ Create Time: 2025-03-12 13:19:04
 # @ Modified by: Aaron Shackelford
 # @ Modified time: 2025-03-12 13:19:19
 # @ Description:
 
  Ogma is a program that edits a word documents' propery values
 '''

import os
import sys

import win32com.client
from win32com.client.dynamic import CDispatch

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",'..'))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)
    
from data.hidden.files import FILES

# Define the properties and their default values
properties = {
    "BOK ID": "WMLSI.XX.XX.XXX.X",
    "Document Name": "Document Name",
    "Company Name": "W. M. Lyles Co.",
    "Division": "System Integration Division",
    "Author": "Lastname, Firstname",
    "Company Address": "9332 Tech Center Drive, Suite 200 | Sacramento, CA 95826",
    "Project Name": "Project Name",
    "Project Number": "WMLSI.XX.XX.XXX.X",
    "End Customer": "End Customer",
    "Site Name": "Site Name",
    "File Name": "DocumentFileName"
}

# Create a Word application object
word_app: CDispatch = win32com.client.Dispatch(dispatch="Word.Application")

# Open the Word document
doc = word_app.Documents.Open(FILES[0])

# Function to check if a property exists
def property_exists(doc, prop_name) -> bool:
    try:
        doc.CustomDocumentProperties(prop_name)
        return True
    except:
        return False

# Set the custom properties
for prop_name, default_val in properties.items():
    if not property_exists(doc, prop_name):
        doc.CustomDocumentProperties.Add(
            Name=prop_name,
            LinkToContent=False,
            Type=4,  # msoPropertyTypeString
            Value=default_val
        )

# Save and close the document
doc.Save()
doc.Close()

# Quit the Word application
word_app.Quit()

