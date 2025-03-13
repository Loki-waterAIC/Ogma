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

# Create a Word application object
word_app: CDispatch = win32com.client.Dispatch("Word.Application")

# Open the Word document
doc = word_app.Documents.Open(FILES[0])

# Run the VBA macro
word_app.Run("YourMacroName")

# Save and close the document
doc.Save()
doc.Close()

# Quit the Word application
word_app.Quit()

