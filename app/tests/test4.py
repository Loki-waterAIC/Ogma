"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-12 13:19:04
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-12 13:19:19
# @ Description:

 Ogma is a program that edits a word documents' propery values
"""

import io
import os
import subprocess
import sys

import docx
import docx.document
import docx.opc
import docx.opc.package
from docx import Document

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

from data.hidden.files import FILES


def property_exists(doc, prop_name: str) -> bool:
    """
    Check if a custom document property exists in the given Word document.

    Args:
        doc: The Word document object.
        prop_name (str): The name of the custom property to check.

    Returns:
        bool: True if the property exists, False otherwise.
    """
    try:
        doc.CustomDocumentProperties(prop_name)
        return True
    except:
        return False


def run_word_macro(doc_path, macro_name):
    vbs_script = r"app/RunWordMacro.vbs"  # Update with the actual path

    try:
        result: subprocess.CompletedProcess[str] = subprocess.run(
            ["cscript", "//nologo", vbs_script, doc_path, macro_name],
            capture_output=True,
            text=True,
            check=True,
        )
        print("Macro ran successfully:", result.stdout)
    except subprocess.CalledProcessError as e:
        print("Error running macro:", e.stderr)


def set_custom_properties(doc_path: str, properties: dict) -> None:
    """
    Set custom document properties in a Word document.

    Args:
        doc_path (str): The path to the Word document.
        properties (dict): A dictionary of property names and their default values.

    Example:
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
    """
    doc: docx.document.Document = Document(docx=doc_path)

    dist: int = max([len(i) for i in properties])

    # checking the current values
    for prop_id in properties:
        prop_val = doc.custom_properties[prop_id]
        print(f"{str(prop_id):>{dist}} || {str(prop_val)}")

    # setting test values
    for prop_id in properties:
        doc.custom_properties[prop_id] = properties[prop_id]

    # properties are assigned but not set....
    # run macro?
    run_word_macro(doc_path=doc_path, macro_name="DocPropMacro")

    # Verifying properies
    for prop_id in properties:
        prop_val = doc.custom_properties[prop_id]
        print(f"{str(prop_id):>{dist}} || {str(prop_val)}")

    # Saving Properties
    doc.save(doc_path)


# Define the properties and their default values
properties: dict[str, str] = {
    "BOK ID": "Python Updated Value",
    "Document Name": "Python Updated Value",
    "Company Name": "Python Updated Value",
    "Division": "Python Updated Value",
    "Author": "Python Updated Value",
    "Company Address": "Python Updated Value",
    "Project Name": "Python Updated Value",
    "Project Number": "Python Updated Value",
    "End Customer": "Python Updated Value",
    "Site Name": "Python Updated Value",
    "File Name": "Python Updated Value",
}

# Path to the Word document
doc_path: str = FILES[2]
print(doc_path)

# Set the custom properties
set_custom_properties(doc_path=doc_path, properties=properties)
