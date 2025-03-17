"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-12 13:19:04
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-12 13:19:19
# @ Description:

 Ogma is a program that edits a word documents' propery values
"""

import os
import sys
from concurrent.futures import ThreadPoolExecutor

import docx
import docx.document

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import callToCScript
from app.cscriptErrors import cscriptError

from data.hidden.files import FILES


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

    document: docx.document.Document = docx.Document(docx=doc_path)

    for k in properties:
        document.custom_properties[k] = properties[k]
    
    document.save(path_or_stream=doc_path)
    
    try:
        callToCScript.update_doc_properties(doc_path=doc_path)
    except cscriptError as e:
        raise cscriptError(f"CScript Error occured:\n{e}")
    except Exception as e:
        raise Exception(f"Generic Error occured:\n{e}")
    return

def file_to_run(file_paths:list[str]) -> None:
    # Define the properties and their default values
    
    properties: dict[str, str] = {
        "BOK ID": "302.EDC",
        "Document Name": "Maciavelli",
        "Company Name": "AIC",
        "Division": "Automation Engineering",
        "Author": "Aaron Shackelford",
        "Company Address": '9332 Tech Center Dr Sacramento Ca | Suite 200',
        "Project Name": "Rocks and Socks",
        "Project Number": "57.9092",
        "End Customer": "W M Lyles",
        "Site Name": "Sacramento City",
        "File Name": "Ventura",
    }
    
    # with ThreadPoolExecutor(max_workers=1 if __debug__ else None) as e:
    #     # Set the custom properties
    #     e.map(lambda x: set_custom_properties(doc_path=x, properties=properties),file_paths)
    for path in file_paths:
        set_custom_properties(path, properties)
    return

if __name__ == "__main__":
    file_to_run(FILES)