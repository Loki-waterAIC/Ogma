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
from docx import Document

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import callToCScript
from app.cscriptErrors import cscriptError

from data.hidden.files import FILES


def __helper_update_properties(doc_path: str, properties: dict) -> None:
    document: docx.document.Document = docx.Document(docx=doc_path)

    for k in properties:
        document.custom_properties[k] = properties[k]

    document.save(path_or_stream=doc_path)


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
    # update the values
    __helper_update_properties(doc_path=doc_path, properties=properties)

    # set the values
    try:
        callToCScript.update_doc_properties(doc_path=doc_path)
    except cscriptError as e:
        raise cscriptError(f"CScript Error occured:\n{e}")
    except Exception as e:
        raise Exception(f"Generic Error occured:\n{e}")
    return


def modify_word_properties(
    file_paths: list[str] | str, properties: dict[str, str] | None = None
) -> None:
    # Define the properties and their default values

    if isinstance(file_paths, str):
        file_paths = [file_paths]

    if properties == None:
        # properties = {
        #     "BOK ID": "302.EDC 20250317-01.20pm",
        #     "Document Name": "Maciavelli 20250317-01.20pm",
        #     "Company Name": "AIC 20250317-01.20pm",
        #     "Division": "Automation Engineering 20250317-01.20pm",
        #     "Author": "Aaron Shackelford 20250317-01.20pm",
        #     "Company Address": "9332 Tech Center Dr Sacramento Ca | Suite 200 20250317-01.20pm",
        #     "Project Name": "Rocks and Socks 20250317-01.20pm",
        #     "Project Number": "57.9092 20250317-01.20pm",
        #     "End Customer": "W M Lyles 20250317-01.20pm",
        #     "Site Name": "Sacramento City 20250317-01.20pm",
        #     "File Name": "Ventura 20250317-01.20pm",
        # }
        
        properties = {
            "BOK ID": "kitty",
            "Document Name": "kitty",
            "Company Name": "kitty",
            "Division": "kitty",
            "Author": "kitty",
            "Company Address": "kitty",
            "Project Name": "kitty",
            "Project Number": "kitty",
            "End Customer": "kitty",
            "Site Name": "kitty",
            "File Name": "kitty",
        }

    file_paths.sort()
    lprint: str = "".join([str(i) + "\n" for i in file_paths])
    

    print(f"Running files:\n{lprint}")
    
    if True:

        with ThreadPoolExecutor() as e:
        # with ThreadPoolExecutor(max_workers=1 if __debug__ else None) as e:
            # Set the custom properties
            e.map(lambda x: set_custom_properties(doc_path=x, properties=properties),file_paths)
    else:
        for path in file_paths:
            set_custom_properties(doc_path=path, properties=properties)

    print("Finished")
    return


if __name__ == "__main__":
    modify_word_properties(FILES)
