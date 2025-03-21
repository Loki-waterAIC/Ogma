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
import datetime

import docx
import docx.document
from docx import Document

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import callToCScript
from app.cscriptErrors import cscriptError


def __helper_update_properties(doc_path: str, properties: dict) -> None:
    try:
        document: docx.document.Document = docx.Document(docx=doc_path)
    except Exception as e:
        err_message: str = f"Exception: can't open ({doc_path})\n\tError >>> {e}"
        print(err_message)
        raise Exception(err_message)

    for k in properties:
        document.custom_properties[k] = properties[k]

    document.save(path_or_stream=doc_path)


def set_custom_properties(doc_paths: list[str], properties: dict) -> None:
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
    try:
        with ThreadPoolExecutor(max_workers=1 if __debug__ else None) as e:
            e.map(lambda x: __helper_update_properties(doc_path=x, properties=properties), doc_paths)
    except Exception as e:
        err_message: str = f"Exception: {e}"
        print(err_message)
        raise Exception(err_message)

    # set the values
    try:
        callToCScript.update_doc_properties(doc_paths=doc_paths)
    except AttributeError as e:
        print(e)
        raise e
    except cscriptError as e:
        raise cscriptError(f"CScript Error occured:\n{e}")
    except Exception as e:
        raise Exception(f"Generic Error occured:\n{e}")
    return


def get_word_properties() -> dict[str, str]:
    return {
        "BOK ID": "",
        "Document Name": "",
        "Company Name": "",
        "Division": "",
        "Author": "",
        "Company Address": "",
        "Project Name": "",
        "Project Number": "",
        "End Customer": "",
        "Site Name": "",
        "File Name": "",
    }


def get_current_datetime_str() -> str:
    # for testing, can be deleted.
    # Get the current datetime
    now: datetime.datetime = datetime.datetime.now()

    # Format the datetime string as YYYYMMDD-HH.MM[AM|PM]
    formatted_datetime: str = now.strftime("%Y%m%d-%I.%M%p")

    return formatted_datetime

def modify_word_properties(file_paths: list[str] | str, properties: dict[str, str] | None = None) -> None:
    # Define the properties and their default values

    if isinstance(file_paths, str):
        file_paths = [file_paths]

    time: str = get_current_datetime_str()

    if properties == None:
        properties = {
            "BOK ID": f"BOK ID {time}",
            "Document Name": f"DOC NAME {time}",
            "Company Name": f"CO NAME {time}",
            "Division": f"DIV {time}",
            "Author": f"AUTH {time}",
            "Company Address": f"ADDR {time}",
            "Project Name": f"PRJ NAME {time}",
            "Project Number": f"PRJ ID {time}",
            "End Customer": f"END CUST {time}",
            "Site Name": f"SITE NAME {time}",
            "File Name": f"FILE NAME {time}",
        }

    set_custom_properties(doc_paths=file_paths, properties=properties)

    return


if __name__ == "__main__":
    from data.hidden.files import FILES

    modify_word_properties(file_paths=FILES)
