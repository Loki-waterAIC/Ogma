import os
import sys
from concurrent.futures import ThreadPoolExecutor

import docx
import docx.document
import filelock
from ogmaGlobal import LOCK_FILE_PATH

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)
    
import callToCScript
from cscriptErrors import cscriptError


def __helper_update_properties(doc_path: str, properties: dict) -> None:
    '''
    __helper_update_properties updates the default values of a property in a document's properties. 

    Args:
        doc_path (str): document path
        properties (dict): dictionary of properties to update. `{"property name" : "property value"}`

    Raises:
        Exception: docx documents have locks, if a document is locked, it can not be updated.
    '''
    try:
        # try to open the document
        document: docx.document.Document = docx.Document(docx=doc_path)
    except Exception as e:
        # document was not found or locked.
        err_message: str = f"Exception: can't open ({doc_path})\n\tError >>> {e}"
        print(err_message)
        raise Exception(err_message)

    for k in properties:
        document.custom_properties[k] = properties[k]

    document.save(path_or_stream=doc_path)


# MARK: START READING HERE
def update_custom_document_properties(doc_paths: list[str], properties: dict) -> None:
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
    # Sanatizing input file paths
    # Throw Error after processing
    path_violation_list: list[str] = list()
    validated_doc_paths: list[str] = list()
    for path in doc_paths:
        try:
            if os.path.exists(path):
                validated_doc_paths.append(path)
            else:
                path_violation_list.append(path)
        except:
            # add to violation list and go to next path
            path_violation_list.append(path)

    # update the values
    try:
        # for each path, update properties in a unique thread
        with ThreadPoolExecutor(max_workers=1 if __debug__ else None) as e:
            e.map(lambda x: __helper_update_properties(doc_path=x, properties=properties), validated_doc_paths)
    except Exception as e:
        # error can occure if a a document is open.
        err_message: str = f"Exception: {e}"
        print(err_message)
        raise Exception(err_message)

    # only one instance of word can be used at once, so we will use locks to prevent multiple instances of word to be open.
    # wait and grab lock
    lock = filelock.FileLock(LOCK_FILE_PATH)

    with lock:
        # set the values
        try:
            callToCScript.update_doc_properties_multi(doc_paths=validated_doc_paths)
        except AttributeError as e:
            print(e)
            raise e
        except cscriptError as e:
            raise cscriptError(f"CScript Error occured:\n{e}")
        except Exception as e:
            raise Exception(f"Generic Error occured:\n{e}")

    if path_violation_list:

        err_message: str = ""
        err_message += "Invalid Files:"
        for invalid_path in path_violation_list:
            err_message += f"\n{str(invalid_path)}"
        print(err_message)
        raise OSError(err_message)
    return
