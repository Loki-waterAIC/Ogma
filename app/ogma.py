"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-12 13:19:04
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-12 13:19:19
# @ Description:

 Ogma is a program that edits a word documents' propery values
"""

from app.cscriptErrors import cscriptError
import callToCScript
import datetime
import json
import os
import sys
import argparse
from concurrent.futures import ThreadPoolExecutor

import docx
import docx.document
import filelock

APP_VERSION = "2.0.0"

LOCK_FILE_PATH: str = os.path.join(os.path.abspath("."), os.path.join("tmp", "ogma_lock.lock"))


# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)


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


def get_current_datetime_str() -> str:
    # for testing, can be deleted.
    # Get the current datetime
    now: datetime.datetime = datetime.datetime.now()

    # Format the datetime string as YYYYMMDD-HH.MM[AM|PM]
    formatted_datetime: str = now.strftime("%Y%m%d-%I.%M%p")

    return formatted_datetime


def modify_word_properties(file_paths: list[str] | str, properties: dict[str, str] | None = None) -> None:
    # Define the properties and their default values

    # If passed a single file, make it a list with a single index
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    time: str = get_current_datetime_str()

    # If no properties are passed, default ones
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

    update_custom_document_properties(doc_paths=file_paths, properties=properties)

    return


if __name__ == "__main__":
    # MARK: START READING HERE

    # TODO
    # [ ] add an arugment for json data
    # ogma.exe "./instructions.json"
    # [ ] make it so the json data is turned into dict

    # TODO
    # refactor RunMacro for the single macro file processer

    program_name = "ogma.exe"
    program_description = (
        "Created by Aaron Shackelford\n"
        "Ogma is named after the Gallic god of writing, Ogma, also spelled Oghma or Ogmios from the Greek.\n"
        "Ogma can run as a service or a CLI tool to modify local Word files.\n"
        "Ogma takes in a JSON object as instructions on how to process Word files.\n"
        "Ogma is known for being detected as ransomware due to how it modifies Word files; and may need to be whitelisted.\n"
        "\nSee below for launch arguments.\n"
    )
    program_epilog = ""

    parser = argparse.ArgumentParser(prog=program_name, description=program_description, epilog=program_epilog)
    # parser.add_argument('--json', '-j',  action="store", type=str, nargs=1,  help='Path to the JSON file to instruct ogma on what to do.')
    parser.add_argument('--json', '-j',  action="store", type=argparse.FileType(mode='r'), nargs='?',  help='Path to the JSON file to instruct ogma on what to do.')
    parser.add_argument('--verbose', '-v', action='store_true', help='Increase output verbosity')
    # Commenting out this section to add back later
    # parser.add_argument("--service", "-s", action="store_true", help="launches ogma in service mode. Will run ogma until closed and process files as single requests.")
    # # 5282 was just the first number that popped into my head -Aaron, there is no significance, but it definitely should be specified when running ogma as a service.
    # parser.add_argument("--port", "-p", type=int, nargs=1, default=5282, help="In combination with --service, this specifies the port that ogma will communicate on. Default is port 5282")
    parser.add_argument('--version', action='version', version=f'%(prog)s {APP_VERSION}')
    # parser.print_help()
    args: argparse.Namespace = parser.parse_args()

    if args.instructions:
        try:
            with open(args.json_file, 'r') as file:
                data = json.load(file)
                if args.verbose:
                    print("Successfully parsed JSON file:")
                print(data)  # This will print the dictionary
        except FileNotFoundError:
            err_message: str = f"File not found: {args.json_file}"
            if args.verbose:
                print(err_message)
            raise FileNotFoundError(err_message)
        except json.JSONDecodeError:
            err_message: str = f"Error decoding JSON from file: {args.json_file}"
            if args.verbose:
                print(err_message)
            raise ValueError(err_message)
