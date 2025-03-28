#!/usr/bin/env python3.11.11
# -*- coding: utf-8 -*-
'''
 # @ Author: Aaron Shackelford
 # @ Create Time: 2025-03-25 11:39:16
 # @ Modified by: Aaron Shackelford
 # @ Modified time: 2025-03-26 13:10:08
 # @ Description:

 This file contains the document property update tool for Ogma

 It takes in request and processes them to send to Word via the callToCScript subroutine

 '''

import os
import sys
from concurrent.futures import ThreadPoolExecutor

import docx
import docx.document

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import app.ogmaScripts.callToCScript as callToCScript
from app.ogmaScripts.cscriptErrors import cscriptError


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
        err_message: str = f"[documentPropertyUpdateTool.__helper_update_properties] Exception: can't open ({doc_path})\n\tError >>> {e}"
        print(err_message)
        raise Exception(err_message)

    for k in properties:
        document.custom_properties[k] = properties[k]

    document.save(path_or_stream=doc_path)


# MARK: START READING HERE
def document_properity_update_tool(doc_paths: list[str], properties: dict) -> None:
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
        _err_message: str = f"[documentPropertyUpdateTool.document_properity_update_tool 0] Exception: {e}"
        print(_err_message)
        raise Exception(_err_message)


    # set the values
    # try range because word is stupid and trying again can help
    errors:list[Exception] = []
    no_success = True
    for _try in range(3):
        try:
            # callToCScript.update_doc_properties_multi(doc_paths=validated_doc_paths)
            callToCScript.update_doc_properties(doc_paths=validated_doc_paths)
            no_success = False
            break
        except Exception as e:
            errors.append(e)

    if no_success:
        loop_err_message:str = ""
        for e in errors:
            if isinstance(e,AttributeError):
                loop_err_message += f"\n[documentPropertyUpdateTool.document_properity_update_tool 1] AttributeError occured:\n{e}"
            elif isinstance(e,cscriptError):
                loop_err_message += f"\n[documentPropertyUpdateTool.document_properity_update_tool 1] cscriptError occured:\n{e}"
            elif isinstance(e,Exception):
                loop_err_message += f"\n[documentPropertyUpdateTool.document_properity_update_tool 2] Exception occured:\n{e}"

        if loop_err_message:
            print(loop_err_message)
            raise Exception(loop_err_message)

    if path_violation_list:
        _err_message: str = ""
        _err_message += "[documentPropertyUpdateTool.document_properity_update_tool 3] Invalid Files:"
        for invalid_path in path_violation_list:
            _err_message += f"\n\t{str(invalid_path)}"
        print(_err_message)
        raise OSError(_err_message)
    return
