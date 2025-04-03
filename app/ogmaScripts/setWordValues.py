#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
 # @ Author: Aaron Shackelford
 # @ Create Time: 2025-04-03 14:54:07
 # @ Modified by: Aaron Shackelford
 # @ Modified time: 2025-04-03 14:55:15
 # @ Description:

Sets document property values

"""

import os
import sys
import warnings
from typing import Any

import filelock
import pythoncom
import win32com
import win32com.client
from win32com.client.dynamic import CDispatch

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

from app.ogmaGlobal import LOCK_FILE_PATH


def run_set_values(
    doc_paths: list[str],
    properties: dict[str, str],
    wordVisible: bool = False,
) -> None:
    """
    Runs a specified macro in a Word document.

    Args:
        doc_paths (list[str]): List of Full path to the Word document paths (.docx paths recommended).
        template_path (str|None): path of the normal.dotm file to use. If None, it will assume the Template is in Normal.dotm or in the Docm file.
        macro_names (str): Name of the macro to run.
        activeDocumentMacro (bool): True If the macro needs to be run on each document individually, False if all the documents can be ran at once. IE ActiveDocument vs all Open Documents.
        wordVisible (bool): display word or not. Default is False

    Raises:
        Exception: If an error occurs during execution.
        OSError: If a path is invalid.
    """

    if os.name != "nt":
        warnings.warn(message="This function is only designed for Windows Machines", category=RuntimeWarning)

    path_violation_list: list[str] = list()  # may use later, for now just stubbing for later code
    validated_doc_paths: list[str] = [i for i in doc_paths if i]

    if validated_doc_paths:

        # only one instance of word can be used at once, so we will use locks to prevent multiple instances of word to be open.
        # wait and grab lock
        lock = filelock.FileLock(LOCK_FILE_PATH)

        with lock:
            # 1
            word: CDispatch | None = None
            # Initialize the COM library for threading
            pythoncom.CoInitialize()

            def sub_func_cleanup_word_0p9s8bgsp3() -> None:
                """
                sub_func_cleanup_word_0p9s8bgsp3 cleans up word
                """
                # 4/6
                nonlocal word

                # Quit the Word application if it was started
                if word:
                    try:
                        # https://learn.microsoft.com/en-us/office/vba/api/word.application.quit(method)
                        word.Quit(
                            SaveChanges=False
                        )  # if could not save and close file from before, assume an error has occured and close everything without saving
                    except:
                        try:
                            word.Quit()  # just quit
                        except:
                            pass  # com objects are stupid and I give up. Parsing the XML would have been easier at this point....
                    finally:
                        word = None  # Sending word to garbage collector.
                        # not using del here becuase it causes a crash in the system.
                        # let it close "gracefully"

            def sub_func_cleanup_doc_0p9s8bgsp3(inner_doc: Any) -> None:
                """
                sub_func_cleanup_doc_0p9s8bgsp3 cleans up the doc if it was opened
                """
                nonlocal word

                # close doc if was opened
                if inner_doc and word:
                    try:
                        # https://learn.microsoft.com/en-us/office/vba/api/word.documents
                        inner_doc.Save()
                        # https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
                        inner_doc.Close(SaveChanges=-1)
                    except:
                        try:
                            inner_doc.Close()
                        except:
                            # if can't save, assume it is closed
                            pass
                    finally:
                        inner_doc = None  # prevent duplication

            try:
                # 2
                # Create word Application object
                # https://learn.microsoft.com/en-us/office/vba/api/word.application
                word = win32com.client.Dispatch(dispatch="Word.Application")
                # https://learn.microsoft.com/en-us/office/vba/api/word.application.visible
                # word.Visible = wordVisible
                word.Visible = str(wordVisible)

                # open each document individually
                for path in validated_doc_paths:
                    # set an empty var
                    doc: Any = None
                    try:
                        # open document # https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
                        doc: Any = word.Documents.Open(path)

                        # run macro # https://learn.microsoft.com/en-us/office/vba/api/word.application.run
                        for key, val in properties.items():
                            # putting things in try blocks cause microsoft garrentees nothing
                            try:
                                # Update custom properties
                                # name of the property
                                # link to content: If the property is tied to a word object, like a paragraph, a table, section, image, text box, etc
                                # type 4, string | 1: msoPropertyTypeNumber (Numeric) 2: msoPropertyTypeBoolean (Boolean) 3: msoPropertyTypeDate (Date) 4: msoPropertyTypeString (String)
                                # value to set the property to
                                doc.CustomDocumentProperties.Add(Name=str(key), LinkToContent=False, Type=4, Value=str(val))
                            except:
                                # mute errors and run next property
                                pass
                    except:
                        # mute errors and run next file
                        pass
                    finally:
                        # close doc if was opened
                        sub_func_cleanup_doc_0p9s8bgsp3(inner_doc=doc)

            except Exception as e:
                # 3
                err_message: str = (
                    f'[runWordMacroWin.run_word_macro_on_files 1] GenericError Occured in one of the files in: "{doc_paths}":\n\t[COM] Generic Error:\n\t{e}'
                )
                print(err_message)
                # 4
                sub_func_cleanup_word_0p9s8bgsp3()
                raise Exception(err_message)
            finally:
                # 3/5
                # 4/6
                sub_func_cleanup_word_0p9s8bgsp3()
                # Uninitialize the COM library for this thread
                pythoncom.CoUninitialize()

    if path_violation_list:

        err_message: str = ""
        err_message += "[runWordMacroWin.run_word_macro_on_files 2] Invalid Files:"
        for invalid_path in path_violation_list:
            err_message += f"\n{str(invalid_path)}"
        print(err_message)
        raise OSError(err_message)

    return
