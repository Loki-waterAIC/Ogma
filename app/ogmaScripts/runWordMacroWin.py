#!/usr/bin/env python3.11.11
# -*- coding: utf-8 -*-
"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-17 09:13:37
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-17 09:29:56
# @ Description: 
# 
#   runs a macro in a given word document at path on windows using Windows Com objects.
#   Documentation for windows word application com objects can be found here:
#   https://learn.microsoft.com/en-us/office/vba/api/overview/word
#
#
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
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)
    
from app.ogmaGlobal import LOCK_FILE_PATH


def run_word_macro_on_files(doc_paths: list[str], macro_names: list[str], template_path: str | None, activeDocumentMacro : bool, wordVisible: bool = False) -> None:
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
    
    if os.name != 'nt':
        warnings.warn(message="This function is only designed for Windows Machines",category=RuntimeWarning)
    
    # only one instance of word can be used at once, so we will use locks to prevent multiple instances of word to be open.
    # wait and grab lock
    lock = filelock.FileLock(LOCK_FILE_PATH)

    with lock:        
        path_violation_list: list[str] = list() # may use later, for now just stubbing for later code
        validated_doc_paths: list[str] = [i for i in doc_paths if i]
        validated_macro_name:list[str] = [i for i in macro_names if i]
        
        # 1
        word: CDispatch | None = None
        # Initialize the COM library for threading
        pythoncom.CoInitialize()

        def sub_func_cleanup_0p9s8bgsp3() -> None:
            """
            sub_func_cleanup_0p9s8bgsp3 cleans up the doc and word file if it was opened
            """
            # 4/6
            nonlocal word

            # Quit the Word application if it was started
            if word:
                try:
                    # https://learn.microsoft.com/en-us/office/vba/api/word.application.quit(method)
                    word.Quit(SaveChanges=False) # if could not save and close file from before, assume an error has occured and close everything without saving
                except:
                    try:
                        word.Quit() # just quit
                    except:
                        pass # com objects are stupid and I give up. Parsing the XML would have been easier at this point.... 
                finally:
                    word = None # Sending word to garbage collector.
                    # not using del here becuase it causes a crash in the system.
                    # let it close "gracefully"

        try:
            # 2
            # Create word Application object
            # https://learn.microsoft.com/en-us/office/vba/api/word.application
            word = win32com.client.Dispatch(dispatch="Word.Application")
            # https://learn.microsoft.com/en-us/office/vba/api/word.application.visible
            # word.Visible = wordVisible
            word.Visible = str(wordVisible)

            # add macro
            if template_path:
                # https://learn.microsoft.com/en-us/office/vba/api/word.addins.add
                word.AddIns.Add(FileName=template_path, Install=True)

            # open the documents
            if activeDocumentMacro:
                # open each document individually
                for path in validated_doc_paths:
                    # set an empty var
                    doc : Any = None
                    try:
                        # open document # https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
                        doc: Any = word.Documents.Open(path)

                        # run macro # https://learn.microsoft.com/en-us/office/vba/api/word.application.run
                        for name in validated_macro_name:
                            if name:
                                word.Application.Run(name)
                    except:
                        # mute errors and run next file
                        pass
                    finally:
                        # close doc if was opened
                        if doc:
                            try:
                                # https://learn.microsoft.com/en-us/office/vba/api/word.documents
                                doc.Save()
                                doc.Close(SaveChanges=True)
                            except:
                                try:
                                    doc.Close()
                                except:
                                    pass
                                # if can't save, assume it is closed
                                pass
                            finally:
                                doc = None  # prevent duplication


            else:
                # run all files at once
                doc_list: list[Any] = list() # they are Applicaiton.Word.Document types but that is not defined in python
                for path in validated_doc_paths:
                    # open the document and add it to the docs list # https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
                    try:
                        doc: Any = word.Documents.Open(path)
                        doc_list.append(doc)
                    except:
                        # mute error and go to next document
                        pass

                # run macro # https://learn.microsoft.com/en-us/office/vba/api/word.application.run
                for name in validated_macro_name:
                    if name:
                        word.Application.Run(name)

                # Save/close the document if it was opened
                for doc in doc_list:
                    if doc:
                        try:
                            # https://learn.microsoft.com/en-us/office/vba/api/word.documents
                            doc.Save()
                            doc.Close(SaveChanges=True)
                        except:
                            # if can't save, assume it is closed
                            pass
                        finally:
                            doc = None  # prevent duplication

        except AttributeError as e:
            # 3
            err_message = f'AttributeError Occured in one of the files in: "{doc_paths}":\n\tCouldn\'t run Macro "{macro_names}"\n\tError: >>> {e}'
            print(err_message)
            # 4
            sub_func_cleanup_0p9s8bgsp3()
            raise AttributeError(err_message)
        except Exception as e:
            # 3
            err_message: str = f'GenericError Occured in one of the files in: "{doc_paths}":\n\tGeneric Error:\n\t{e}'
            print(err_message)
            # 4
            sub_func_cleanup_0p9s8bgsp3()
            raise Exception(err_message)
        finally:
            # 3/5
            # 4/6
            sub_func_cleanup_0p9s8bgsp3()
            # Uninitialize the COM library for this thread
            pythoncom.CoUninitialize()
            
        if path_violation_list:
            
            err_message:str = ""
            err_message += "Invalid Files:"
            for invalid_path in path_violation_list:
                err_message += f"\n{str(invalid_path)}"
            print(err_message)
            raise OSError(err_message)

    return


if __name__ == "__main__":
    # tester
    from data.hidden.files import FILES, MACRO_FILES  # This can be removed

    # Example usage
    file: str | list[str] = FILES[0]  # making it so it works both single and multiple file tests
    if isinstance(file, str):
        file = [file]

    run_word_macro_on_files(
        doc_paths=file,
        template_path=MACRO_FILES[0],
        macro_names=["ogmaMacro"],
        activeDocumentMacro=True,
        wordVisible=True,
    )
