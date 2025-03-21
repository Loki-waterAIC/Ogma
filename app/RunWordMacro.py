#!/usr/bin/env python3.11.11
# -*- coding: utf-8 -*-
"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-17 09:13:37
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-17 09:29:56
# @ Description: runs a macro in a given word document at path.
"""

import os
import sys
from typing import Any

import pythoncom
import win32com
import win32com.client
from win32com.client.dynamic import CDispatch

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)


def run_word_macro_on_files(doc_paths: list[str], macro_name: str, template_path: str | None, wordVisible: bool = False) -> None:
    """
    Runs a specified macro in a Word document.

    Args:
        doc_path (str): Full path to the Word document (.docm recommended).
        template_path (str|None): path of the normal.dotm file to use. If None, it will assume the Template is in Normal.dotm or in the Docm file.
        macro_name (str): Name of the macro to run.
        wordVisible (bool): display word or not. Default is False

    Raises:
        Exception: If an error occurs during execution.
    """
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
            word.Quit()
            word = None

    try:
        # TODO:
        # MAKE SURE MACRO ISN"T LOCKED OUT FROM PREVIOUS FILE...
        # [x] or maybe one word instance and thread the word docs.....
        # [x] Open the files one at the time in the same word instance

        # [ ] open all files one at a time?
        
        # [ ] open all files at once then run all at once?
        #   [ ] make sure to use locks to prevent files from opening when running
        
        
        # 2
        # Create word Application object
        word = win32com.client.Dispatch(dispatch="Word.Application")
        word.Visible = wordVisible

        # add macro
        if template_path:
            word.AddIns.Add(FileName=template_path, Install=True)
            # word.AddIns(template_path).Installed = False

        # open the documents
        doc_list: list[Any] = list() # they are Applicaiton.Word.Document types but that is not defined in python
        for path in doc_paths:
            # open the document and add it to the docs list
            # https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
            doc: Any = word.Documents.Open(path)
            doc_list.append(doc)
            
        # run macro
        # [ ] does this work on all files?
        word.Application.Run(macro_name)
        
        # Save/close the document if it was opened
        for doc in doc_list:
            if doc:
                try:
                    doc.Save()
                    doc.Close(SaveChanges=True)
                except:
                    # if can't save, assume it is closed
                    pass
                doc = None  # prevent duplication

    except AttributeError as e:
        # 3
        err_message = f'AttributeError Occured in "{doc_paths}":\n\tCouldn\'t run Macro "{macro_name}"\n\tError: >>> {e}'
        print(err_message)
        # 4
        sub_func_cleanup_0p9s8bgsp3()
        raise AttributeError(err_message)
    except Exception as e:
        # 3
        err_message: str = f'GenericError Occured in "{doc_paths}":\n\tGeneric Error:\n\t{e}'
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

    return


if __name__ == "__main__":
    from data.hidden.files import FILES, MACRO_FILES  # This can be removed

    # Example usage
    file: str | list[str] = FILES[0]  # making it so it works both single and multiple file tests
    if isinstance(file, str):
        file = [file]

    run_word_macro_on_files(
        doc_paths=file,
        template_path=MACRO_FILES[0],
        macro_name="ogmaMacro",
        wordVisible=True,
    )
