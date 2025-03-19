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

from data.hidden.files import FILES  # This can be removed


def run_word_macro(doc_path: str, macro_name: str, wordVisible: bool) -> None:
    """
    Runs a specified macro in a Word document.

    Args:
        doc_path (str): Full path to the Word document (.docm recommended).
        macro_name (str): Name of the macro to run.

    Raises:
        Exception: If an error occurs during execution.
    """
    # 1
    doc: Any = None
    word: CDispatch | None = None
    # Initialize the COM library for threading
    pythoncom.CoInitialize()

    def sub_func_cleanup_0p9s8bgsp3() -> None:
        """
        sub_func_cleanup_0p9s8bgsp3 cleans up the doc and word file if it was opened
        """
        # 4/6
        nonlocal doc, word

        # Save/close the document if it was opened
        if doc:
            doc.Save()
            doc.Close()
            doc = None  # prevent duplication

        # Quit the Word application if it was started
        if word:
            word.Quit()
            word = None  # prevent duplication

    try:
        # 2
        # Create word Application object
        word = win32com.client.Dispatch(dispatch="Word.Application")
        word.Visible = wordVisible

        # open the word document
        doc = word.Documents.Open(doc_path)

        # run the macro
        # doc.Run(macro_name)
        word.Application.Run(macro_name)

    except AttributeError as e:
        # 3
        err_message = f'AttributeError Occured in "{doc_path}":\n\tCouldn\'t run Macro "{macro_name}"\n\tError: >>> {e}'
        print(err_message)
        # 4
        sub_func_cleanup_0p9s8bgsp3()
        raise AttributeError(err_message)
    except Exception as e:
        # 3
        err_message: str = (
            f'GenericError Occured in "{doc_path}":\n\tGeneric Error:\n\t{e}'
        )
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
    # Example usage
    run_word_macro(
        doc_path=FILES[0],
        macro_name="ogmaMacro",
        wordVisible=True,
    )
