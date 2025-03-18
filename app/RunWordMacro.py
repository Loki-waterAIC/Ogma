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
    doc: Any = None
    word: CDispatch | None = None
    # Initialize the COM library for threading
    pythoncom.CoInitialize()
    try:
        # Create word Application object
        word = win32com.client.Dispatch(dispatch="Word.Application")
        word.Visible = wordVisible

        # open the word document
        doc = word.Documents.Open(doc_path)

        # run the macro
        # doc.Run(macro_name)
        word.Application.Run(macro_name)

    except AttributeError as e:
        print(f'AttributeError Occured in "{doc_path}":\n\tCouldn\'t run Macro "{macro_name}"\n\tError: >>> {e}')
    except Exception as e:
        print(f'GenericError Occured in "{doc_path}":\n\tGeneric Error:\n\t{e}')
    finally:
        # Save/close the document if it was opened
        if doc:
            doc.Save()
            doc.Close()

        # Quit the Word application if it was started
        if word:
            word.Quit()

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
