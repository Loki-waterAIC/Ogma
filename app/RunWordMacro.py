"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-17 09:13:37
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-17 09:29:56
# @ Description: runs a macro in a given word document at path.
"""

from typing import Any
import win32com
import win32com.client
from win32com.client.dynamic import CDispatch

def run_word_macro(doc_path: str, macro_name:str) -> None:
    '''
    Runs a specified macro in a Word document.

    Args:
        doc_path (str): Full path to the Word document (.docm recommended).
        macro_name (str): Name of the macro to run.

    Raises:
        Exception: If an error occurs during execution.
    '''
    try:
        # Create word Application object
        word: CDispatch = win32com.client.Dispatch(dispatch="Word.Application")
        word.Visible = True if __debug__ else False # if in __debug__ mode show word for macro debug too

        # open the word document
        doc: Any = word.Documents.Open(doc_path)

        # run the macro
        doc.Run(macro_name)

        # save/close the document
        doc.Save()
        doc.Close()

    except Exception as e:
        print(f"Generic Error:\n{e}")
        
    return


if __name__ == "__main__":
    # Example usage    
    run_word_macro(doc_path=r"data\hidden\1. Revision History.docx", macro_name="UpdateAllFields")
