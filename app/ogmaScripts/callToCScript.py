#!/usr/bin/env python3.11.11
# -*- coding: utf-8 -*-
'''
 # @ Author: Aaron Shackelford
 # @ Create Time: 2025-03-25 11:39:16
 # @ Modified by: Aaron Shackelford
 # @ Modified time: 2025-03-26 13:10:08
 # @ Description: 
 
    callToCScript originally controlled word using cscripts in windows
    but has now has changed to a controlling script called runWordMacroWin
    
    !! runWordMacroWin can not be called more than once at a time. !!
    
    to insure this, we use locks 
 
 '''

import os
import sys

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import app.ogmaScripts.runWordMacroWin as runWordMacroWin

# True if Should word be visible; False if word should not be visible
WORDVISIBLITY = False

def template_path_func() -> str:
    abs_path: str = os.path.abspath(".")
    dir_path: str = os.path.join(abs_path, "app")
    dir_path: str = os.path.join(dir_path, "documentTemplateMacros")
    dir_path: str = os.path.join(dir_path, "ogma.dotm")
    if not os.path.exists(path=dir_path):
        raise OSError(f"Template Path Does not Exist!\n\t>>> {dir_path}")
    return dir_path


def update_doc_properties_multi(doc_paths: list[str]) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro on the word files at the given path at the same time

    Args:
        doc_paths (list[str]): files to process
    """

    # set the macro
    macro: str = r"ogmaMacroAllFiles"
    template_path: str = template_path_func()
    wordVisible: bool = WORDVISIBLITY
    
    
    runWordMacroWin.run_word_macro_on_files(
        doc_paths=doc_paths,
        macro_names=[macro],
        template_path=template_path,
        activeDocumentMacro=False,
        wordVisible=wordVisible,
    )
    return


def update_doc_properties(doc_paths: list[str]) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro on the word file at the given path

    Args:
        doc_path (str): files to process
    """

    # set the macro
    macro: str = r"ogmaMacro"
    template_path: str = template_path_func()
    wordVisible: bool = WORDVISIBLITY


    runWordMacroWin.run_word_macro_on_files(
        doc_paths=doc_paths,
        macro_names=[macro],
        template_path=template_path,
        activeDocumentMacro=True,
        wordVisible=wordVisible,
    )

    return


if __name__ == "__main__":
    from data.hidden.files import FILES  # This can be removed

    # Example usage
    file: str | list[str] = FILES[0]  # making it so it works both single and multiple file tests
    if isinstance(file, str):
        file = [file]
    update_doc_properties(doc_paths=file)
