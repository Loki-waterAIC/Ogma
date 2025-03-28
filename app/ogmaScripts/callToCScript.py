#!/usr/bin/env python3.11.11
# -*- coding: utf-8 -*-
"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-25 11:39:16
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-26 13:10:08
# @ Description:

   callToCScript originally controlled word using cscripts in windows
   but has now has changed to a controlling script called runWordMacroWin

   !! runWordMacroWin can not be called more than once at a time. !!

   to insure this, we use locks

"""

import os
import sys
import subprocess
import filelock

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import app.ogmaScripts.runWordMacroWin as runWordMacroWin
from app.ogmaScripts.cscriptErrors import cscriptError
from app.ogmaGlobal import LOCK_FILE_PATH

# True if Should word be visible; False if word should not be visible
WORDVISIBLITY = False


def __template_path_func() -> str:
    abs_path: str = os.path.abspath(".")
    dir_path: str = os.path.join(abs_path, "app")
    dir_path: str = os.path.join(dir_path, "documentTemplateMacros")
    dir_path: str = os.path.join(dir_path, "ogma.dotm")
    if not os.path.exists(path=dir_path):
        raise OSError(f"[callToCScript.template_path_func] Template Path Does not Exist!\n\t>>> {dir_path}")
    return dir_path


def __docx_to_pdf_vbs_path() -> str:
    abs_path: str = os.path.abspath(".")
    dir_path: str = os.path.join(abs_path, "app")
    dir_path: str = os.path.join(dir_path, "vbScripts")
    dir_path: str = os.path.join(dir_path, "docxToPdf.vbs")
    if not os.path.exists(path=dir_path):
        raise OSError(f"[callToCScript.template_path_func] Template Path Does not Exist!\n\t>>> {dir_path}")
    return dir_path


def docx_to_pdf(doc_paths: list[str], retry: int = 0) -> None:
    # recurse end
    if retry > 3:
        return

    retry_list: list[str] = []
    for input_docx in doc_paths:
        output_pdf: str = input_docx.removesuffix(r".docx") + r".pdf"
        output_pdf: str = os.path.normpath(os.path.abspath(output_pdf))
        vbs_script: str = __docx_to_pdf_vbs_path()

        lock = filelock.FileLock(LOCK_FILE_PATH)
        with lock:
            try:
                # Run the VBS script with input/output arguments
                subprocess.run(["cscript", vbs_script, input_docx, output_pdf], shell=True)
            except:
                retry_list.append(input_docx)

    # too lazy to not recurse
    if retry_list:
        docx_to_pdf(retry_list, retry=retry + 1)


def update_doc_properties_multi(doc_paths: list[str], export_pdf: bool = False) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro on the word files at the given path at the same time

    Args:
        doc_paths (list[str]): files to process
    """

    # set the macro
    macro: str = r"ogmaMacroAllFiles"
    template_path: str = __template_path_func()
    wordVisible: bool = WORDVISIBLITY

    runWordMacroWin.run_word_macro_on_files(
        doc_paths=doc_paths,
        macro_names=[macro],
        template_path=template_path,
        activeDocumentMacro=False,
        wordVisible=wordVisible,
    )
    if export_pdf:
        docx_to_pdf(doc_paths=doc_paths)
    return


def update_doc_properties(doc_paths: list[str], export_pdf: bool = False) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro on the word file at the given path

    Args:
        doc_path (str): files to process
    """

    # set the macro
    macro: str = r"ogmaMacro"
    template_path: str = __template_path_func()
    wordVisible: bool = WORDVISIBLITY

    runWordMacroWin.run_word_macro_on_files(
        doc_paths=doc_paths,
        macro_names=[macro],
        template_path=template_path,
        activeDocumentMacro=True,
        wordVisible=wordVisible,
    )
    if export_pdf:
        docx_to_pdf(doc_paths=doc_paths)
    return


if __name__ == "__main__":
    from data.hidden.files import FILES  # This can be removed

    # Example usage
    file: str | list[str] = FILES[0]  # making it so it works both single and multiple file tests
    if isinstance(file, str):
        file = [file]
    # update_doc_properties(doc_paths=file)
    docx_to_pdf(file)