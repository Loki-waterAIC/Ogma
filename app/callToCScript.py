import os
import sys

import runWordMacro

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)


VISIBILITY = False


def template_path_func() -> str:
    abs_path: str = os.path.abspath(".")
    dir_path: str = os.path.join(abs_path, "app")
    dir_path: str = os.path.join(dir_path, "documentTemplateMacros")
    dir_path: str = os.path.join(dir_path, "ogma.dotm")
    if not os.path.exists(path=dir_path):
        raise OSError(f"Template Path Does not Exist!\n\t>>> {dir_path}")
    return dir_path


def run_macro_on_doc(doc_paths: list[str], macro_path: str, macro_name: str, visibility: bool = False) -> None:
    """
    run_macro_on_doc runss a dotm macro from a given path on a list of documents.

    Args:
        doc_paths (list[str]): word document paths
        macro_path (str): macro file path
        macro_name (str): macro name in the macro file
        visibility (bool): show word or not. Default False
    """

    runWordMacro.run_word_macro_on_files(
        doc_paths=doc_paths,
        macro_name=macro_name,
        template_path=macro_path,
        wordVisible=visibility,
    )

    return


def update_doc_properties_multi(doc_paths: list[str]) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro on the word files at the given path at the same time

    Args:
        doc_path (str): _description_

    Raises:
        cscriptError: _description_
    """

    # set the macro
    macro: str = r"ogmaMacroAllFiles"
    template_path: str = template_path_func()
    visible = VISIBILITY

    run_macro_on_doc(doc_paths=doc_paths, macro_path=template_path, macro_name=macro, visibility=visible)

    return


def update_doc_properties(doc_paths: list[str]) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro on the word file at the given path

    Args:
        doc_path (str): _description_

    Raises:
        cscriptError: _description_
    """

    # set the macro
    macro: str = r"ogmaMacro"
    template_path: str = template_path_func()
    visible = VISIBILITY

    for doc in doc_paths:
        run_macro_on_doc(doc_paths=[doc], macro_path=template_path, macro_name=macro, visibility=visible)

    return


if __name__ == "__main__":
    from data.hidden.files import FILES  # This can be removed

    # Example usage
    file: str | list[str] = FILES[0]  # making it so it works both single and multiple file tests
    if isinstance(file, str):
        file = [file]
    update_doc_properties(doc_paths=file)
