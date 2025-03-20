import os
import sys

import RunWordMacro

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)


def template_path_func() -> str:
    abs_path: str = os.path.abspath(".")
    dir_path: str = os.path.join(abs_path, "app")
    dir_path: str = os.path.join(dir_path, "ogma.dotm")
    return dir_path


def update_doc_properties(doc_paths: list[str]) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro in the word file at the given path

    Args:
        doc_path (str): _description_

    Raises:
        cscriptError: _description_
    """

    # set the macro
    macro: str = r"ogmaMacro"
    template_path:str = template_path_func()
    visible = True

    RunWordMacro.run_word_macro_on_files(
        doc_paths=doc_paths,
        macro_name=macro,
        template_path=template_path,
        wordVisible=visible,
    )

    return


if __name__ == "__main__":
    from data.hidden.files import FILES  # This can be removed

    # Example usage
    file: str | list[str] = FILES[0] # making it so it works both single and multiple file tests
    if isinstance(file, str):
        file = [file]
    update_doc_properties(doc_paths=file)
