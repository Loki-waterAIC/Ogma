import os
import sys

import RunWordMacro

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)


def update_doc_properties(doc_path: str) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro in the word file at the given path

    Args:
        doc_path (str): _description_

    Raises:
        cscriptError: _description_
    """

    # set the macro
    macro: str = r"ogmaMacro"
    template_path=r".\app\ogma.dotm"
    visible = True

    RunWordMacro.run_word_macro(
        doc_path=doc_path,
        macro_name=macro,
        template_path=template_path,
        wordVisible=visible,
    )

    return


if __name__ == "__main__":
    from data.hidden.files import FILES  # This can be removed

    # Example usage
    update_doc_properties(doc_path=FILES[0])
