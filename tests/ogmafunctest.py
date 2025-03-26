import os
import sys

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

from app.ogmaScripts.ogmaTestValues import modify_word_properties
from app.ogmaScripts.documentPropertyUpdateTool import update_custom_document_properties

# MARK: Start Reading Here
if __name__ == "__main__":    
    from data.hidden.files import FILES

    values: tuple[list[str], dict[str, str]] = modify_word_properties(file_paths=FILES)
    files: list[str] = values[0]
    props: dict[str, str] = values[1]
    update_custom_document_properties(doc_paths=files, properties=props)