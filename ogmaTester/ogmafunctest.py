import os
import sys

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

from ogmaTester.ogmaTestValues import modify_word_properties
from app.ogmaScripts.documentPropertyUpdateTool import document_properity_update_tool

def run_scripts(in_files:list[str]|str) -> None:
    values: tuple[list[str], dict[str, str]] = modify_word_properties(file_paths=in_files)
    files: list[str] = values[0]
    props: dict[str, str] = values[1]
    document_properity_update_tool(doc_paths=files, properties=props)

# MARK: Start Reading Here
if __name__ == "__main__":    
    from data.hidden.files import FILES
    
    run_scripts(in_files=FILES)

    