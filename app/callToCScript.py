import os
import subprocess
import sys

import RunWordMacro
from cscriptErrors import cscriptError

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)
    
from data.hidden.files import FILES # This can be removed

RUN_SUBPROCESS:bool = False

def update_doc_properties(doc_path:str) -> None:
    '''
    update_doc_properties runs the "UpdateDocumentProperties" macro in the word file at the given path

    Args:
        doc_path (str): _description_

    Raises:
        cscriptError: _description_
    '''
    
    # set the macro
    macro:str = r"UpdateAllFields"
    visible = False
    # visible = True # keep true, as of 20250318, 
    #   word has a bug that when ran in invisible mode (false).
    #   the OGMA macro currently does not work correctly if word is invisible.
    #   Word also will not close properly and keep the file resource open.
    #   this will prevent us from using the file later
    
    
    if RUN_SUBPROCESS:
        # run macro as a cscript
        vbs_script = r"app\RunWordMacro.vbs"  # Update with the actual path
        
        visible = str(visible)
        
        # check if the macro exist

        try:
            result: subprocess.CompletedProcess[str] = subprocess.run(
                args=[r"cscript", r"/nologo", r"/b", vbs_script, doc_path, macro, visible],
                capture_output=True,
                text=True,
                check=True,
            )
            print(f"Macro ran without error for {doc_path}\n\tstdout >>> ", result.stdout, "\n\tstdout >>> ", str(result.returncode))
        except subprocess.CalledProcessError as e:
            raise cscriptError("Error running macro:\n", e.stderr)
        
    else:
        # run as a py script
        RunWordMacro.run_word_macro(doc_path=doc_path, macro_name=macro, wordVisible=visible)
        # raise Exception("Bad Path, shouldn't be going this way. Make sure RUN_SUBPROCESS is True")
        
    return

if __name__ == "__main__":
    # Example usage    
    update_doc_properties(doc_path=FILES[0])