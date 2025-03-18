import subprocess
from cscriptErrors import cscriptError
import RunWordMacro

RUN_SUBPROCESS:bool = False
RUN_SUBPROCESS:bool = True

def update_doc_properties(doc_path:str) -> None:
    '''
    update_doc_properties runs the "UpdateDocumentProperties" macro in the word file at the given path

    Args:
        doc_path (str): _description_

    Raises:
        cscriptError: _description_
    '''
    
    # set the macro
    macro:str = r"ogma"
    visible = True
    
    
    if RUN_SUBPROCESS:
        # run macro as a cscript
        vbs_script = r"app\RunWordMacro.vbs"  # Update with the actual path
        
        visible = str(visible)

        try:
            result: subprocess.CompletedProcess[str] = subprocess.run(
                args=[r"cscript", r"/nologo", r"/b", vbs_script, doc_path, macro, visible],
                capture_output=True,
                text=True,
                check=True,
            )
            print(f"Macro ran without error for {doc_path}\n\tstdout >>> ", result.stdout)
        except subprocess.CalledProcessError as e:
            raise cscriptError("Error running macro:\n", e.stderr)
        
    else:
        # run as a py script
        RunWordMacro.run_word_macro(doc_path=doc_path, macro_name=macro, wordVisible=visible)
        
    return

if __name__ == "__main__":
    # Example usage    
    update_doc_properties(doc_path=r"data\hidden\1. Revision History.docx")