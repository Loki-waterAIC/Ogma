import subprocess
from cscriptErrors import cscriptError


def update_doc_properties(doc_path) -> None:
    vbs_script = r"app\RunWordMacro.vbs"  # Update with the actual path

    try:
        result: subprocess.CompletedProcess[str] = subprocess.run(
            [r"cscript", r"/nologo", r"/b", vbs_script, doc_path, r"UpdateDocumentProperties"],
            capture_output=True,
            text=True,
            check=True,
        )
        print(f"Macro ran successfully for {doc_path}:\n", result.stdout)
    except subprocess.CalledProcessError as e:
        raise cscriptError("Error running macro:\n", e.stderr)
    return