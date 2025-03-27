import json
import os
import subprocess
import sys

# project path
OGMA_PATH: str = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import ogmaTester.ogmaTestValues as OTV
from data.hidden.files import FILES, OGMA_PYTHON_LOCATION, OGMA_PYTHON_SCRIPT_LOCATION


# MARK: RUNNER
def ogma_run(run_files: list[str] | str) -> None:
    # make values
    values: tuple[list[str], dict[str, str]] = OTV.modify_word_properties(file_paths=run_files)
    files: list[str] = values[0]
    props: dict[str, str] = values[1]

    # make json
    data: dict[str, dict[str, str] | list[str]] = {"files": files, "doc_properties": props}
    # print("json object created:\n" f"{json.dumps(data,indent=3)}")
    abs_path: str = os.path.abspath(".")
    json_path: str = os.path.join(abs_path, "ogmaTester")
    json_path: str = os.path.join(json_path, "data.json")
    with open(file=json_path, mode="w", encoding="utf-8") as f:
        json.dump(obj=data, fp=f, indent=3)

    # send json to ogma
    subprocess.run(args=[OGMA_PYTHON_LOCATION, "-OO", OGMA_PYTHON_SCRIPT_LOCATION, "-j", json_path])


if __name__ == "__main__":
    ogma_run(run_files=FILES[0])
