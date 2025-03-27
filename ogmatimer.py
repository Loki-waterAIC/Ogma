import os
import sys
import timeit
from datetime import datetime as dt
import json
import subprocess
import tempfile

# project path
OGMA_PATH: str = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import ogmaTester.ogmaTestValues as OTV
from data.hidden.files import FILES, OGMA_PYTHON_LOCATION, OGMA_PYTHON_SCRIPT_LOCATION


# MARK: RUNNER
def ogma_run(run_files: list[str] | str, times:int) -> float:
    # make values
    values: tuple[list[str], dict[str, str]] = OTV.modify_word_properties(file_paths=run_files)
    files: list[str] = values[0]
    props: dict[str, str] = values[1]

    # make json
    data: dict[str, dict[str, str] | list[str]] = {"files": files, "doc_properties": props}
    json_path: str = tempfile.gettempprefix() + ".json"
    with open(file=json_path, mode="w", encoding="utf-8") as f:
        json.dump(obj=data, fp=f, indent=3)

    # send json to ogma and time it
    time: float = timeit.timeit(
        stmt=lambda: subprocess.run(args=[OGMA_PYTHON_LOCATION, "-OO", OGMA_PYTHON_SCRIPT_LOCATION, "-j", json_path]),
        number=times,
    )

    # remove/Delete the file
    try:
        os.remove(json_path)
        print(f"{json_path} has been deleted successfully.")
    except FileNotFoundError:
        print(f"{json_path} does not exist.")
    except PermissionError:
        print(f"Permission denied: unable to delete {json_path}.")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        return time


# MARK: TEST RUNNER
def test_runner() -> None:

    # making run output title
    now: dt = dt.now()  # separate line to insure nothing breaks
    # formatted_datetime: str = now.strftime(r"%Y%m%d %I:%M %p").lower()
    formatted_datetime: str = now.strftime(r"%Y%m%d %I:%M %p").lower()
    break_string = "\n\n" + ("=" * 10) + "run output for " + formatted_datetime + ("=" * 10) + "\n\n"

    # data obj
    data: list[list[str | int | float]] = list()

    # run tests
    for i in range(len(FILES)):
        print(f"--runing {i} files")
        if FILES[:i]:
            times = 10
            time: float = ogma_run(run_files=FILES[:i], times=times)
            info: list[str | int | float] = ["number of files:", i, "avg run time:", time, "over", times, " times"]
            data.append(info)
        print(f"finished {i} files")

    # format data
    o_str: str = ""
    for i in data:
        j = map(lambda x: str(x), i)
        o_str += " ".join(j) + "\n"

    with open(r"ogmaTester\data_output.txt", "a", encoding="utf-8") as f:
        # Format the datetime string as "yyyymmdd h:mm am/pm"
        f.write(break_string)
        f.write(o_str)

    print(break_string)
    print(o_str)

if __name__ == "__main__":
    test_runner()