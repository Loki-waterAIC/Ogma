import os
import sys
import timeit
from datetime import datetime as dt

# project path
OGMA_PATH: str = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import ogmaTester.ogmaCLITester as OCT
from data.hidden.files import FILES

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
            time: float = timeit.timeit(stmt=lambda: OCT.ogma_run(run_files=FILES[:i]), number=times)
            info: list[str | int | float] = ["number of files:", i, "avg run time:", time, "over", times, " times"]
            data.append(info)
        print(f"finished {i} files")

    # format data
    o_str:str = ""
    for i in data:
        j = map(lambda x: str(x), i)
        o_str += " ".join(j) + "\n"

    with open(r"ogmaTester\data_output.txt", "a", encoding="utf-8") as f:
        # Format the datetime string as "yyyymmdd h:mm am/pm"
        f.write(break_string)
        f.write(o_str)

    print(break_string)
    print(o_str)