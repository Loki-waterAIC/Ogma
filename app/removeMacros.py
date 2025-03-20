"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-20 11:35:00
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-20 11:37:39
# @ Description: This function is meant to be run as a separate process to remove files outside the main program.
"""

import multiprocessing
import os
import time
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from collections.abc import Iterator


def _delete_file_continuously(file_path: str | Path) -> None:
    """
    Continuously attempts to delete the specified file until successful.

    Args:
        file_path (str | Path): The path to the file to be deleted.
    """

    while os.path.exists(path=file_path) and os.path.isfile(path=file_path):
        try:
            os.remove(path=file_path)
            print(f"Successfully deleted {file_path}")
            break
        except FileNotFoundError:
            # file is already gone
            break
        except OSError:
            # File is locked or not accessible, wait and try again
            time.sleep(1)


def _start_deletion_process(file_path: str | Path) -> multiprocessing.Process:
    """
    Starts a separate process to delete the specified file.

    Args:
        file_path (str | Path): The path to the file to be deleted.

    Returns:
        multiprocessing.Process: The process that will attempt to delete the file.
    """
    deletion_process = multiprocessing.Process(target=_delete_file_continuously, args=(file_path,))
    deletion_process.start()
    return deletion_process


def delete_async(paths: str | Path | list[str] | list[Path] | list[str | Path]) -> multiprocessing.Process | list[multiprocessing.Process]:
    """
    Initiates the asynchronous deletion of the specified file.

    Args:
        file (str | Path): The path to the file to be deleted.

    Returns:
        multiprocessing.Process: The process that will attempt to delete the file.
    """
    if isinstance(paths, list):
        with ThreadPoolExecutor() as e:
            pids: Iterator[multiprocessing.Process] = e.map(_start_deletion_process, paths)
        return list(pids)
    else:
        return _start_deletion_process(file_path=paths)
