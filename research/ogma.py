#!/usr/bin/env python3

"""
 # @ Author: Aaron Shackelford
 # @ Create Time: 2025-07-16 16:24:33
 # @ Description:

Ogma is a program that updates a lot of document
properties at once.

 # @ Modified by: Aaron Shackelford
 # @ Modified time: 2025-07-16 16:24:37
"""

from concurrent.futures import ThreadPoolExecutor
from logging import Logger, getLogger
from pathlib import Path

from app.common.utils.pathFunctions import path_sanitization
from app.ogma.service.docx.update_custom_properties import update_custom_properties
from app.ogma.service.docx.update_properties_text import update_docx_properties
from app import DEBUG

logger: Logger = getLogger("streamline." + __name__)

# This is so contrived... why did I do this to myself.
# I could have been so simple and clean, but nooooooooooo
# - Aaron.


def __mass_update_helper(data: tuple[bool, Path], props: dict[str, str]) -> int:
    if not data[0]:
        return 0
    path: Path = data[1]
    modified_bin: bytes = b""

    # Automatically update file name
    if "File Name" in props and props["File Name"] == "<automatic>":
        props["File Name"] = path.stem

    # update the custom properties
    try:
        docx_bin: bytes = b""
        with open(file=path, mode="rb") as f:
            docx_bin = f.read()
        modified_bin: bytes = update_custom_properties(docx_bin=docx_bin, properties=props)
    except Exception as e:
        logger.error(e)
        return 0

    # update the custom properties visuals
    docx_return: tuple[bytes, bool] = update_docx_properties(docx_bin=modified_bin)
    modified_bin = docx_return[0]
    was_bad_file: bool = docx_return[1]

    if was_bad_file:
        logger.warning(f"Bad Briareus File was found at {path}")

    try:
        with open(file=path, mode="wb") as f:
            f.write(modified_bin)
    except Exception as e:
        logger.error(e)
        return 0

    return -1 if was_bad_file else 1


def __restore_file(path: Path) -> None:
    """
    __restore_file restores ogma file to docx file.

    Args:
        path (Path): path to restore.
    """
    if path.with_suffix(".ogma").is_file():
        try:
            with open(path.with_suffix(".ogma"), "rb") as in_file:
                with open(path, "wb") as out_file:
                    out_file.write(in_file.read())
        except Exception as e:
            logger.error(e)
            # we are just going to continue as if that didn't happen....


def __duplicate_file(path: Path) -> None:
    """
    __duplicate_file Duplicate file given, duplicated file will be in the same dir with suffix ".ogma"

    Args:
        path (Path): path to duplicate
    """
    with open(path, "rb") as in_file:
        with open(path.with_suffix(".ogma"), "wb") as out_file:
            out_file.write(in_file.read())


def __test_and_duplicate(path: Path) -> bool:
    """
    __test_file Tests to make sure the files can be opened.

    Args:
        path (Path): path to test and duplicate

    Returns:
        bool: file could be opened and duplicated successfully
    """
    try:
        __duplicate_file(path=path)
    except Exception as e:
        # some of these errors could be put into info, but I don't know them all yet.
        logger.error(e)
        return False
    return True


def __is_path_excluded(path: Path, exclusion_paths: set[Path]) -> bool:
    if path in exclusion_paths:
        return True
    return False


def __is_file_excluded(path: Path, exclusions_files: set[str]) -> bool:
    if path.stem in exclusions_files:
        return True
    if path.name in exclusions_files:
        return True
    if path.name.startswith("~$"):
        # ~$ is a word temp file.
        return True
    return False


def __get_documents(path: str, exclude: list[str]) -> list[Path]:
    dir_path: Path = path_sanitization(path=path)

    exclusion_path_strings: set[str] = set()
    exclusion_file_names: set[str] = set()
    for i in exclude:
        if "\\" or "/" in i:
            # exclude path
            exclusion_path_strings.add(i)
        else:
            # exclude file name
            exclusion_file_names.add(i)

    exclusion_path_paths: set[Path] = set([(path_sanitization(i)) for i in exclusion_path_strings])

    docx_files: list[Path] = []
    for r, ds, fs in dir_path.walk():
        if __is_path_excluded(r, exclusion_paths=exclusion_path_paths):
            continue
        for f in fs:
            jtf = Path(f)  # just the file
            if jtf.suffix == ".docx":
                fp: Path = r / f  # file path
                if __is_file_excluded(path=jtf, exclusions_files=exclusion_file_names):
                    continue
                docx_files.append(fp)
    return docx_files


def __test_open(path: Path) -> bool:
    if path.name.startswith("~$"):
        # temp file found!
        return False
    try:
        t_bytes: bytes = b""
        with open(file=path, mode="rb") as f:
            t_bytes = f.read()
        with open(file=path, mode="wb") as f:
            f.write(t_bytes)
    except Exception:
        return False
    return True


def test_files(directory: str, exclusions: list[str]) -> list[tuple[bool, str]]:
    docx_files_paths: list[Path] = __get_documents(path=directory, exclude=exclusions)

    good_returns: list[bool] = []
    with ThreadPoolExecutor(max_workers=None) as e:
        good_returns = list(e.map(__test_open, docx_files_paths))

    good_paths: zip[tuple[bool, Path]] = zip(good_returns, docx_files_paths)
    return [(i[0], str(i[1])) for i in good_paths]


def mass_document_property_update(directory: str, props: dict[str, str], exclusions: list[str]) -> str:
    """
    mass_document_property_update updates all the docx files in a given folder path

    Args:
        dir (str): directory to update files in
        props (dict[str,str]): properties to update

    Returns:
        str | None: error message bytes or None if no errors occurred.
    """

    docx_files: list[Path] = __get_documents(path=directory, exclude=exclusions)

    good_paths: list[bool] = []
    with ThreadPoolExecutor(max_workers=None) as e:
        good_paths = list(e.map(__test_and_duplicate, docx_files))

    run_it: zip[tuple[bool, Path]] = zip(good_paths, docx_files)

    successes: list[int] = []
    with ThreadPoolExecutor(max_workers=1 if DEBUG else None) as e:
        successes = list(e.map(lambda x: __mass_update_helper(data=x, props=props), run_it))

    bad_conversions: zip[tuple[bool, int, Path]] = zip(good_paths, successes, docx_files)

    # generate error message
    m: str = ""
    restore_path: list[Path] = []
    for i in bad_conversions:
        if not i[0]:
            m += f"Path failed to open, Does someone else have the file open? {i[2].__str__()}" + "\n"
        if i[0] and (i[1] < 1):
            m += f"Error Occurred when modifying the following Docx file:"
            if i[1] == -1:
                # bad briareus file
                m += f"\n\tFile was a bad Briareus File from Briareus version 0.1, properties were set but not updated in the document."
                m += "\n\t Open the file, let Word resolve the issue, then resave it and overwrite the original file."
                m += "\n\t Note that this file will not convert until this is fixed."
                m += f"\n\t{i[2].__str__()}" + "\n"
            else:
                m += f": {i[2].__str__()}" + "\n"
                restore_path.append(i[2])

    # Restore files that failed conversion?
    with ThreadPoolExecutor(max_workers=None) as e:
        e.map(__restore_file, restore_path)

    return m
