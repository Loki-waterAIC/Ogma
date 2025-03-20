import copy
import io
import os
import secrets
import string
import sys
import threading
from collections.abc import Iterator
from concurrent.futures import ThreadPoolExecutor

# project path

OGMA_PATH: str = os.path.abspath(path=os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

import removeMacros
import runWordMacros

# lock for macro binary copy
BIN_COPY_LOCK = threading.Lock()


def generate_random_string(length: int) -> str:
    """
    Generate a random string of alphanumeric characters.

    Args:
        length (int): The length of the random string to generate.

    Returns:
        str: A random string of alphanumeric characters of the specified length.
    """
    # Define the characters to choose from (alphanumeric)
    characters: str = string.ascii_letters + string.digits
    # Generate a random string of the specified length (using secrets to avoid the chance of two threads producing the same result)
    random_string: str = "".join([secrets.choice(characters) for i in range(length)])
    return random_string


def multiply_macros(macro_path: str, num_to_multiply_to: int) -> list[str]:
    """
    multiply_macros  mutilpies a parent macro

    Args:
        macro_path (str): parent macro to copy
        num_to_multiply_to (int): number of copies to have
        already_copied (list[str]): list of already copied macros
    """

    def sub_thread_make_copies(macro_bin: io.BytesIO, inner_path_obj: dict[str, str]) -> str:
        inner_parent = io.BytesIO()
        with BIN_COPY_LOCK:
            inner_parent: io.BytesIO = copy.deepcopy(macro_bin)
        inner_name: str = inner_path_obj["name"]
        inner_ext: str = inner_path_obj["ext"]
        inner_dir: str = inner_path_obj["dir"]
        new_name: str = f"{inner_name}{generate_random_string(length=5)}{inner_ext}"
        macro_copy_path: str = os.path.normpath(os.path.join(inner_dir, new_name))
        with open(file=macro_copy_path, mode="wb") as child:
            inner_parent.seek(0)
            child.write(inner_parent.read())
        return macro_copy_path

    # Split the file path into directory, base name, and extension
    directory, base_name = os.path.split(macro_path)
    name, ext = os.path.splitext(base_name)
    parent_path_obj: dict[str, str] = {"name": name, "ext": ext, "dir": directory}

    # load macro binary into memory
    macro_bin = io.BytesIO()
    with open(file=macro_path, mode="rb") as parent:
        parent.seek(0)
        macro_bin = io.BytesIO(initial_bytes=parent.read())

    with ThreadPoolExecutor(max_workers=1 if __debug__ else None) as e:
        paths: Iterator[str] = e.map(lambda: sub_thread_make_copies(macro_bin=macro_bin, inner_path_obj=parent_path_obj), [i for i in range(num_to_multiply_to + 1)])
    return list(paths)


def template_path_func() -> str:
    abs_path: str = os.path.abspath(".")
    dir_path: str = os.path.join(abs_path, "app")
    dir_path: str = os.path.join(dir_path, "documentTemplateMacros")
    dir_path: str = os.path.join(dir_path, "ogma.dotm")
    if not os.path.exists(path=dir_path):
        raise OSError(f"Template Path Does not Exist!\n\t>>> {dir_path}")
    return dir_path


def update_doc_properties(doc_paths: list[str]) -> None:
    """
    update_doc_properties runs the "UpdateDocumentProperties" macro in the word file at the given path

    Args:
        doc_path (str): _description_

    Raises:
        cscriptError: _description_
    """

    # set the macro
    macro: str = r"ogmaMacro"
    template_path: str = template_path_func()
    visible = True

    macro_paths: list[str] = multiply_macros(macro_path=template_path, num_to_multiply_to=len(doc_paths))

    # TODO:
    # check to make sure the number of macros is more or equal to than the number of docs
    if len(macro_paths) < len(doc_paths):
        raise ValueError("You dun fucked up AARON!\n" "Files are the wrong sizes:\n\t" f"Macros >>> {len(macro_paths)} | Docs >>> {len(doc_paths)}")

    to_process = list(zip(doc_paths, macro_paths))
    with ThreadPoolExecutor(max_workers=1 if __debug__ else None) as e:
        e.map(
            lambda x: runWordMacros.run_word_macro(doc_path=x[0], macro_name=macro, template_path=x[1], wordVisible=visible),
            to_process,
        )

    # delete cloned macros
    removeMacros.delete_async(paths=macro_paths)
    return


if __name__ == "__main__":
    from data.hidden.files import FILES  # This can be removed

    # Example usage
    files: list[str] | str = FILES[0]
    if isinstance(files, str):
        files = [files]  # if str, throw in a list of itself
    update_doc_properties(doc_paths=files)

    # Example usage
    # file1 = "path/to/first/file"
    # file2 = "path/to/second/file"
    # print(f"File hashes are {'the same' if valid_copies(parent_path=file1, children_copies=[file2]) else 'different'}")
