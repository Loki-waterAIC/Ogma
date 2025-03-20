import hashlib
import os
import sys
from concurrent.futures import Future, ThreadPoolExecutor
from collections.abc import Iterator

import RunWordMacro

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

HASH_ALG = "sha256"


def template_path_func() -> str:
    abs_path: str = os.path.abspath(".")
    dir_path: str = os.path.join(abs_path, "app")
    dir_path: str = os.path.join(dir_path, "ogma.dotm")
    return dir_path


def get_file_hash(file_path: str, hash_algorithm: str = HASH_ALG) -> str:
    """
    Computes the hash of a file using the specified hash algorithm.

    Args:
        file_path (str): The path to the file to be hashed.
        hash_algorithm (str, optional): The hash algorithm to use (default is 'sha256').

    Returns:
        str: The hexadecimal hash of the file.
    """
    with open(file_path, "rb") as f:
        hash_func: hashlib._Hash = hashlib.new(name=hash_algorithm, data=f.read(), usedforsecurity=False)
        return hash_func.hexdigest()


def compare_file_copy_hashes(parent_path: str, children_copies: list[str], hash_algorithm: str = HASH_ALG) -> list[str]:
    """
    Compares the hashes of a parent to its children copies.

    Args:
        parent_path (str): The path to the first file.
        children_copies (list[str]): The path to the second file.
        hash_algorithm (str, optional): The hash algorithm to use (default is 'sha256').

    Returns:
        list[str]: list of the children that matched the parent plus the parent
    """
    with ThreadPoolExecutor() as e:
        parent_future: Future[str] = e.submit(lambda: get_file_hash(file_path=parent_path, hash_algorithm=hash_algorithm))
        children_futures: Iterator[str] = e.map(
            lambda x: get_file_hash(file_path=x, hash_algorithm=hash_algorithm),
            children_copies,
        )
    parent_hash: str = parent_future.result()
    children_hashes: list[str] = list(children_futures)

    good_paths: list[str] = [parent_path]

    children_data = list(zip(children_hashes, children_copies))

    good_paths.extend([x[1] for x in children_data if x[0] == parent_hash])
    return good_paths


def get_macro_paths(macro_path: str) -> list[str]:
    """
    get_macro_paths return the paths for the macro copies.

    Args:
        macro_path (str): parent macro path

    Returns:
        list[str]: macro copy paths + parent path
    """
    raise Warning("Function not implemented yet")

    # Make sure all the macros are the same as the parent macro (hash comp)
    # use increment file name function as a template
    
    
    return compare_file_copy_hashes(macro_path,)


def num_macros(macro_path: str) -> int:
    """
    num_macros returns the number of macro copies from a given parent macro

    Args:
        macro_path (str): path of the parent macro

    Returns:
        int: number of macro copies including the parent
    """
    raise Warning("Function not implemented yet")
    return 0


def multiply_macros(macro_path: str, num_to_multiply_to: int, already_copied: list[str]) -> None:
    """
    multiply_macros  mutilpies a parent macro

    Args:
        macro_path (str): parent macro to copy
        num_to_multiply_to (int): number of copies to have
        already_copied (list[str]): list of already copied macros
    """
    raise Warning("Function not implemented yet")
    # Make sure there are num_to_multiply_to number of macro copies avaible
    # Make macro folder
    # add "Macro (n).dotm" is added to .gitignore

    # if macro copy in already_copied, skip


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

    # TODO:
    # check if the number of macros is less than the number of docs
    macro_paths: list[str] = get_macro_paths(macro_path=template_path)
    if len(macro_paths) < len(doc_paths):
        multiply_macros(
            macro_path=template_path,
            num_to_multiply_to=len(doc_paths),
            already_copied=macro_paths,
        )

    RunWordMacro.run_word_macro(
        doc_path=doc_paths,
        macro_name=macro,
        template_path=template_path,
        wordVisible=visible,
    )

    return


if __name__ == "__main__":
    from data.hidden.files import FILES  # This can be removed

    # Example usage
    files: list[str] | str = FILES[0]
    if isinstance(files, str):
        files = [files]  # if str, throw in a list of itself
    update_doc_properties(doc_paths=files)

    # Example usage
    file1 = "path/to/first/file"
    file2 = "path/to/second/file"
    print(f"File hashes are {'the same' if compare_file_copy_hashes(file1, file2) else 'different'}")
