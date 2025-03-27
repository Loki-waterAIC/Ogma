import os
import zipfile
from concurrent.futures import ThreadPoolExecutor
import sys

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)


def unzip_file_multithreaded(zip_file_path: str, file_ext: list[str] = [".zip"], output_dir: str | None = None) -> str:
    """
    Unzips a ZIP file to the same directory in a multi-threaded fashion.

    Args:
        zip_file_path (str): The path to the ZIP file to be unzipped.

    Raises:
        FileNotFoundError: If the ZIP file does not exist.
        zipfile.BadZipFile: If the ZIP file is corrupted.
    """
    if not os.path.exists(zip_file_path):
        raise FileNotFoundError(f"[unzip_docx.unzip_file_multithreaded] The file {zip_file_path} does not exist.")

    if not output_dir:
        # init the output_dir directory
        # edge case
        zip_file_path = os.path.normpath(zip_file_path)
        period: int | None = zip_file_path.rfind(".")
        slash: int = zip_file_path.rfind("/")
        if slash < period:
            period = None
        # init
        output_dir = os.path.abspath(zip_file_path[:period])

        # remove ending
        for i in file_ext:
            if zip_file_path.endswith(i):
                output_dir = os.path.abspath(zip_file_path).removesuffix(i)
                break

    # make the dir
    os.makedirs(name=output_dir, exist_ok=True)

    with zipfile.ZipFile(file=zip_file_path, mode="r") as zip_ref:
        zip_objects: list[zipfile.ZipInfo] = zip_ref.infolist()
        tree: list[str] = zip_ref.namelist()
        tree.sort()
        tree = list(map(lambda x: x + "\n", tree))

        # with open(file= os.path.join(output_dir,"tree.txt"),mode="w",encoding="utf-8") as tree_file:
        #     tree_file.writelines(tree)

        def extract_member(zip_info: zipfile.ZipInfo) -> None:
            """
            Extracts a single member from the ZIP file.

            Args:
                zip_info (zipfile.ZipInfo): The ZIP file member to extract.
            """
            zip_ref.extract(member=zip_info, path=output_dir)

        with ThreadPoolExecutor(max_workers=1 if __debug__ else None) as executor:
            executor.map(extract_member, zip_objects)

    return output_dir


if __name__ == "__main__":
    from data.hidden import files

    # Example usage
    extns: list[str] = [".docx", ".docm"]
    proc: list[str] = []
    excl: list[str] = []
    if False:
        for file in os.listdir(r"."):
            file_path = os.path.normpath(file)
            for ext in extns:
                if file.endswith(ext):
                    try:
                        unzip_file_multithreaded(file_path, [ext])
                    except:
                        pass
                    proc.append(f"ran: {file_path}")
                    break
            else:
                excl.append(f"skipped: {file_path}")

        for i in proc:
            print(i)

        for i in excl:
            print(i)
    else:
        unzip_file_multithreaded(zip_file_path=files.FILES[0], file_ext=extns)
