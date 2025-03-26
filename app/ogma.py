#!/usr/bin/env python3.11.11
# -*- coding: utf-8 -*-
"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-03-12 13:19:04
# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-03-12 13:19:19
# @ Description:

 Ogma is a program that edits a word documents' propery values
"""


import argparse
import json
import os
import sys

# project path
OGMA_PATH: str = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)
    
import app.ogmaGlobal as ogmaGlobal
import app.ogmaScripts.runWordMacroWin as runWordMacroWin
from app.ogmaScripts.documentPropertyUpdateTool import document_properity_update_tool

def run_json_list(json_paths:list[str]):
    for path in json_paths:
        run_json(json_path=path)

def run_json(json_path: str) -> None:
    '''
        run_json processes the following json string:\n
        ```
        {
            "dotm_path": "./path.dotm",
            "macros": [
                "macroName0",
                "macroNameN"
            ],
            "singleFileMacro": true,
            "files": [
                "./file0.docx",
                "./fileN.docx"
            ],
            "doc_properties": {
                "BOK ID": "BOK ID {time}",
                "Document Name": "DOC NAME {time}",
                "Company Name": "CO NAME {time}",
                "Division": "DIV {time}",
                "Author": "AUTH {time}",
                "Company Address": "ADDR {time}",
                "Project Name": "PRJ NAME {time}",
                "Project Number": "PRJ ID {time}",
                "End Customer": "END CUST {time}",
                "Site Name": "SITE NAME {time}",
                "File Name": "FILE NAME {time}"
            }
        }
        ```

        Args:
            json_path (str): path to json object

        Raises:
            FileNotFoundError: no file found
            ValueError: file can't be opened
    '''
    try:
        # TODO
        # [ ] make sure windows paths are turned into WSL paths
        json_path = os.path.normpath(json_path)

        with open(json_path, 'r') as file:
            data: dict = dict(json.load(file))
            if ogmaGlobal.VERBOSE_LEVEL:
                print("Successfully parsed JSON file:")
                print(data)  # This will print the dictionary

            # MARK: detect doc_properties and files
            if all([(i in data) for i in ["files", "doc_properties"]]):
                # grab data
                files: list[str] = data["files"]
                doc_props: dict[str, str] = data["doc_properties"]
                if files and doc_props:
                    # run doc property update tool
                    # [ ] should we validate the data?
                    document_properity_update_tool(doc_paths=files, properties=doc_props)

            # if dotm_paths[0] to files[3] in data
            if all([(i in data) for i in ["dotm_path", "macros", "singleFileMacro", "files"]]):
                # grab data
                dotm_path: str = data['dotm_path']
                macros: list[str] = data['macros']
                singleFileMacro: bool = data["singleFileMacro"]
                files: list[str] = data['files']
                # run a macro
                runWordMacroWin.run_word_macro_on_files(doc_paths=files, macro_names=macros,
                                                        template_path=dotm_path, activeDocumentMacro=singleFileMacro)

    except FileNotFoundError:
        err_message: str = f"File not found: {args.json}"
        if ogmaGlobal.VERBOSE_LEVEL:
            print(err_message)
        raise FileNotFoundError(err_message)
    except json.JSONDecodeError:
        err_message: str = f"Error decoding JSON from file: {args.json}"
        if ogmaGlobal.VERBOSE_LEVEL:
            print(err_message)
        raise ValueError(err_message)


if __name__ == "__main__":
    # MARK: START READING HERE

    # CLI program information
    program_name: str = "ogma.exe"
    program_description: str = (
        "Created by Aaron Shackelford\n"
        "Ogma is named after the Gallic god of writing, Ogma, also spelled Oghma or Ogmios from the Greek.\n"
        "Ogma can run as a service or a CLI tool to modify local Word files.\n"
        "Ogma takes in a JSON object as instructions on how to process Word files.\n"
        "Ogma is known for being detected as ransomware due to how it modifies Word files; and may need to be whitelisted.\n"
        "\nSee below for launch arguments.\n"
    )
    program_epilog: str = ""

    # defining CLI tool
    parser = argparse.ArgumentParser(prog=program_name, description=program_description,
                                     epilog=program_epilog, allow_abbrev=False, formatter_class=argparse.RawTextHelpFormatter)

    # adding CLI tool arguments
    # parser.add_argument('--json', '-j',  action="store", type=str, nargs=1,  help='Path to the JSON file to instruct ogma on what to do.')
    parser.add_argument('--jsonPath','--jsonPaths', '--json', '-j', dest="jsonPaths", action="store", type=str,
                        nargs='*', metavar='"./json_path.json"', help='Path(s) to the JSON file to instruct ogma on what to do.')
    parser.add_argument('--verbose', '-v', dest="verbose", action='store_true', help='Increase output verbosity')
    parser.add_argument('--version', "--v", dest="version", action='version', version='%(prog)s 'f'version {ogmaGlobal.APP_VERSION}') # FIX: FIX THE FORMATTING HERE TO MAKE IT RETURN A VERSION!!
    # parse inputs
    args: argparse.Namespace = parser.parse_args()
    
    # if no arguments, ogma.exe/py[1] --arg[2]
    if len(sys.argv) < 2:
        parser.print_help()
        sys.exit(1)

    ogmaGlobal.VERBOSE_LEVEL = int(args.verbose)

    if args.jsonPaths:
        run_json_list(json_paths=args.jsonPaths)
