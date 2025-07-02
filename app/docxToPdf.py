#!/usr/bin/env python3.12.10
# -*- coding: utf-8 -*-

"""
# @ Author: Aaron Shackelford
# @ Create Time: 2025-06-13 16:15:39
# @ Description:

This script is being used to test the Docx to PDF api Aaron created in Power Automate

# @ Modified by: Aaron Shackelford
# @ Modified time: 2025-06-13 16:40:49
"""

import base64
import json
import os
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from pathlib import Path

import requests
from dotenv import load_dotenv
from requests import HTTPError

PROJECT_PATH: Path = Path(__file__).resolve().parent
DATA_PATH: Path = PROJECT_PATH / "data"
ENV_LOCATION: Path = PROJECT_PATH / ".env"

load_dotenv(dotenv_path=ENV_LOCATION)
GET_AIC_API_URL: str = os.getenv(key="PDF_CONVERSION_API", default="")
GET_AIC_USERS_API: str = os.getenv(key="PDF_CONVERSION_AUTH", default="")
FILE_NAME: str = os.getenv(key="FILE_NAME", default="")


def get_pdf_binaries(docx_bin: BytesIO) -> BytesIO:
    """
    Converts a DOCX file (provided as a BytesIO object) to a PDF by sending it to an external API.
    Args:
        docx_bin (BytesIO): The binary content of the DOCX file to be converted.
    Returns:
        BytesIO: The binary content of the resulting PDF file.
    Raises:
        ValueError: If required API URLs are not set or if the API response is empty.
        HTTPError: If the API request fails with an HTTP error.
        Exception: For any other unexpected errors during the process.
    Notes:
        - Requires the global variables GET_AIC_API_URL and GET_AIC_USERS_API to be set.
        - The DOCX file is base64-encoded and sent as JSON to the API.
        - The API is expected to return a base64-encoded PDF.
    """
    docx_bin.seek(0)
    docx_base64: str = base64.b64encode(docx_bin.read()).decode("utf-8")
    text = ""
    if not GET_AIC_API_URL and not GET_AIC_USERS_API:
        raise ValueError(f"GET_AIC_API_URL and GET_AIC_USERS_API must have values. | {GET_AIC_API_URL} | {GET_AIC_USERS_API}")
    else:
        try:
            print("sending data to power automate")
            response: requests.Response = requests.Response()
            try:
                data: dict[str, str] = {
                    "Authorization": GET_AIC_USERS_API,
                    "file": docx_base64,
                }
                payload: str = json.dumps(data)

                # Set the headers
                headers: dict[str, str] = {"Content-Type": "application/json"}

                response = requests.post(
                    url=GET_AIC_API_URL,
                    headers=headers,
                    data=payload,
                    timeout=300,
                )
                response.raise_for_status()  # Raises HTTPError for bad responses
            except HTTPError as e:
                # Request error occurred
                raise
            except Exception as e:
                # unexpected error occurred
                raise
            else:
                print("got data from power automate")
                text: str = response.text

        except Exception:
            raise
        else:
            if text:
                pdf_data: BytesIO = BytesIO(initial_bytes=base64.b64decode(text))
            else:
                raise ValueError("No value to convert")
    return pdf_data


def file_to_pdf(in_path: Path, out_path: Path) -> None:
    docx: Path = in_path
    docx_bin: BytesIO = BytesIO()
    with open(file=docx, mode="rb") as file:
        docx_bin = BytesIO(initial_bytes=file.read())

    try:
        pdf_bin: BytesIO = get_pdf_binaries(docx_bin=docx_bin)
        os.makedirs(out_path.parent, exist_ok=True)
        with open(file=out_path, mode="wb") as file:
            pdf_bin.seek(0)
            file.write(pdf_bin.read())
    except Exception as e:
        print(f"{e}\n{in_path}")