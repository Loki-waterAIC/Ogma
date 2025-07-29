#!/usr/bin/env python3
"""
 # @ Author: Aaron Shackelford
 # @ Create Time: 2025-07-17 15:32:24
 # @ Description:

This file takes in a docx binary and updates all the visual text properties
in the docx file to their current property values.

Initially created by ChatGPT 4o on 20250717
https://chatgpt.com/s/t_6879818920488191a3ec713aa3b4920e

Reviewed and modified to suit our ogma's needs.

 # @ Modified by: Aaron Shackelford
 # @ Modified time: 2025-07-17 16:01:19
"""

import io
from re import T
import zipfile
from logging import Logger, getLogger
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.etree.ElementTree import Element
from zipfile import BadZipFile

logger: Logger = getLogger("streamline." + __name__)


def __extract_custom_properties(xml: bytes) -> dict[str, str]:
    """
    Extracts custom properties from the DOCX custom properties XML.

    Args:
        xml (bytes): The XML content of the custom properties file.

    Returns:
        dict[str, str]: A dictionary mapping property names to their string values.
    """
    tree: ET.Element[str] = ET.fromstring(text=xml)
    ns: dict[str, str] = {
        "cp": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
        "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    }
    return {
        prop.attrib["name"]: prop.find(path="./vt:lpwstr", namespaces=ns).text or ""  # type: ignore
        for prop in tree.findall(path="cp:property", namespaces=ns)
        if prop.find(path="./vt:lpwstr", namespaces=ns) is not None
    }


def __update_docproperty_fields(xml: bytes, property_map: dict[str, str]) -> bytes:
    """
    Updates all DOCPROPERTY fields in the given DOCX XML with values from the property map.

    Args:
        xml (bytes): The XML content of a DOCX part (e.g., document, header, footer).
        property_map (dict[str, str]): A dictionary mapping property names to their updated values.

    Returns:
        bytes: The updated XML content as bytes.
    """
    tree: ET.Element[str] = ET.fromstring(xml)
    ns: dict[str, str] = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    # Simple fields: <w:fldSimple w:instr=" DOCPROPERTY  &quot;Project Name&quot;  \* MERGEFORMAT ">
    for field in tree.findall(path=".//w:fldSimple", namespaces=ns):
        instruction: str = field.attrib.get(f'{ns["w"]}instr', "")
        for property, value in property_map.items():
            if f'DOCPROPERTY  "{property}"' in instruction or f"DOCPROPERTY  &quot;{property}&quot;" in instruction:
                text_elem: ET.Element[str] | None = field.find(path=".//w:t", namespaces=ns)
                if text_elem is not None:
                    text_elem.text = value

    # Complex fields: <w:fldChar/> ... <w:instrText> DOCPROPERTY ... </w:instrText> ... <w:t>value</w:t>
    document_paragraphs: list[Element[str]] = tree.findall(path=".//w:p", namespaces=ns)
    for paragraphs in document_paragraphs:
        instruction = ""
        inside_instruction = False
        target_proproperty = ""
        for node in paragraphs:
            if node.tag == f'{{{ns["w"]}}}r':
                for child in node:
                    if child.tag == f'{{{ns["w"]}}}fldChar':
                        fld_type: str | None = child.attrib.get(f'{ns["w"]}fldCharType')
                        if fld_type == "begin":
                            inside_instruction = True
                        elif fld_type == "end":
                            inside_instruction = False
                            target_proproperty = ""
                    elif inside_instruction and child.tag == f'{{{ns["w"]}}}instrText':
                        instruction = child.text or ""
                        for property in property_map:
                            if f'DOCPROPERTY  "{property}"' in instruction:
                                target_proproperty: str = property
                    elif not inside_instruction and child.tag == f'{{{ns["w"]}}}t' and target_proproperty:
                        child.text = property_map[target_proproperty]

    return ET.tostring(element=tree, encoding="UTF-8", xml_declaration=True)


def update_docx_properties(docx_bin: bytes) -> tuple[bytes, bool]:
    """
    Updates all visual text properties in a DOCX file to their current property values.

    Args:
        docx_file (bytes): The binary content of the DOCX file.

    Returns:
        bytes: The updated DOCX file as binary content.
    """
    in_memory_buffer = io.BytesIO(initial_bytes=docx_bin)
    out_memory_buffer = io.BytesIO()
    bad_briareus_file: bool = False
    with zipfile.ZipFile(file=in_memory_buffer, mode="r") as docx_in:
        with zipfile.ZipFile(file=out_memory_buffer, mode="w") as docx_out:
            custom_props: dict[str, str] = {}

            # First, extract custom properties
            for item in docx_in.infolist():
                if item.filename == "docProps/custom.xml":
                    custom_props = __extract_custom_properties(xml=docx_in.read(name=item))
                docx_out.writestr(zinfo_or_arcname=item, data=docx_in.read(name=item))

            if custom_props:
                # Now replace XMLs where fields are used
                parts_to_patch: list[str] = [
                    f
                    for f in docx_in.namelist()
                    if f.startswith("word/")
                    and f.endswith(".xml")
                    and ("header" in f or "footer" in f or f == "word/document.xml")
                ]

                for filename in parts_to_patch:
                    try:
                        xml: bytes = docx_in.read(name=filename)
                    except BadZipFile:
                        # BUG!
                        #   The following note is incorrect, this bug is produced from 
                        #   the binaries from the property update function.
                        # TODO:
                        #   Fix the bug produced by the property update function!
                        # NOTE:
                        # In the first version of Briareus, we created Word Documents with
                        # incorrect xml data. This is an expected error to occur when
                        # these bad files docx files are encountered.

                        # For now, skip this xml and try the next one.
                        bad_briareus_file = True
                        continue
                    updated_xml: bytes = __update_docproperty_fields(xml=xml, property_map=custom_props)
                    docx_out.writestr(zinfo_or_arcname=docx_in.getinfo(name=filename), data=updated_xml)

    out_memory_buffer.seek(0)
    return (out_memory_buffer.read(), bad_briareus_file)
