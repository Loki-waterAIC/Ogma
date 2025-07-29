import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path


def extract_custom_properties(xml: bytes) -> dict[str, str]:
    tree: ET.Element[str] = ET.fromstring(text=xml)
    ns: dict[str, str] = {
        "cp": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
        "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    }

    # dictionary comprehension
    return {
        prop.attrib["name"]: prop.find("./vt:lpwstr", ns).text or ""  # type: ignore
        for prop in tree.findall("cp:property", ns)
        if prop.find("./vt:lpwstr", ns) is not None
    }
