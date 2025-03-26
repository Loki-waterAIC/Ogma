import os
import sys
import xml.etree.ElementTree as ET
from typing import Any

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)

from data.hidden.files import XML_FILES


# Parse an XML file
tree: ET.ElementTree = ET.parse(source=XML_FILES[0])
root: ET.Element | Any = tree.getroot()


# Find all {http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent elements
txbx_contents = root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent")

# Iterate through each txbxContent
for txbx_content in txbx_contents:
    # Check if txbxContent contains {http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText
    if txbx_content.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText") is not None:
        # Find all elements with attributes ending with 't'
        for elem in txbx_content.iter():
            for attr in elem.attrib:
                if attr.endswith("t"):
                    print(elem.tag, elem.attrib, elem.text)
