import xml.etree.ElementTree as ET
from typing import Any
import os
import sys

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)
from data.hidden.files import XML_FILES

# Parse an XML file
tree: ET.ElementTree = ET.parse(source=XML_FILES[0])
root: ET.Element | Any = tree.getroot()

# # Access elements
# for child in root:
#     print(child.tag, child.attrib)

# Find all {http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent elements
txbx_contents: list[ET.Element] | Any = root.findall(
    ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent"
)

# Iterate through each txbxContent and find all {http://schemas.openxmlformats.org/wordprocessingml/2006/main}t elements
for txbx_content in txbx_contents:
    t_elements: list[ET.Element] | Any = txbx_content.findtext(r' DOCPROPERTY  "Company Name"  \* MERGEFORMAT ')
    # t_elements: list[ET.Element] | Any = txbx_content.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText')
    if t_elements:
        t0_elements: list[ET.Element] | Any = txbx_content.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
        )
        for t_elem in t0_elements:
            print(t_elem.tag, t_elem.text)
    print("=" * 10)
