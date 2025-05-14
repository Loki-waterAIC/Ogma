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
tree: ET.ElementTree = ET.parse(source=XML_FILES[0]) # type: ignore
root: ET.Element | Any = tree.getroot()

# # Access elements
# for child in root:
#     print(child.tag, child.attrib)

# Find all {http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent elements
text_box_contents: list[ET.Element] | Any = root.findall(
    ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent"
)

# Iterate through each text box Content and find all {http://schemas.openxmlformats.org/wordprocessingml/2006/main}t elements
for text_box_content in text_box_contents:
    t_elements: list[ET.Element] | Any = text_box_content.findall(
        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
    )
    for t_elem in t_elements:
        print(t_elem.tag, t_elem.text)
