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

INTXBXCONT = 2
DOCPROP = 1
TEXT = 0


def traverse_tree(
    element: ET.Element,
    marks: list[bool] = [False] * 3,
    els: list[ET.Element | None] = [None] * 2,
) -> None:  # FMT: off
    if element.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent":
        marks[INTXBXCONT] = True
    if marks[INTXBXCONT] and element.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText":
        if element.text:
            marks[DOCPROP] = True
            els[DOCPROP] = element
    if marks[INTXBXCONT] and element.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t":
        if element.text:
            marks[TEXT] = True
            els[TEXT] = element

    # if all elements in marks is true
    if all(marks):
        # assume they have values
        docprop: str = str(els[DOCPROP].text)
        text: str = str(els[TEXT].text)
        print(docprop + " " + text)
        return

    for child in element:
        traverse_tree(element=child, marks=marks, els=els)


traverse_tree(element=root)
