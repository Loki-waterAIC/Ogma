import os
import sys

from lxml import etree

# project path
OGMA_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if OGMA_PATH not in sys.path:
    sys.path.append(OGMA_PATH)
from data.hidden.files import XML_FILES

# Parse an XML file
tree = etree.parse(source=XML_FILES[0])
root = tree.getroot()

# Access elements
for child in root:
    print(child.tag, child.attrib)

# Find specific elements using XPath
for elem in root.xpath("//tag_name"):
    print(elem.text)
