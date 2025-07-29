import io
import zipfile
from xml.etree import ElementTree as ET


def update_custom_properties(docx_bin: bytes, properties: dict[str, str]) -> bytes:
    in_memory_buffer = io.BytesIO(initial_bytes=docx_bin)
    out_memory_buffer = io.BytesIO()
    with zipfile.ZipFile(file=in_memory_buffer, mode="r") as docx_in:
        with zipfile.ZipFile(file=out_memory_buffer, mode="w") as docx_out:
            for item in docx_in.infolist():
                if item.filename != "docProps/custom.xml":
                    docx_out.writestr(zinfo_or_arcname=item, data=docx_in.read(name=item.filename))
                else:
                    # Parse and modify custom.xml
                    xml_content: bytes = docx_in.read(name=item.filename)
                    tree: ET.Element[str] = ET.fromstring(text=xml_content)
                    ns: dict[str, str] = {
                        "cp": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
                        "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                    }
                    for prop in tree.findall(path="cp:property", namespaces=ns):
                        prop_name: str | None = prop.get("name")
                        if prop_name in properties:
                            vt_elem: ET.Element[str] | None = prop.find(path="./vt:lpwstr", namespaces=ns)
                            if vt_elem is not None:
                                vt_elem.text = properties[prop_name]
                    modified_xml: bytes = ET.tostring(element=tree, encoding="UTF-8", xml_declaration=True)
                    docx_out.writestr(zinfo_or_arcname=item.filename, data=modified_xml)

    out_memory_buffer.seek(0)
    return out_memory_buffer.read()
