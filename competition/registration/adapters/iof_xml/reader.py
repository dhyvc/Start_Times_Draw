import xml.etree.ElementTree as ET

_NS = "{http://www.orienteering.org/datastandard/3.0}"


def read_iof_xml(file_storage):
    """Parse an IOF results XML and yield raw competitor dicts.

    Each dict contains the PersonResult element and its class name.
    """
    tree = ET.parse(file_storage.stream)
    root = tree.getroot()
    for class_result in root.findall(f"{_NS}ClassResult"):
        class_el = class_result.find(f"{_NS}Class")
        class_name_el = class_el.find(f"{_NS}Name") if class_el is not None else None
        class_name = class_name_el.text if class_name_el is not None and class_name_el.text else ""
        for person_result in class_result.findall(f"{_NS}PersonResult"):
            yield {"element": person_result, "class_name": class_name}
