"""XML names used by the package and WordprocessingML helpers."""

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}
W = "{" + NS["w"] + "}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def local_name(element: etree._Element) -> str:
    """Return an element's local name, ignoring XML comments and processing instructions."""
    return etree.QName(element).localname if isinstance(element.tag, str) else ""
