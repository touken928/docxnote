"""Generate documents and inspect ZIP parts without invoking docxnote internals."""

from io import BytesIO
import zipfile

from docx import Document
from docx.document import Document as WordDocument
from lxml import etree


def save_docx(document: WordDocument) -> bytes:
    buffer = BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def build_docx(paragraph_texts: list[str]) -> bytes:
    document = Document()
    for text in paragraph_texts:
        document.add_paragraph(text)
    return save_docx(document)


def read_parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(BytesIO(data)) as archive:
        names = archive.namelist()
        assert len(names) == len(set(names)), "duplicate ZIP entries"
        return {name: archive.read(name) for name in names}


def write_parts(parts: dict[str, bytes]) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return buffer.getvalue()


def xml_part(data: bytes, name: str = "word/document.xml") -> etree._Element:
    with zipfile.ZipFile(BytesIO(data)) as archive:
        return etree.fromstring(archive.read(name))
