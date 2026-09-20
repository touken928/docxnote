"""DOCX ZIP parts and the relationships required by the comments part."""

import io
import posixpath
import zipfile

from lxml import etree

COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
DOCUMENT_PART = "word/document.xml"
DOCUMENT_RELS_PART = "word/_rels/document.xml.rels"
CONTENT_TYPES_PART = "[Content_Types].xml"


class DocxPackage:
    """Keep source parts intact and replace only parts owned by the comment engine."""

    def __init__(self, data: bytes) -> None:
        self._archive = zipfile.ZipFile(io.BytesIO(data))
        self.comments_part = "word/comments.xml"

    def read(self, path: str) -> bytes:
        return self._archive.read(path)

    def find_comments_part(self) -> str:
        """Resolve the comments relationship relative to the document part."""
        try:
            root = etree.fromstring(self.read(DOCUMENT_RELS_PART))
        except KeyError:
            return "word/comments.xml"
        for relationship in root:
            if relationship.get("Type") == COMMENTS_REL_TYPE:
                target = relationship.get("Target")
                if not target or relationship.get("TargetMode") == "External":
                    raise ValueError("comments relationship must target a package part")
                return posixpath.normpath(posixpath.join("/word", target)).lstrip("/")
        return "word/comments.xml"

    def render(
        self,
        document_xml: bytes,
        comments_xml: bytes | None,
        *,
        keep_comments: bool,
    ) -> bytes:
        include_comments = comments_xml is not None
        replacements = {
            DOCUMENT_PART: document_xml,
            self.comments_part: comments_xml,
            CONTENT_TYPES_PART: self._content_types(include_comments),
        }
        relationships = self._document_relationships(include_comments)
        if relationships is not None:
            replacements[DOCUMENT_RELS_PART] = relationships
        if not keep_comments or not include_comments:
            replacements[self._comments_relationships_part()] = None

        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
            for name in self._archive.namelist():
                if name not in replacements:
                    target.writestr(name, self.read(name))
            for name, content in replacements.items():
                if content is not None:
                    target.writestr(name, content)
        return output.getvalue()

    def _comments_relationships_part(self) -> str:
        return posixpath.join(
            posixpath.dirname(self.comments_part),
            "_rels",
            posixpath.basename(self.comments_part) + ".rels",
        )

    def _document_relationships(self, include_comments: bool) -> bytes | None:
        try:
            root = etree.fromstring(self.read(DOCUMENT_RELS_PART))
        except KeyError:
            if not include_comments:
                return None
            root = etree.Element(
                "Relationships",
                nsmap={
                    None: "http://schemas.openxmlformats.org/package/2006/relationships"
                },
            )
        has_comments = False
        for relationship in list(root):
            if relationship.get("Type") == COMMENTS_REL_TYPE:
                if include_comments:
                    has_comments = True
                else:
                    root.remove(relationship)
        if include_comments and not has_comments:
            etree.SubElement(
                root,
                "Relationship",
                attrib={
                    "Id": self._next_relationship_id(root),
                    "Type": COMMENTS_REL_TYPE,
                    "Target": posixpath.relpath(self.comments_part, "word"),
                },
            )
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8")

    @staticmethod
    def _next_relationship_id(root: etree._Element) -> str:
        maximum = 0
        for relationship in root:
            identifier = relationship.get("Id", "")
            if identifier.startswith("rId"):
                try:
                    maximum = max(maximum, int(identifier[3:]))
                except ValueError:
                    pass
        return f"rId{maximum + 1}"

    def _content_types(self, include_comments: bool) -> bytes:
        root = etree.fromstring(self.read(CONTENT_TYPES_PART))
        namespace = root.nsmap.get(
            None, "http://schemas.openxmlformats.org/package/2006/content-types"
        )
        for override in list(root):
            if override.get("PartName") == "/" + self.comments_part:
                root.remove(override)
        if include_comments:
            etree.SubElement(
                root,
                f"{{{namespace}}}Override",
                attrib={
                    "PartName": "/" + self.comments_part,
                    "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
                },
            )
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8")
