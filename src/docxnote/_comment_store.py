"""Comment bodies and source XML, independent of paragraph anchors and ZIP paths."""

from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime, timezone

from lxml import etree

from ._xml.namespaces import NS, W
from ._xml.text import run_text, set_text


@dataclass(frozen=True)
class CommentRecord:
    comment_id: int
    text: str
    author: str
    date: datetime | None
    source_element: etree._Element | None = None


class CommentStore:
    """Own IDs, metadata, and source elements under the document's shared lock.

    The ordered records are authoritative; the lookup maps IDs to list positions.
    Original XML is kept separately from its lossy plain-text/date interpretation
    inside each record, so rendering does not reconstruct existing comments.
    """

    def __init__(self) -> None:
        self._records: list[CommentRecord] = []
        self._positions: dict[int, int] = {}
        self._root_template: etree._Element | None = None
        self._next_id = 0

    @classmethod
    def parse(cls, data: bytes) -> "CommentStore":
        store = cls()
        root = etree.fromstring(data)
        store._root_template = deepcopy(root)
        for element in root:
            raw_id = element.get(W + "id")
            if not raw_id:
                continue
            try:
                comment_id = int(raw_id)
            except ValueError:
                continue
            store._append(
                CommentRecord(
                    comment_id,
                    _comment_text(element),
                    element.get(W + "author", "") or "",
                    _parse_date(element.get(W + "date")),
                    deepcopy(element),
                )
            )
            store._next_id = max(store._next_id, comment_id + 1)
        return store

    def create(
        self, text: str, author: str, date: datetime | None = None
    ) -> CommentRecord:
        record = CommentRecord(
            self._next_id,
            text,
            author,
            date if date is not None else datetime.now().astimezone(),
        )
        self._append(record)
        self._next_id += 1
        return record

    def _append(self, record: CommentRecord) -> None:
        self._positions[record.comment_id] = len(self._records)
        self._records.append(record)

    def get(self, comment_id: int) -> CommentRecord | None:
        position = self._positions.get(comment_id)
        return self._records[position] if position is not None else None

    def render(self) -> bytes | None:
        if not self._records:
            return None
        if self._root_template is None:
            root = etree.Element(W + "comments", nsmap=NS)
        else:
            root = deepcopy(self._root_template)
            for child in list(root):
                root.remove(child)
        for record in self._records:
            # Match source ID lookup semantics even for duplicate IDs: the last
            # source element wins, while the original record order is retained.
            current = self.get(record.comment_id)
            assert current is not None
            if current.source_element is not None:
                root.append(deepcopy(current.source_element))
            else:
                root.append(_new_comment_element(record))
        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )


def _comment_text(element: etree._Element) -> str:
    # Comment bodies retain the existing descendant-paragraph interpretation;
    # the addressable host-paragraph view has a narrower traversal scope.
    return "\n".join(
        "".join(run_text(run) for run in paragraph.findall(".//w:r", NS))
        for paragraph in element.findall(".//w:p", NS)
    )


def _parse_date(value: str | None) -> datetime | None:
    """Interpret w:date without changing the preserved source attribute."""
    if not value or not value.strip():
        return None
    value = value.strip()
    if value.endswith("Z"):
        value = value[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def _format_date(value: datetime) -> str:
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _new_comment_element(record: CommentRecord) -> etree._Element:
    assert record.date is not None
    element = etree.Element(
        W + "comment",
        attrib={
            W + "id": str(record.comment_id),
            W + "author": record.author,
            W + "date": _format_date(record.date),
            W + "initials": record.author[0].upper() if record.author else "D",
        },
    )
    for line in record.text.split("\n"):
        paragraph = etree.SubElement(element, W + "p")
        run = etree.SubElement(paragraph, W + "r")
        for index, segment in enumerate(line.split("\t")):
            if index:
                etree.SubElement(run, W + "tab")
            if segment or "\t" not in line:
                set_text(etree.SubElement(run, W + "t"), segment)
    return element
