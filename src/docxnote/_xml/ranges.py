"""Read and insert comment anchors in host-paragraph text coordinates."""

from dataclasses import dataclass

from lxml import etree

from .namespaces import NS, W, local_name
from .text import iter_inline_elements, run_text, split_run_at_boundary, text_length


class UnsupportedCommentRangeError(ValueError):
    """A comment crosses paragraphs or lacks paired range markers.

    Only range views raise this error. Parsing and rendering can still preserve
    the original XML without interpreting its anchors.
    """


@dataclass(frozen=True)
class CommentRange:
    comment_id: int
    start: int
    end: int
    sequence: int


@dataclass(frozen=True)
class RangeStart:
    offset: int
    sequence: int


@dataclass(frozen=True)
class InsertionPoint:
    parent: etree._Element
    index: int


def collect_comment_ranges(paragraph: etree._Element, path: str) -> list[CommentRange]:
    """Collect closed ranges in start-marker order; reject unsupported anchors."""
    ranges: list[CommentRange] = []
    open_starts: dict[int, RangeStart] = {}
    position = 0
    sequence = 0
    for element in iter_inline_elements(paragraph):
        tag = local_name(element)
        if tag == "r":
            position += len(run_text(element))
            continue
        comment_id = _marker_id(element)
        if comment_id is None:
            continue
        if tag == "commentRangeStart":
            if comment_id not in open_starts:
                open_starts[comment_id] = RangeStart(position, sequence)
                sequence += 1
        elif tag == "commentRangeEnd":
            start = open_starts.pop(comment_id, None)
            if start is None:
                raise UnsupportedCommentRangeError(
                    f"comment range {comment_id} on paragraph '{path}' has "
                    "a commentRangeEnd without a matching commentRangeStart "
                    "in this paragraph; comment ranges that cross paragraphs "
                    "or are left unclosed are not supported by the range view"
                )
            ranges.append(
                CommentRange(comment_id, start.offset, position, start.sequence)
            )
    if open_starts:
        unclosed = ", ".join(str(comment_id) for comment_id in open_starts)
        raise UnsupportedCommentRangeError(
            f"comment range(s) [{unclosed}] on paragraph '{path}' "
            "have no matching commentRangeEnd in this paragraph; comment "
            "ranges that cross paragraphs or are left unclosed are not "
            "supported by the range view"
        )
    ranges.sort(key=lambda item: (item.start, item.sequence))
    return ranges


def _marker_id(element: etree._Element) -> int | None:
    value = element.get(W + "id")
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def insert_comment_markers(
    paragraph: etree._Element, comment_id: int, start: int, end: int
) -> None:
    """Insert both anchors and the reference after splitting boundary runs."""
    split_run_at_boundary(paragraph, end)
    split_run_at_boundary(paragraph, start)
    attributes = {W + "id": str(comment_id)}
    start_marker = etree.Element(W + "commentRangeStart", attrib=attributes)
    end_marker = etree.Element(W + "commentRangeEnd", attrib=attributes)
    reference = etree.Element(W + "r")
    etree.SubElement(reference, W + "commentReference", attrib=attributes)
    location = _find_boundary(paragraph, start)
    location.parent.insert(location.index, start_marker)
    location = _find_boundary(paragraph, end)
    location.parent.insert(location.index, end_marker)
    location.parent.insert(location.index + 1, reference)


def _find_boundary(
    container: etree._Element, boundary: int, position: int = 0
) -> InsertionPoint:
    for index, child in enumerate(container):
        tag = local_name(child)
        if tag in {"commentRangeStart", "commentRangeEnd"}:
            continue
        length = text_length(child)
        if length == 0:
            continue
        if tag == "r":
            if boundary <= position:
                return InsertionPoint(container, index)
        elif position <= boundary <= position + length:
            return _find_boundary(child, boundary, position)
        position += length
    return InsertionPoint(container, len(container))


def strip_comment_markers(root: etree._Element) -> None:
    """Remove anchors throughout the XML, including non-addressable text boxes."""
    for tag in ("commentRangeStart", "commentRangeEnd"):
        for element in root.findall(f".//w:{tag}", NS):
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
    for reference in root.findall(".//w:commentReference", NS):
        run = reference.getparent()
        if run is None:
            continue
        run.remove(reference)
        if len(run) == 0 and run.text is None and not run.tail:
            parent = run.getparent()
            if parent is not None:
                parent.remove(run)
