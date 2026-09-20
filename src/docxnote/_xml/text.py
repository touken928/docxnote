"""One text coordinate system for paragraph reads, ranges, and run splitting."""

from collections.abc import Iterator
from copy import deepcopy
from dataclasses import dataclass

from lxml import etree

from .namespaces import XML_SPACE, local_name


@dataclass(frozen=True)
class RunSpan:
    element: etree._Element
    parent: etree._Element
    start: int
    end: int


def iter_inline_elements(container: etree._Element) -> Iterator[etree._Element]:
    """Walk wrappers, stopping at runs and range markers.

    A run is a leaf in the host paragraph's coordinate system: drawings and
    text boxes inside it must not contribute nested text or comment markers.
    """
    for child in container:
        if local_name(child) in {"r", "commentRangeStart", "commentRangeEnd"}:
            yield child
        else:
            yield from iter_inline_elements(child)


def iter_runs(container: etree._Element) -> Iterator[etree._Element]:
    for element in iter_inline_elements(container):
        if local_name(element) == "r":
            yield element


def text_atom(element: etree._Element) -> str:
    """Map a direct run child to its contribution to the text view."""
    match local_name(element):
        case "t":
            return element.text or ""
        case "br":
            return "\n"
        case "tab":
            return "\t"
        case _:
            return ""


def run_text(run: etree._Element) -> str:
    return "".join(text_atom(child) for child in run)


def paragraph_text(element: etree._Element) -> str:
    return "".join(run_text(run) for run in iter_runs(element))


def text_length(element: etree._Element) -> int:
    if local_name(element) == "r":
        return len(run_text(element))
    return sum(len(run_text(run)) for run in iter_runs(element))


def iter_run_spans(container: etree._Element) -> Iterator[RunSpan]:
    offset = 0
    for run in iter_runs(container):
        end = offset + len(run_text(run))
        parent = run.getparent()
        assert parent is not None
        yield RunSpan(run, parent, offset, end)
        offset = end


def set_text(element: etree._Element, text: str) -> None:
    """Preserve leading/trailing whitespace when Word reads a text node."""
    element.text = text
    if text != text.strip():
        element.set(XML_SPACE, "preserve")


def split_run_at_boundary(paragraph: etree._Element, boundary: int) -> None:
    for span in iter_run_spans(paragraph):
        if span.start < boundary < span.end:
            split_run(span, boundary - span.start)
            return


def split_run(span: RunSpan, offset: int) -> None:
    """Split a run without losing properties or non-text children."""
    if not 0 < offset < span.end - span.start:
        return
    run = span.element
    before = etree.Element(run.tag, attrib=dict(run.attrib), nsmap=run.nsmap)
    after = etree.Element(run.tag, attrib=dict(run.attrib), nsmap=run.nsmap)
    position = 0
    for child in run:
        tag = local_name(child)
        if tag == "rPr":
            before.append(deepcopy(child))
            after.append(deepcopy(child))
            continue
        text = text_atom(child)
        if tag == "t":
            split = max(0, min(offset - position, len(text)))
            for target, part in ((before, text[:split]), (after, text[split:])):
                if part:
                    node = deepcopy(child)
                    set_text(node, part)
                    target.append(node)
        else:
            # Zero-width drawings, fields, and references stay on the same
            # side of the boundary as their position in the original run.
            target = before if position < offset else after
            target.append(deepcopy(child))
        position += len(text)
    index = span.parent.index(run)
    span.parent.remove(run)
    span.parent.insert(index, before)
    span.parent.insert(index + 1, after)
