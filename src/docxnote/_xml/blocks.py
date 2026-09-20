"""Block traversal through transparent content controls."""

from collections.abc import Iterator

from lxml import etree

from .namespaces import NS, W


def iter_block_elements(container: etree._Element) -> Iterator[etree._Element]:
    """Yield paragraphs and tables without descending into their contents."""
    for child in container:
        if child.tag in {W + "p", W + "tbl"}:
            yield child
        elif child.tag == W + "sdt":
            content = child.find("./w:sdtContent", NS)
            if content is not None:
                yield from iter_block_elements(content)
