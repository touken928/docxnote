"""Internal traversal of blocks through transparent content controls."""

from .namespaces import NS


def iter_block_elements(container):
    """Yield paragraphs/tables without descending into their own contents."""
    for child in container:
        if child.tag in {f"{{{NS['w']}}}p", f"{{{NS['w']}}}tbl"}:
            yield child
        elif child.tag == f"{{{NS['w']}}}sdt":
            content = child.find("./w:sdtContent", NS)
            if content is not None:
                yield from iter_block_elements(content)
