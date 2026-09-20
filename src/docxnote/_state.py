"""Shared document state; every view uses this single reentrant lock."""

from dataclasses import dataclass, field
from threading import RLock
from typing import Protocol

from lxml import etree

from ._comment_store import CommentStore
from ._package import DocxPackage, DOCUMENT_PART
from ._xml.namespaces import NS
from ._xml.ranges import strip_comment_markers


@dataclass
class DocumentState:
    package: DocxPackage
    comments: CommentStore = field(default_factory=CommentStore)
    lock: RLock = field(default_factory=RLock)
    root: etree._Element | None = None
    body: etree._Element | None = None
    keep_comments: bool = False

    def load(self, *, keep_comments: bool) -> None:
        """Initialize XML and comment policy before the document is published."""
        self.root = etree.fromstring(self.package.read(DOCUMENT_PART))
        self.body = self.root.find(".//w:body", NS)
        self.keep_comments = keep_comments
        self.package.comments_part = self.package.find_comments_part()
        if keep_comments:
            try:
                data = self.package.read(self.package.comments_part)
            except KeyError:
                return
            self.comments = CommentStore.parse(data)
        else:
            strip_comment_markers(self.root)


class DocumentOwner(Protocol):
    """A document or view that can supply shared state to a child view."""

    _state: DocumentState
