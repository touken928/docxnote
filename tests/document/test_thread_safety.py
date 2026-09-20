"""Shared locks protect writes, rendered snapshots, and paused iterators."""

from concurrent.futures import ThreadPoolExecutor
import threading

import pytest

from docxnote import DocxDocument
from tests.support.docx import build_docx


@pytest.fixture
def document():
    return DocxDocument.parse(build_docx(["ABCDEFGHIJ", "Second paragraph"]))


def test_concurrent_comments_keep_unique_ids_and_exact_ranges(document):
    paragraph = next(document.iter_paragraphs())

    def add(index):
        start = index % 9
        paragraph.comment(str(index), start, start + 1, author=f"a{index}")

    with ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(add, range(40)))
    reopened = DocxDocument.parse(document.render(), keep_comments=True)
    comments = reopened.comments()
    assert len({c.path for c in comments}) == len(comments) == 40
    assert {c.text: (c.start, c.end, c.author) for c in comments} == {
        str(index): (index % 9, index % 9 + 1, f"a{index}") for index in range(40)
    }


def test_concurrent_render_never_observes_partially_written_anchors(document):
    paragraph = next(document.iter_paragraphs())
    start = threading.Barrier(2)

    def add_many():
        start.wait(timeout=5)
        for index in range(30):
            paragraph.comment(str(index), index % 9, index % 9 + 1)

    def render_many():
        start.wait(timeout=5)
        for _ in range(20):
            snapshot = DocxDocument.parse(document.render(), keep_comments=True)
            comments = snapshot.comments()
            assert len({c.path for c in comments}) == len(comments)
            assert {int(c.text) for c in comments} == set(range(len(comments)))
            assert all(
                (c.start, c.end) == (int(c.text) % 9, int(c.text) % 9 + 1)
                for c in comments
            )

    with ThreadPoolExecutor(max_workers=2) as pool:
        futures = [pool.submit(add_many), pool.submit(render_many)]
        for future in futures:
            future.result(timeout=10)
    assert len(document.comments()) == 30


def test_iter_paragraphs_releases_lock_while_paused(document):
    iterator = document.iter_paragraphs()
    next(iterator)
    finished = threading.Event()

    def add_comment():
        next(document.iter_paragraphs()).comment("concurrent", 0, 1)
        finished.set()

    worker = threading.Thread(target=add_comment, daemon=True)
    worker.start()
    try:
        assert finished.wait(timeout=5)
    finally:
        iterator.close()
        worker.join(timeout=5)
    assert not worker.is_alive()
