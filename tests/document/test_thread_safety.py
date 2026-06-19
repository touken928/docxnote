"""同一 DocxDocument 在多线程下的串行化访问"""

import threading
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO

from lxml import etree

from docxnote import DocxDocument, Paragraph
from docxnote.namespaces import NS


def _comment_count_in_output(docx_bytes: bytes) -> int:
    with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
        root = etree.fromstring(z.read("word/comments.xml"))
    return len(root.findall(f"{{{NS['w']}}}comment"))


def test_concurrent_comments_same_paragraph(simple_doc):
    doc = DocxDocument.parse(simple_doc)
    paras = [b for b in doc.blocks() if isinstance(b, Paragraph)]
    assert len(paras) >= 1
    p = paras[0]
    text = p.text
    assert len(text) >= 2

    n = 40

    def work(i: int) -> None:
        # 不同区间，避免实现层对重叠范围的假设干扰本测试
        start = i % (len(text) - 1)
        end = start + 1
        p.comment(f"t{i}", start=start, end=end, author=f"a{i}")

    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = [pool.submit(work, i) for i in range(n)]
        for f in as_completed(futures):
            f.result()

    out = doc.render()
    assert len(out) > 100
    assert _comment_count_in_output(out) == n


def test_concurrent_comment_and_render(simple_doc):
    doc = DocxDocument.parse(simple_doc)
    paras = [b for b in doc.blocks() if isinstance(b, Paragraph)]
    p = paras[0]
    text = p.text
    span = max(1, len(text) - 1)

    errors: list[BaseException] = []

    def add_many() -> None:
        try:
            for i in range(30):
                start = i % span
                p.comment(f"x{i}", start=start, end=start + 1, author="t")
        except BaseException as e:
            errors.append(e)

    def render_many() -> None:
        try:
            for _ in range(20):
                b = doc.render()
                assert len(b) > 100
        except BaseException as e:
            errors.append(e)

    t1 = threading.Thread(target=add_many)
    t2 = threading.Thread(target=render_many)
    t1.start()
    t2.start()
    t1.join()
    t2.join()

    assert not errors
    assert len(doc.render()) > 100
