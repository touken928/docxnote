"""测试 docxnote.paths 工具函数"""

import pytest

from docxnote import build_segment, comment_path, join_path, parse_path


class TestBuildSegment:
    def test_simple(self):
        assert build_segment("p", 0) == "p:0"
        assert build_segment("t", 3) == "t:3"
        assert build_segment("r", 12) == "r:12"
        assert build_segment("c", 7) == "c:7"

    def test_invalid_kind(self):
        with pytest.raises(ValueError):
            build_segment("x", 0)

    def test_negative_index(self):
        with pytest.raises(ValueError):
            build_segment("p", -1)


class TestJoinPath:
    def test_join_segments(self):
        assert join_path("t:0", "r:1", "c:2", "p:0") == "t:0/r:1/c:2/p:0"

    def test_skip_empty(self):
        assert join_path("", "p:0") == "p:0"
        assert join_path("t:0", "") == "t:0"
        assert join_path("", "") == ""


class TestCommentPath:
    def test_build(self):
        assert comment_path("p:0", 3) == "p:0#3"
        assert comment_path("t:0/r:0/c:0/p:1", 17) == "t:0/r:0/c:0/p:1#17"

    def test_negative_id(self):
        with pytest.raises(ValueError):
            comment_path("p:0", -1)


class TestParsePath:
    def test_top_paragraph(self):
        segs, cid = parse_path("p:0")
        assert segs == [("p", 0)]
        assert cid is None

    def test_top_table(self):
        segs, cid = parse_path("t:5")
        assert segs == [("t", 5)]
        assert cid is None

    def test_deep(self):
        segs, cid = parse_path("t:0/r:1/c:2/p:0")
        assert segs == [("t", 0), ("r", 1), ("c", 2), ("p", 0)]
        assert cid is None

    def test_with_comment(self):
        segs, cid = parse_path("t:0/r:1/c:2/p:0#9")
        assert segs == [("t", 0), ("r", 1), ("c", 2), ("p", 0)]
        assert cid == 9

    def test_empty_path(self):
        with pytest.raises(ValueError):
            parse_path("")

    def test_missing_comment_id(self):
        with pytest.raises(ValueError):
            parse_path("p:0#")

    def test_invalid_kind(self):
        with pytest.raises(ValueError):
            parse_path("x:0")

    def test_invalid_segment(self):
        with pytest.raises(ValueError):
            parse_path("p0")

    def test_invalid_index(self):
        with pytest.raises(ValueError):
            parse_path("p:abc")

    def test_non_str(self):
        with pytest.raises(TypeError):
            parse_path(123)  # type: ignore[arg-type]
