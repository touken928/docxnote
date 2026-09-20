"""Path syntax is separate from document-dependent structural validation."""

import pytest

from docxnote import build_segment, comment_path, join_path, parse_path


@pytest.mark.parametrize("kind,index", [("p", 0), ("t", 3), ("r", 12), ("c", 7)])
def test_build_segment(kind, index):
    assert build_segment(kind, index) == f"{kind}:{index}"


@pytest.mark.parametrize("kind,index", [("x", 0), ("p", -1)])
def test_invalid_segment(kind, index):
    with pytest.raises(ValueError):
        build_segment(kind, index)


@pytest.mark.parametrize(
    "parts,expected",
    [
        (("t:0", "r:1", "c:2", "p:0"), "t:0/r:1/c:2/p:0"),
        (("", "p:0"), "p:0"),
        (("t:0", ""), "t:0"),
        (("", ""), ""),
    ],
)
def test_join_path_ignores_empty_fragments(parts, expected):
    assert join_path(*parts) == expected


@pytest.mark.parametrize("path,identifier", [("p:0", 3), ("t:0/r:0/c:0/p:1", 17)])
def test_comment_path(path, identifier):
    assert comment_path(path, identifier) == f"{path}#{identifier}"


def test_negative_comment_id():
    with pytest.raises(ValueError):
        comment_path("p:0", -1)


@pytest.mark.parametrize(
    "path,segments,identifier",
    [
        ("p:0", [("p", 0)], None),
        ("t:5", [("t", 5)], None),
        ("t:0/r:1/c:2/p:0", [("t", 0), ("r", 1), ("c", 2), ("p", 0)], None),
        ("t:0/r:1/c:2/p:0#9", [("t", 0), ("r", 1), ("c", 2), ("p", 0)], 9),
    ],
)
def test_parse_path(path, segments, identifier):
    assert parse_path(path) == (segments, identifier)


@pytest.mark.parametrize(
    "path", ["", "p:0#", "x:0", "p0", "p:abc", "p:-1", "p:0#bad", "p:0#-1", "/"]
)
def test_invalid_path(path):
    with pytest.raises(ValueError):
        parse_path(path)


def test_non_string_path():
    with pytest.raises(TypeError):
        parse_path(123)  # type: ignore[arg-type, ty:invalid-argument-type]
