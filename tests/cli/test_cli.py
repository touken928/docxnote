"""docxnote CLI 的端到端测试"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docxnote import DocxDocument
from docxnote.cli import main


def _write(tmp_path: Path, name: str, data: bytes) -> Path:
    p = tmp_path / name
    p.write_bytes(data)
    return p


class TestListCommand:
    def test_plain_output_contains_paragraph_paths(self, capsys, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        rc = main(["list", str(infile)])
        assert rc == 0
        out = capsys.readouterr().out.splitlines()
        assert out
        assert all(line.startswith(("p:", "t:")) for line in out)
        # 至少应看到 p:0
        assert any(line == "p:0" for line in out)

    def test_json_output_schema(self, capsys, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        rc = main(["list", str(infile), "--json"])
        assert rc == 0
        payload = json.loads(capsys.readouterr().out)
        assert isinstance(payload, list) and payload
        assert {"path", "text"} <= set(payload[0].keys())

    def test_text_preview_is_truncated(self, capsys, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        rc = main(["list", str(infile), "--text", "--text-limit", "3", "--json"])
        assert rc == 0
        payload = json.loads(capsys.readouterr().out)
        # 至少有一个文本被截断
        assert any(p["text"].endswith("…") for p in payload)


class TestShowCommand:
    def test_show_paragraph_json(self, capsys, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        rc = main(["show", str(infile), "p:0", "--json"])
        assert rc == 0
        data = json.loads(capsys.readouterr().out)
        assert data["type"] == "paragraph"
        assert data["path"] == "p:0"
        assert "text" in data
        assert isinstance(data["comments"], list)

    def test_show_table_shape(self, capsys, tmp_path, table_doc):
        infile = _write(tmp_path, "in.docx", table_doc)
        rc = main(["show", str(infile), "t:0", "--json"])
        assert rc == 0
        data = json.loads(capsys.readouterr().out)
        assert data["type"] == "table"
        assert data["shape"] == [3, 3]

    def test_show_cell_bounds(self, capsys, tmp_path, table_doc):
        infile = _write(tmp_path, "in.docx", table_doc)
        rc = main(["show", str(infile), "t:0/r:1/c:2", "--json"])
        assert rc == 0
        data = json.loads(capsys.readouterr().out)
        assert data["type"] == "cell"
        assert data["path"] == "t:0/r:1/c:2"
        assert len(data["bounds"]) == 4

    def test_show_invalid_path_returns_error(self, capsys, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        rc = main(["show", str(infile), "p:9999", "--json"])
        assert rc == 2
        err = capsys.readouterr().err
        assert "error" in err.lower()


class TestAnnotateCommand:
    def test_single_annotation(self, capsys, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        outfile = tmp_path / "out.docx"

        rc = main(
            [
                "annotate",
                str(infile),
                str(outfile),
                "--path",
                "p:1",
                "--text",
                "please revise",
                "--start",
                "0",
                "--end",
                "3",
                "--author",
                "reviewer",
                "--json",
            ]
        )
        assert rc == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload["output"] == str(outfile)
        assert len(payload["added"]) == 1
        added = payload["added"][0]
        assert added["paragraph"] == "p:1"
        assert added["author"] == "reviewer"
        assert added["text"] == "please revise"

        # 验证写出文件里真的能再读到批注
        doc2 = DocxDocument.parse(outfile.read_bytes(), keep_comments=True)
        comments = doc2.comments()
        assert any(c.text == "please revise" for c in comments)

    def test_requires_path_and_text_together(self, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        outfile = tmp_path / "out.docx"
        rc = main(
            [
                "annotate",
                str(infile),
                str(outfile),
                "--path",
                "p:0",
                # missing --text
            ]
        )
        assert rc == 2

    def test_spec_file_batch(self, capsys, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        outfile = tmp_path / "out.docx"
        spec = tmp_path / "ops.json"
        spec.write_text(
            json.dumps(
                [
                    {"path": "p:0", "text": "a", "start": 0, "end": 2, "author": "u1"},
                    {"path": "p:1", "text": "b", "start": 0, "end": 1, "author": "u2"},
                ]
            ),
            encoding="utf-8",
        )

        rc = main(
            [
                "annotate",
                str(infile),
                str(outfile),
                "--spec",
                str(spec),
                "--json",
            ]
        )
        assert rc == 0
        payload = json.loads(capsys.readouterr().out)
        assert len(payload["added"]) == 2
        authors = {c["author"] for c in payload["added"]}
        assert authors == {"u1", "u2"}

        doc2 = DocxDocument.parse(outfile.read_bytes(), keep_comments=True)
        texts = [c.text for c in doc2.comments()]
        assert "a" in texts and "b" in texts

    def test_invalid_path_target(self, capsys, tmp_path, table_doc):
        """annotate 到非段落应报错。"""
        infile = _write(tmp_path, "in.docx", table_doc)
        outfile = tmp_path / "out.docx"
        rc = main(
            [
                "annotate",
                str(infile),
                str(outfile),
                "--path",
                "t:0",
                "--text",
                "x",
            ]
        )
        assert rc == 2


class TestCommentsCommand:
    def test_comments_after_annotate(self, capsys, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        outfile = tmp_path / "out.docx"

        main(
            [
                "annotate",
                str(infile),
                str(outfile),
                "--path",
                "p:1",
                "--text",
                "alpha",
                "--end",
                "2",
                "--author",
                "A",
            ]
        )
        capsys.readouterr()  # drop annotate output

        rc = main(["comments", str(outfile), "--json"])
        assert rc == 0
        items = json.loads(capsys.readouterr().out)
        assert items
        assert items[0]["text"] == "alpha"
        assert items[0]["author"] == "A"
        assert items[0]["paragraph"] == "p:1"


class TestFileErrors:
    def test_missing_file_reports_error(self, capsys, tmp_path):
        rc = main(["list", str(tmp_path / "nope.docx")])
        assert rc == 2
        assert "error" in capsys.readouterr().err.lower()


class TestHelp:
    def test_help_does_not_crash(self, capsys):
        with pytest.raises(SystemExit) as exc:
            main(["--help"])
        assert exc.value.code == 0
        assert "docxnote" in capsys.readouterr().out.lower()
