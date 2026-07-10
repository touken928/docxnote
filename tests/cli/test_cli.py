"""docxnote CLI 的端到端测试"""

from __future__ import annotations

import json
import os
import stat
from pathlib import Path

import pytest

import docxnote.cli as cli_module
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

    @pytest.mark.parametrize(
        "bad_op",
        [
            {"path": "p:0", "text": "x", "start": True},
            {"path": "p:0", "text": "x", "start": 1.5},
            {"path": "p:0", "text": "x", "end": []},
            {"path": "p:0", "text": "x", "extra": 1},
            {"path": "", "text": "x"},
            {"path": "p:0", "text": 1},
            {"path": "p:0", "text": "x", "author": False},
        ],
    )
    def test_spec_rejects_invalid_schema(self, capsys, tmp_path, simple_doc, bad_op):
        infile = _write(tmp_path, "in.docx", simple_doc)
        outfile = _write(tmp_path, "out.docx", b"prior output")
        spec = tmp_path / "ops.json"
        spec.write_text(json.dumps([bad_op]), encoding="utf-8")

        assert main(["annotate", str(infile), str(outfile), "--spec", str(spec)]) == 2
        assert outfile.read_bytes() == b"prior output"
        assert "error" in capsys.readouterr().err.lower()

    def test_later_batch_failure_preserves_existing_output(self, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        outfile = _write(tmp_path, "out.docx", b"prior output")
        spec = tmp_path / "ops.json"
        spec.write_text(
            json.dumps(
                [
                    {"path": "p:0", "text": "first"},
                    {"path": "p:999", "text": "later"},
                ]
            ),
            encoding="utf-8",
        )

        assert main(["annotate", str(infile), str(outfile), "--spec", str(spec)]) == 2
        assert outfile.read_bytes() == b"prior output"

    def test_batch_target_validation_precedes_apply(
        self, monkeypatch, tmp_path, simple_doc
    ):
        infile = _write(tmp_path, "in.docx", simple_doc)
        outfile = tmp_path / "out.docx"
        spec = tmp_path / "ops.json"
        spec.write_text(
            json.dumps(
                [
                    {"path": "p:0", "text": "first"},
                    {"path": "p:999", "text": "later"},
                ]
            ),
            encoding="utf-8",
        )
        applied = []
        monkeypatch.setattr(
            cli_module,
            "_apply_annotate_op",
            lambda target, op: applied.append(op),
        )

        assert main(["annotate", str(infile), str(outfile), "--spec", str(spec)]) == 2
        assert applied == []

    @pytest.mark.parametrize(
        "alias_kind", ["direct", "relative", "symlink", "hardlink"]
    )
    def test_input_output_aliases_are_rejected(
        self, capsys, monkeypatch, tmp_path, simple_doc, alias_kind
    ):
        infile = _write(tmp_path, "in.docx", simple_doc)
        input_bytes = infile.read_bytes()
        if alias_kind == "direct":
            input_name = output_name = str(infile)
        elif alias_kind == "relative":
            monkeypatch.chdir(tmp_path)
            input_name, output_name = "in.docx", "./sub/../in.docx"
            (tmp_path / "sub").mkdir()
        elif alias_kind == "symlink":
            output = tmp_path / "out.docx"
            output.symlink_to(infile)
            input_name, output_name = str(infile), str(output)
        else:
            output = tmp_path / "out.docx"
            os.link(infile, output)
            input_name, output_name = str(infile), str(output)

        assert (
            main(["annotate", input_name, output_name, "--path", "p:0", "--text", "x"])
            == 2
        )
        assert "different files" in capsys.readouterr().err
        assert infile.read_bytes() == input_bytes

    def test_spec_explicit_null_end_empty_author_and_omitted_defaults(
        self, capsys, tmp_path, simple_doc
    ):
        infile = _write(tmp_path, "in.docx", simple_doc)
        input_bytes = infile.read_bytes()
        outfile = tmp_path / "out.docx"
        spec = tmp_path / "ops.json"
        spec.write_text(
            json.dumps(
                [
                    {"path": "p:0", "text": "null", "end": None},
                    {"path": "p:1", "text": "empty", "author": ""},
                    {"path": "p:0", "text": "defaults"},
                ]
            ),
            encoding="utf-8",
        )

        assert (
            main(["annotate", str(infile), str(outfile), "--spec", str(spec), "--json"])
            == 0
        )
        added = json.loads(capsys.readouterr().out)["added"]
        assert added[0]["end"] == len("测试文档")
        assert added[0]["author"] == "docxnote"
        assert added[1]["author"] == ""
        assert added[2]["start"] == 0
        assert added[2]["author"] == "docxnote"
        assert infile.read_bytes() == input_bytes

    @pytest.mark.skipif(os.name != "posix", reason="permission bits are POSIX-specific")
    def test_atomic_output_modes(self, tmp_path, simple_doc):
        infile = _write(tmp_path, "in.docx", simple_doc)
        existing = _write(tmp_path, "existing.docx", b"prior output")
        existing.chmod(0o640)
        assert (
            main(
                ["annotate", str(infile), str(existing), "--path", "p:0", "--text", "x"]
            )
            == 0
        )
        assert stat.S_IMODE(existing.stat().st_mode) == 0o640

        output = tmp_path / "new.docx"
        old_umask = os.umask(0o027)
        try:
            assert (
                main(
                    [
                        "annotate",
                        str(infile),
                        str(output),
                        "--path",
                        "p:0",
                        "--text",
                        "x",
                    ]
                )
                == 0
            )
        finally:
            os.umask(old_umask)
        assert stat.S_IMODE(output.stat().st_mode) == 0o640

    def test_windows_atomic_mode_path_does_not_use_fchmod(
        self, monkeypatch, tmp_path, simple_doc
    ):
        outfile = _write(tmp_path, "out.docx", b"prior output")
        outfile.chmod(0o640)
        monkeypatch.setattr(cli_module.os, "name", "nt")

        def fail_fchmod(*args):
            raise AssertionError("Windows must not call os.fchmod")

        monkeypatch.setattr(cli_module.os, "fchmod", fail_fchmod)
        cli_module._write_atomic(outfile, simple_doc)
        assert outfile.read_bytes() == simple_doc

    def test_output_replace_failure_cleans_temp_and_preserves_output(
        self, monkeypatch, tmp_path, simple_doc
    ):
        infile = _write(tmp_path, "in.docx", simple_doc)
        outfile = _write(tmp_path, "out.docx", b"prior output")

        def fail_replace(source, destination):
            raise OSError("replace failed")

        monkeypatch.setattr(cli_module.os, "replace", fail_replace)
        assert (
            main(
                ["annotate", str(infile), str(outfile), "--path", "p:0", "--text", "x"]
            )
            == 2
        )
        assert outfile.read_bytes() == b"prior output"
        assert list(tmp_path.glob(f".{outfile.name}.*.tmp")) == []


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
