"""The command line, always with the mock or without a key: never a real call."""

import io
from pathlib import Path

import pytest

from slidegen import fit
from slidegen.__main__ import main

NOTES = "Phase 1 done on time\n- 312 luminaires replaced\n- no downtime"


@pytest.fixture(autouse=True)
def estimated_widths(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(fit, "load_font", lambda family, bold: None)


def test_mock_writes_a_deck(tmp_path: Path, capsys: pytest.CaptureFixture) -> None:
    out = tmp_path / "deck.pptx"
    assert main([NOTES, "--mock", "--out", str(out)]) == 0
    assert out.exists()
    assert "first try" in capsys.readouterr().out


def test_notes_from_stdin(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("sys.stdin", io.StringIO(NOTES))
    out = tmp_path / "deck.pptx"
    assert main(["--mock", "--out", str(out)]) == 0
    assert out.exists()


@pytest.mark.parametrize("mode", ["repair", "fallback"])
def test_mock_modes_show_repair_and_fallback(
    mode: str, tmp_path: Path, capsys: pytest.CaptureFixture
) -> None:
    out = tmp_path / "deck.pptx"
    assert main([NOTES, "--mock", mode, "--verbose", "--out", str(out)]) == 0
    output = capsys.readouterr().out
    assert ("repaired" if mode == "repair" else "fallback") in output
    assert "call 2:" in output  # two calls in both cases, never more


def test_empty_notes_exit_2(capsys: pytest.CaptureFixture) -> None:
    assert main(["   ", "--mock"]) == 2
    assert "notes are empty" in capsys.readouterr().err


def test_missing_key_exits_3(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture
) -> None:
    monkeypatch.chdir(tmp_path)  # no .env here, and the test fixture removed the variable
    assert main([NOTES, "--out", str(tmp_path / "deck.pptx")]) == 3
    assert "OPENROUTER_API_KEY" in capsys.readouterr().err


def test_unknown_template_file_exits_2(tmp_path: Path, capsys: pytest.CaptureFixture) -> None:
    missing = tmp_path / "nope.pptx"
    assert (
        main([NOTES, "--mock", "--template", str(missing), "--out", str(tmp_path / "d.pptx")]) == 2
    )
    assert "error:" in capsys.readouterr().err


def test_team_template_is_used(tmp_path: Path, capsys: pytest.CaptureFixture) -> None:
    assert main([NOTES, "--mock", "--team", "finance", "--out", str(tmp_path / "d.pptx")]) == 0
    assert "template finance.pptx" in capsys.readouterr().out
