"""Fitting text into a box: keep the size, shrink one point at a time, cut only at the floor."""

import pytest

from slidegen import fit
from slidegen.fit import ELLIPSIS, fit_text, fits, load_font, wrap
from slidegen.template import Slot


@pytest.fixture(autouse=True)
def estimated_widths(monkeypatch: pytest.MonkeyPatch) -> None:
    """Use the catalog's width estimate, so results are the same with or without the font."""
    monkeypatch.setattr(fit, "load_font", lambda family, bold: None)


def box(width: float = 500, height: float = 200, size: float = 20, kind: str = "list") -> Slot:
    """A body box like the Master's: Arial Narrow, 6 pt before each paragraph."""
    return Slot("body", 18, kind, 0, 0, size, width_pt=width, height_pt=height, gap_pt=6,
                font="Arial Narrow")  # fmt: skip


def test_short_text_keeps_the_template_size() -> None:
    result = fit_text(box(), ["312 luminaires replaced", "No downtime"])
    assert (result.size, result.cut) == (20, False)
    assert result.paragraphs == ["312 luminaires replaced", "No downtime"]


def test_long_text_shrinks_one_point_at_a_time_until_it_fits() -> None:
    bullets = ["Energy use down 33 percent versus the legacy fittings at both sites"] * 6
    slot = box()
    result = fit_text(slot, bullets)
    assert 12 <= result.size < 20 and not result.cut
    assert fits(slot, bullets, result.size)
    assert not fits(slot, bullets, result.size + 1)  # the largest size that fits


def test_text_is_cut_with_an_ellipsis_only_at_the_floor() -> None:
    slot = box(height=60)
    result = fit_text(slot, ["word " * 400])
    assert result.size == 12  # 60% of 20 pt, kept at the 12 pt minimum
    assert result.cut and result.paragraphs[-1].endswith(ELLIPSIS)
    assert fits(slot, result.paragraphs, result.size)


def test_floor_is_never_below_12_pt() -> None:
    result = fit_text(box(height=20, size=14), ["word " * 400])
    assert result.size == 12  # 60% of 14 would be 8.4


def test_a_text_slot_can_hold_line_breaks() -> None:
    result = fit_text(box(kind="text"), "Plant management update\nTunis, 17 November 2026")
    assert result.paragraphs == ["Plant management update", "Tunis, 17 November 2026"]


def test_wrap_breaks_lines_and_splits_overlong_words() -> None:
    slot = box(width=100)
    lines = wrap(slot, "short words " + "x" * 60, 20)
    assert lines[0] == "short words"
    assert len(lines) >= 3 and "".join(lines[1:]) == "x" * 60


def test_uppercase_styles_are_measured_wider() -> None:
    lower, caps = box(), box()
    caps.caps = True
    text = ["A title that states the point of the slide clearly"] * 5
    assert fit_text(caps, text).size <= fit_text(lower, text).size


def test_real_font_is_used_when_installed(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.undo()  # back to the real font loader
    if load_font("Arial Narrow", False) is None:
        pytest.skip("Arial Narrow is not installed here")
    slot = box()
    assert fit.width(slot, "WWWW", 20) > fit.width(slot, "iiii", 20)  # real glyph widths
