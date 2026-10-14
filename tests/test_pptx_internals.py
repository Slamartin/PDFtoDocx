"""Unit tests for the PPTX converter's building blocks: colour maths, inherited style resolution,
bullets and spacing, OOXML helpers, layout arithmetic, reading order and paragraph emission.

The end-to-end deck conversion lives in test_convert.py; these tests pin down the pieces it is
built from so a regression points at the function that broke.
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from types import SimpleNamespace

import pytest
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn as wqn
from docx.shared import Emu, Pt
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches

from app.converters import pptx_to_docx as m
from app.converters.pptx_to_docx import (
    Box,
    Ctx,
    ParaStyle,
    RunStyle,
    Theme,
    _advance,
    _apply_run_style,
    _bullet_from,
    _classify,
    _column_widths,
    _emit_notes,
    _emit_paragraph,
    _emit_picture,
    _estimate_text_height,
    _first_run_rPr,
    _flatten,
    _font_scale,
    _group_rows,
    _insert_in_order,
    _ordered_shapes,
    _para_style,
    _picture_box,
    _plan_slide,
    _pPr_chain,
    _prepare_picture,
    _run_style,
    _set_paragraph_mark_size,
    _setup_page,
    _shade_cell,
    _shade_paragraph,
    _spacing_pts,
    _split_symbol_runs,
    _take_gap,
    _trimmed_paragraphs,
    convert_pptx,
)
from app.report import ConversionReport, CorruptFileError

USABLE_W, USABLE_H = 8229600, 5943600  # a 16:9-ish 10in x 7.5in slide minus the margins


# ---------- helpers ----------


def el(tag: str, *children, **attrs):
    """Builds a DrawingML element: el("a:rPr", el("a:latin", typeface="Arial"), sz="2400")."""
    node = etree.Element(qn(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    for child in children:
        node.append(child)
    return node


def solid(color_el):
    return el("a:solidFill", color_el)


def fake_theme(colors=None, clr_map=None, major="Cambria", minor="Calibri") -> Theme:
    """A Theme without a presentation behind it, so colour and font lookups can be tested alone."""
    theme = Theme.__new__(Theme)
    theme.colors = colors if colors is not None else {"dk1": "111111", "lt1": "FFFFFF", "accent1": "4F81BD"}
    theme.clr_map = clr_map if clr_map is not None else {"tx1": "dk1", "bg1": "lt1", "accent1": "accent1"}
    theme.major_font = major
    theme.minor_font = minor
    theme.tx_styles = {}
    theme.default_text_style = None
    return theme


def make_ctx(theme=None, prs=None, usable_width=USABLE_W, usable_height=USABLE_H) -> Ctx:
    return Ctx(
        doc=Document(),
        report=ConversionReport(),
        usable_width=usable_width,
        usable_height=usable_height,
        prs=prs,
        theme=theme,
    )


@pytest.fixture
def deck():
    """A fresh presentation with one blank slide, its master theme and a matching Ctx."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme = Theme(prs, prs.slide_masters[0])
    return SimpleNamespace(prs=prs, slide=slide, theme=theme, ctx=make_ctx(theme=theme, prs=prs))


def textbox(slide, text="", left=Inches(1), top=Inches(1), width=Inches(4), height=Inches(1)):
    box = slide.shapes.add_textbox(left, top, width, height)
    if text:
        box.text_frame.text = text
    return box


def png_bytes(width=120, height=80, mode="RGB") -> bytes:
    out = BytesIO()
    Image.new(mode, (width, height), (90, 90, 90) if mode == "RGB" else None).save(out, format="PNG")
    return out.getvalue()


def picture_shape(blob: bytes, width: int, height: int, rotation: float = 0.0, **crops):
    crop = {f"crop_{side}": crops.get(side, 0.0) for side in ("left", "top", "right", "bottom")}
    return SimpleNamespace(
        image=SimpleNamespace(blob=blob, ext="png"), width=width, height=height, rotation=rotation, name="pic", **crop
    )


def box(top, left, bottom, right, kind="other", shape=None) -> Box:
    return Box(top=top, left=left, bottom=bottom, right=right, shape=shape, kind=kind)


# ---------- theme colour maths ----------


class TestThemeModify:
    def test_no_modifiers_returns_uppercase_hex(self):
        assert Theme._modify("ff00aa", el("a:srgbClr")) == "FF00AA"

    def test_lumMod_halves_lightness(self):
        assert Theme._modify("FFFFFF", el("a:srgbClr", el("a:lumMod", val="50000"))) == "808080"

    def test_lumMod_then_lumOff_apply_in_order(self):
        mods = el("a:schemeClr", el("a:lumMod", val="75000"), el("a:lumOff", val="25000"))
        assert Theme._modify("000000", mods) == "404040"

    def test_tint_moves_towards_white(self):
        assert Theme._modify("0000FF", el("a:srgbClr", el("a:tint", val="50000"))) == "8080FF"

    def test_shade_moves_towards_black(self):
        assert Theme._modify("FF0000", el("a:srgbClr", el("a:shade", val="50000"))) == "800000"

    def test_lightness_is_clamped(self):
        assert Theme._modify("FFFFFF", el("a:srgbClr", el("a:lumOff", val="50000"))) == "FFFFFF"

    def test_percent_string_modifier_is_applied(self):
        # ST_Percentage may also be written as "50%"; it must not be skipped as unparseable.
        assert Theme._modify("FFFFFF", el("a:srgbClr", el("a:lumMod", val="50%"))) == "808080"

    def test_unparseable_modifier_is_skipped(self):
        assert Theme._modify("FF0000", el("a:srgbClr", el("a:lumMod", val="abc"))) == "FF0000"

    def test_unknown_modifier_is_ignored(self):
        assert Theme._modify("FF0000", el("a:srgbClr", el("a:alpha", val="50000"))) == "FF0000"


class TestThemeBaseColor:
    def test_srgb(self):
        assert Theme._base_color(el("a:srgbClr", val="ABCDEF")) == "ABCDEF"

    def test_system_colour_uses_last_rendered_value(self):
        assert Theme._base_color(el("a:sysClr", val="windowText", lastClr="000000")) == "000000"

    def test_preset_colour_by_name(self):
        assert Theme._base_color(el("a:prstClr", val="red")) == "FF0000"

    def test_unknown_preset_is_none(self):
        assert Theme._base_color(el("a:prstClr", val="chartreuse")) is None

    def test_unsupported_colour_kind_is_none(self):
        assert Theme._base_color(el("a:hslClr", hue="0", sat="0", lum="0")) is None


class TestThemeColorLookup:
    def test_scheme_colour_follows_the_master_colour_map(self):
        theme = fake_theme()
        assert theme.color_el(el("a:schemeClr", val="tx1")) == "111111"  # tx1 -> dk1

    def test_scheme_colour_without_map_entry_uses_the_name_directly(self):
        theme = fake_theme(clr_map={})
        assert theme.color_el(el("a:schemeClr", val="accent1")) == "4F81BD"

    def test_scheme_colour_with_modifier(self):
        theme = fake_theme(colors={"lt1": "FFFFFF"}, clr_map={"bg1": "lt1"})
        assert theme.color_el(el("a:schemeClr", el("a:lumMod", val="50000"), val="bg1")) == "808080"

    def test_unknown_scheme_colour_is_none(self):
        assert fake_theme().color_el(el("a:schemeClr", val="accent9")) is None

    def test_malformed_srgb_is_none(self):
        assert fake_theme().color_el(el("a:srgbClr", val="FFF")) is None

    def test_color_reads_solid_fill_of_parent(self):
        rPr = el("a:rPr", solid(el("a:srgbClr", val="123456")))
        assert fake_theme().color(rPr) == "123456"

    def test_color_without_solid_fill_is_none(self):
        assert fake_theme().color(el("a:rPr", el("a:noFill"))) is None

    def test_color_with_empty_solid_fill_is_none(self):
        assert fake_theme().color(el("a:rPr", el("a:solidFill"))) is None

    def test_text_colour_default_is_tx1(self):
        assert fake_theme().text_color_default() == "111111"


class TestThemeFont:
    def test_major_and_minor_references(self):
        theme = fake_theme(major="Cambria", minor="Calibri")
        assert theme.font("+mj-lt") == "Cambria"
        assert theme.font("+mn-lt") == "Calibri"

    def test_literal_typeface_passes_through(self):
        assert fake_theme().font("Arial") == "Arial"

    def test_missing_typeface_is_none(self):
        assert fake_theme().font(None) is None
        assert fake_theme().font("") is None


def test_theme_reads_default_template(deck):
    theme = deck.theme
    assert theme.colors["dk1"] == "000000" and theme.colors["lt1"] == "FFFFFF"
    assert all(len(v) == 6 for v in theme.colors.values())
    assert theme.clr_map["tx1"] == "dk1"
    assert theme.major_font and theme.minor_font
    assert set(theme.tx_styles) == {"titleStyle", "bodyStyle", "otherStyle"}
    assert all(v is not None for v in theme.tx_styles.values())
    assert theme.default_text_style is not None


# ---------- bullets and spacing ----------


class TestBulletFrom:
    def test_none(self):
        assert _bullet_from(el("a:pPr", el("a:buNone"))) == "plain"

    def test_auto_number(self):
        assert _bullet_from(el("a:pPr", el("a:buAutoNum", type="arabicPeriod"))) == "number"

    def test_char(self):
        assert _bullet_from(el("a:pPr", el("a:buChar", char="•"))) == "bullet"

    def test_picture_bullet(self):
        assert _bullet_from(el("a:pPr", el("a:buBlip"))) == "bullet"

    def test_unspecified_is_none_so_the_chain_continues(self):
        assert _bullet_from(el("a:pPr", algn="ctr")) is None


class TestSpacingPts:
    def test_percent(self):
        assert _spacing_pts(el("a:lnSpc", el("a:spcPct", val="150000"))) == (1.5, None)

    def test_points(self):
        assert _spacing_pts(el("a:spcBef", el("a:spcPts", val="1200"))) == (None, 12.0)

    def test_empty(self):
        assert _spacing_pts(el("a:spcAft")) == (None, None)

    def test_percent_string_form(self):
        # ST_TextSpacingPercentOrPercentString allows "150%" as well as "150000".
        assert _spacing_pts(el("a:lnSpc", el("a:spcPct", val="150%"))) == (1.5, None)

    def test_garbage_value_is_ignored(self):
        assert _spacing_pts(el("a:lnSpc", el("a:spcPct", val="lots"))) == (None, None)

    def test_garbage_points_value_is_ignored(self):
        assert _spacing_pts(el("a:spcBef", el("a:spcPts", val="12pt"))) == (None, None)


class TestParaStyle:
    def test_defaults_when_chain_is_empty(self):
        assert _para_style([]) == ParaStyle()

    def test_most_specific_element_wins_per_property(self):
        chain = [el("a:pPr", algn="r"), el("a:lvl1pPr", algn="ctr", marL="100", indent="-50")]
        ps = _para_style(chain)
        assert (ps.align, ps.mar_left, ps.indent) == ("r", 100, -50)

    def test_spacing_in_points(self):
        ps = _para_style(
            [el("a:pPr", el("a:spcBef", el("a:spcPts", val="600")), el("a:spcAft", el("a:spcPts", val="300")))]
        )
        assert (ps.before, ps.before_pct, ps.after, ps.after_pct) == (6.0, None, 3.0, None)

    def test_spacing_in_percent_of_font_size(self):
        ps = _para_style([el("a:pPr", el("a:spcBef", el("a:spcPct", val="20000")))])
        assert (ps.before, ps.before_pct) == (0.0, 0.2)

    def test_line_spacing(self):
        assert _para_style([el("a:pPr", el("a:lnSpc", el("a:spcPct", val="90000")))]).line_pct == 0.9
        assert _para_style([el("a:pPr", el("a:lnSpc", el("a:spcPts", val="1800")))]).line_pts == 18.0

    def test_explicit_no_bullet_shadows_inherited_bullet(self):
        chain = [el("a:pPr", el("a:buNone")), el("a:lvl1pPr", el("a:buChar", char="•"))]
        assert _para_style(chain).bullet == "plain"

    def test_inherited_bullet_applies_when_paragraph_is_silent(self):
        chain = [el("a:pPr"), el("a:lvl1pPr", el("a:buChar", char="•"))]
        assert _para_style(chain).bullet == "bullet"

    def test_numbering(self):
        assert _para_style([el("a:pPr", el("a:buAutoNum", type="arabicPeriod"))]).bullet == "number"


# ---------- run style ----------


class TestRunStyle:
    def test_explicit_run_properties(self):
        rPr = el("a:rPr", el("a:latin", typeface="Arial"), sz="2400", b="1", i="0", u="sng")
        rs = _run_style(rPr, [], make_ctx())
        assert rs == RunStyle(size=24.0, bold=True, italic=False, underline=True, color=None, font="Arial")

    def test_run_beats_paragraph_defaults(self):
        rPr = el("a:rPr", sz="2400")
        chain = [el("a:pPr", el("a:defRPr", sz="1800", b="true"))]
        rs = _run_style(rPr, chain, make_ctx())
        assert (rs.size, rs.bold) == (24.0, True)

    def test_defaults_used_when_run_has_none(self):
        chain = [el("a:pPr", el("a:defRPr", sz="1800", i="1"))]
        rs = _run_style(None, chain, make_ctx())
        assert (rs.size, rs.italic) == (18.0, True)

    def test_font_scale_shrinks_size(self):
        assert _run_style(el("a:rPr", sz="2000"), [], make_ctx(), font_scale=0.5).size == 10.0

    def test_underline_none_is_false(self):
        assert _run_style(el("a:rPr", u="none"), [], make_ctx()).underline is False

    def test_without_theme_colour_and_font_stay_unset(self):
        rs = _run_style(el("a:rPr"), [], make_ctx(theme=None))
        assert rs.color is None and rs.font is None

    def test_theme_supplies_default_colour_and_font(self):
        rs = _run_style(el("a:rPr"), [], make_ctx(theme=fake_theme()))
        assert (rs.color, rs.font) == ("111111", "Calibri")

    def test_theme_font_reference_is_resolved(self):
        rPr = el("a:rPr", el("a:latin", typeface="+mj-lt"))
        assert _run_style(rPr, [], make_ctx(theme=fake_theme())).font == "Cambria"

    def test_explicit_fill_is_used(self):
        rPr = el("a:rPr", solid(el("a:srgbClr", val="FF0000")))
        assert _run_style(rPr, [], make_ctx(theme=fake_theme())).color == "FF0000"

    def test_scheme_fill_in_defaults(self):
        chain = [el("a:pPr", el("a:defRPr", solid(el("a:schemeClr", val="accent1"))))]
        assert _run_style(None, chain, make_ctx(theme=fake_theme())).color == "4F81BD"

    def test_unresolvable_explicit_fill_stops_the_chain(self):
        rPr = el("a:rPr", el("a:noFill"))
        chain = [el("a:pPr", el("a:defRPr", solid(el("a:srgbClr", val="00FF00"))))]
        assert _run_style(rPr, chain, make_ctx(theme=fake_theme())).color == "111111"

    def test_shape_style_colour_beats_theme_default(self):
        ctx = make_ctx(theme=fake_theme())
        ctx.font_ref_color = "ABCDEF"
        assert _run_style(el("a:rPr"), [], ctx).color == "ABCDEF"

    def test_explicit_fill_beats_shape_style_colour(self):
        ctx = make_ctx(theme=fake_theme())
        ctx.font_ref_color = "ABCDEF"
        assert _run_style(el("a:rPr", solid(el("a:srgbClr", val="FF0000"))), [], ctx).color == "FF0000"


class TestApplyRunStyle:
    def test_sets_word_run_properties(self):
        run = Document().add_paragraph().add_run("x")
        _apply_run_style(run, RunStyle(size=11.5, bold=True, italic=None, underline=True, color="1F5C8A", font="Arial"))
        assert run.bold is True and run.italic is False and run.underline is True
        assert run.font.size == Pt(11.5) and run.font.name == "Arial" and str(run.font.color.rgb) == "1F5C8A"

    def test_tiny_size_is_floored_at_one_point(self):
        run = Document().add_paragraph().add_run("x")
        _apply_run_style(run, RunStyle(size=0.2))
        assert run.font.size == Pt(1)

    def test_unset_fields_leave_defaults(self):
        run = Document().add_paragraph().add_run("x")
        _apply_run_style(run, RunStyle(color=""))
        assert run.font.size is None and run.font.name is None and run.font.color.rgb is None
        assert run.underline is None


class TestFontScale:
    def test_no_autofit_is_one(self, deck):
        assert _font_scale(textbox(deck.slide, "x")) == 1.0

    def test_norm_autofit_scale(self, deck):
        shape = textbox(deck.slide, "x")
        body_pr = shape.text_frame._txBody.find(qn("a:bodyPr"))
        body_pr.append(el("a:normAutofit", fontScale="75000"))
        assert _font_scale(shape) == 0.75

    def test_norm_autofit_percent_string(self, deck):
        shape = textbox(deck.slide, "x")
        shape.text_frame._txBody.find(qn("a:bodyPr")).append(el("a:normAutofit", fontScale="75%"))
        assert _font_scale(shape) == 0.75


# ---------- symbol-font runs ----------


class TestSplitSymbolRuns:
    def test_no_symbol_font_is_one_piece(self):
        assert list(_split_symbol_runs("ab", None)) == [("ab", None)]

    def test_no_symbol_chars_is_one_piece(self):
        assert list(_split_symbol_runs("plain", "Wingdings")) == [("plain", None)]

    def test_symbol_chars_get_the_font_and_keep_a_trailing_space(self):
        assert list(_split_symbol_runs("a b", "Wingdings")) == [("a", None), (" ", "Wingdings"), ("b", None)]

    def test_leading_symbol(self):
        assert list(_split_symbol_runs("x", "Wingdings")) == [("", "Wingdings"), ("x", None)]

    def test_empty_text(self):
        assert list(_split_symbol_runs("", "Wingdings")) == [("", None)]


# ---------- OOXML helpers ----------


class TestInsertInOrder:
    def test_inserts_before_first_successor(self):
        parent = el("a:pPr", el("a:buNone"), el("a:defRPr"))
        _insert_in_order(parent, el("a:spcBef"), ("a:defRPr", "a:extLst"))
        assert [etree.QName(c).localname for c in parent] == ["buNone", "spcBef", "defRPr"]

    def test_appends_when_no_successor_present(self):
        parent = el("a:pPr", el("a:buNone"))
        _insert_in_order(parent, el("a:spcBef"), ("a:defRPr",))
        assert [etree.QName(c).localname for c in parent] == ["buNone", "spcBef"]


class TestShading:
    def test_paragraph_shading_precedes_justification(self):
        p = Document().add_paragraph("x")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shade_paragraph(p, "1F5C8A")
        tags = [etree.QName(c).localname for c in p._p.pPr]
        assert tags.index("shd") < tags.index("jc")
        assert p._p.pPr.find(wqn("w:shd")).get(wqn("w:fill")) == "1F5C8A"

    def test_paragraph_shading_is_replaced_not_duplicated(self):
        p = Document().add_paragraph("x")
        _shade_paragraph(p, "111111")
        _shade_paragraph(p, "222222")
        shds = p._p.pPr.findall(wqn("w:shd"))
        assert len(shds) == 1 and shds[0].get(wqn("w:fill")) == "222222"

    def test_cell_shading_precedes_vertical_alignment(self):
        cell = Document().add_table(1, 1).cell(0, 0)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _shade_cell(cell, "ABCDEF")
        tags = [etree.QName(c).localname for c in cell._tc.tcPr]
        assert tags.index("shd") < tags.index("vAlign")
        assert cell._tc.tcPr.find(wqn("w:shd")).get(wqn("w:fill")) == "ABCDEF"


class TestParagraphMarkSize:
    def test_sets_half_points_on_both_size_elements(self):
        p = Document().add_paragraph("x")
        _set_paragraph_mark_size(p, 12.5)
        rPr = p._p.pPr.find(wqn("w:rPr"))
        assert rPr.find(wqn("w:sz")).get(wqn("w:val")) == "25"
        assert rPr.find(wqn("w:szCs")).get(wqn("w:val")) == "25"

    def test_second_call_updates_in_place(self):
        p = Document().add_paragraph("x")
        _set_paragraph_mark_size(p, 12)
        _set_paragraph_mark_size(p, 8)
        rPr = p._p.pPr.find(wqn("w:rPr"))
        assert len(rPr.findall(wqn("w:sz"))) == 1 and rPr.find(wqn("w:sz")).get(wqn("w:val")) == "16"


# ---------- page setup and layout arithmetic ----------


class TestSetupPage:
    def test_landscape_slide_keeps_its_size(self):
        doc, report = Document(), ConversionReport()
        uw, uh = _setup_page(doc, SimpleNamespace(slide_width=9144000, slide_height=6858000), report)
        sec = doc.sections[0]
        assert sec.page_width == Emu(9144000) and sec.page_height == Emu(6858000)
        assert sec.orientation == m.WD_ORIENT.LANDSCAPE
        assert sec.left_margin == sec.right_margin == sec.top_margin == sec.bottom_margin == m.MARGIN
        assert (uw, uh) == (9144000 - 2 * m.MARGIN, 6858000 - 2 * m.MARGIN)
        assert report.warnings == []

    def test_portrait_slide(self):
        doc = Document()
        _setup_page(doc, SimpleNamespace(slide_width=6858000, slide_height=9144000), ConversionReport())
        assert doc.sections[0].orientation == m.WD_ORIENT.PORTRAIT

    def test_tiny_slide_is_clamped_up_with_a_warning(self):
        doc, report = Document(), ConversionReport()
        _setup_page(doc, SimpleNamespace(slide_width=100, slide_height=100), report)
        assert doc.sections[0].page_width == Emu(m.MIN_PAGE) and doc.sections[0].page_height == Emu(m.MIN_PAGE)
        assert any("clamped" in w for w in report.warnings)

    def test_huge_slide_is_clamped_down(self):
        doc, report = Document(), ConversionReport()
        _setup_page(doc, SimpleNamespace(slide_width=40 * m.EMU_PER_INCH, slide_height=30 * m.EMU_PER_INCH), report)
        assert doc.sections[0].page_width == Emu(m.MAX_PAGE) and doc.sections[0].page_height == Emu(m.MAX_PAGE)
        assert any("clamped" in w for w in report.warnings)


class TestGapAndCursor:
    def test_gap_is_distance_from_cursor(self):
        ctx = make_ctx(usable_height=1000)
        ctx.cursor, ctx.gap_budget = 100, 1000
        assert _take_gap(ctx, 150) == 50 and ctx.gap_budget == 950

    def test_gap_is_capped_by_max_gap_share(self):
        ctx = make_ctx(usable_height=1000)
        ctx.cursor, ctx.gap_budget = 0, 1000
        assert _take_gap(ctx, 900) == int(1000 * m.MAX_GAP)

    def test_gap_is_capped_by_budget_and_spends_it(self):
        ctx = make_ctx(usable_height=1000)
        ctx.cursor, ctx.gap_budget = 0, 30
        assert _take_gap(ctx, 150) == 30
        assert ctx.gap_budget == 0
        assert _take_gap(ctx, 400) == 0

    def test_overlapping_shape_has_no_gap(self):
        ctx = make_ctx(usable_height=1000)
        ctx.cursor, ctx.gap_budget = 500, 1000
        assert _take_gap(ctx, 200) == 0 and ctx.gap_budget == 1000

    def test_advance_never_moves_backwards(self):
        ctx = make_ctx()
        _advance(ctx, 100, 50)
        assert ctx.cursor == 150
        _advance(ctx, 10, 20)
        assert ctx.cursor == 150
        _advance(ctx, 200, None)
        assert ctx.cursor == 200


class TestEstimateTextHeight:
    def test_flattened_url_counts_towards_the_height(self, deck):
        shape = textbox(deck.slide, width=Inches(3))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = "docs"
        plain = _estimate_text_height(shape.text_frame, shape, deck.ctx, Inches(3))
        run.hyperlink.address = "https://example.com/" + "a" * 120
        assert _estimate_text_height(shape.text_frame, shape, deck.ctx, Inches(3)) > plain


class TestPictureBox:
    def test_small_picture_keeps_its_size(self):
        ctx = make_ctx(usable_width=1000, usable_height=1000)
        assert _picture_box(SimpleNamespace(width=400, height=200, rotation=0), ctx) == (400, 200)

    def test_width_cap_scales_proportionally(self):
        ctx = make_ctx(usable_width=1000, usable_height=1000)
        assert _picture_box(SimpleNamespace(width=400, height=200, rotation=0), ctx, max_width=200) == (200, 100)

    def test_usable_width_is_the_default_cap(self):
        ctx = make_ctx(usable_width=1000, usable_height=1000)
        assert _picture_box(SimpleNamespace(width=4000, height=200, rotation=0), ctx) == (1000, 50)

    def test_rotated_picture_swaps_extents(self):
        ctx = make_ctx(usable_width=1000, usable_height=1000)
        assert _picture_box(SimpleNamespace(width=400, height=200, rotation=90.0), ctx) == (200, 400)
        assert _picture_box(SimpleNamespace(width=400, height=200, rotation=270.0), ctx) == (200, 400)

    def test_tall_picture_is_limited_by_page_height(self):
        ctx = make_ctx(usable_width=1000, usable_height=1000)
        w, h = _picture_box(SimpleNamespace(width=100, height=2000, rotation=0), ctx)
        assert h <= int(1000 * m.MAX_PICTURE_HEIGHT) and abs(w / h - 100 / 2000) < 0.01

    def test_zero_size_falls_back_to_half_page(self):
        ctx = make_ctx(usable_width=1000, usable_height=1000)
        assert _picture_box(SimpleNamespace(width=0, height=None, rotation=None), ctx) == (500, 500)


class TestGroupRows:
    def test_side_by_side_shapes_share_a_row_sorted_by_left(self):
        right, left = box(0, 200, 100, 300), box(0, 0, 100, 100)
        rows = _group_rows([right, left])
        assert rows == [[left, right]]

    def test_stacked_shapes_get_their_own_rows(self):
        upper, lower = box(0, 0, 100, 100), box(200, 0, 300, 100)
        assert _group_rows([lower, upper]) == [[upper], [lower]]

    def test_horizontally_overlapping_shapes_are_not_columns(self):
        a, b = box(0, 0, 100, 100), box(10, 50, 110, 150)
        assert _group_rows([a, b]) == [[a], [b]]

    def test_slight_vertical_overlap_is_not_a_row(self):
        a, b = box(0, 0, 100, 100), box(90, 200, 300, 300)  # 10 of 100 overlap: 10 %, below the 30 % threshold
        assert _group_rows([a, b]) == [[a], [b]]

    def test_at_most_four_columns(self):
        boxes = [box(0, i * 100, 50, i * 100 + 80) for i in range(5)]
        rows = _group_rows(boxes)
        assert [len(r) for r in rows] == [4, 1]

    def test_rows_are_ordered_top_to_bottom(self):
        a, b, c = box(500, 0, 600, 100), box(0, 0, 100, 100), box(250, 0, 350, 100)
        assert _group_rows([a, b, c]) == [[b], [c], [a]]


class TestColumnWidths:
    def test_single_box_takes_full_width(self):
        assert _column_widths([box(0, 0, 10, 10)], make_ctx(usable_width=1000)) == [1000]

    def test_widths_follow_horizontal_spans_and_fill_the_page(self):
        row = [box(0, 0, 100, 100, kind="picture"), box(0, 500, 100, 600, kind="picture")]
        widths = _column_widths(row, make_ctx(usable_width=1000))
        assert widths[0] > widths[1]
        assert abs(sum(widths) - 1000) <= 2

    def test_narrow_column_gets_a_minimum_share(self):
        row = [box(0, 0, 100, 900, kind="picture"), box(0, 990, 100, 1000, kind="picture")]
        widths = _column_widths(row, make_ctx(usable_width=1000))
        assert widths[1] >= 100


class TestPlanSlide:
    def test_content_that_fits_keeps_full_size_and_leaves_gap_budget(self):
        ctx = make_ctx()
        pic = box(
            0, 0, 1000000, 1000000, kind="picture", shape=SimpleNamespace(width=1000000, height=1000000, rotation=0)
        )
        _plan_slide([], [[pic]], ctx)
        assert (pic.pic_w, pic.pic_h, pic.height) == (1000000, 1000000, 1000000)
        assert pic.col_width == USABLE_W
        assert ctx.gap_budget == int(USABLE_H * m.PAGE_FILL) - m.HEADING_FALLBACK_HEIGHT - 1000000

    def test_oversized_picture_is_scaled_to_fit_the_page(self):
        ctx = make_ctx()
        pic = box(
            0, 0, 20000000, 3000000, kind="picture", shape=SimpleNamespace(width=3000000, height=20000000, rotation=0)
        )
        _plan_slide([], [[pic]], ctx)
        available = int(USABLE_H * m.PAGE_FILL)
        assert pic.pic_h < int(USABLE_H * m.MAX_PICTURE_HEIGHT)
        assert m.HEADING_FALLBACK_HEIGHT + pic.pic_h <= available
        assert 0 <= ctx.gap_budget < 0.01 * available
        assert abs(pic.pic_w / pic.pic_h - 3 / 20) < 0.01

    def test_other_kinds_have_fixed_or_bounded_heights(self):
        ctx = make_ctx()
        table = box(0, 0, 50000000, 1000, kind="table")
        other = box(0, 0, 100, 100, kind="other")
        _plan_slide([], [[table], [other]], ctx)
        assert table.height == USABLE_H // 2
        assert other.height == 20 * m.EMU_PER_PT


# ---------- reading order, classification and text-frame helpers ----------


class TestFlattenAndOrder:
    def test_body_is_sorted_by_row_then_left_and_title_is_separate(self, deck):
        slide = deck.prs.slides.add_slide(deck.prs.slide_layouts[5])  # title only
        slide.shapes.title.text = "T"
        lower = textbox(slide, "lower", left=Inches(1), top=Inches(2))
        right = textbox(slide, "right", left=Inches(3), top=Inches(1))
        left = textbox(slide, "left", left=Inches(0.5), top=Inches(1.05))  # same half-centimetre row as `right`
        titles, body = _ordered_shapes(slide)
        assert [t[2] for t in titles] == [slide.shapes.title]
        assert [t[2] for t in body] == [left, right, lower]

    def test_missing_offsets_count_as_zero(self):
        acc: list = []
        _flatten([SimpleNamespace(top=None, left=None, shape_type=None)], acc, 0, 0)
        assert acc[0][:2] == (0, 0)

    def test_group_child_keeps_its_slide_position(self, deck):
        # PowerPoint (and python-pptx) store group children in a child space whose origin, chOff,
        # usually equals the group's own offset, so the child's slide position is its own top/left.
        group = deck.slide.shapes.add_group_shape()
        child = group.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
        child.text_frame.text = "in group"
        assert (group.top, group.left) == (Inches(2), Inches(1))
        acc: list = []
        _flatten(deck.slide.shapes, acc, 0, 0)
        assert acc == [(Inches(2), Inches(1), child)]

    def test_moved_group_shifts_its_children(self, deck):
        group = deck.slide.shapes.add_group_shape()
        child = group.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
        group.left, group.top = Inches(4), Inches(5)  # move the group: chOff stays, off changes
        acc: list = []
        _flatten(deck.slide.shapes, acc, 0, 0)
        assert acc == [(Inches(5), Inches(4), child)]

    def test_resized_group_scales_child_offsets(self, deck):
        group = deck.slide.shapes.add_group_shape()
        group.shapes.add_textbox(Inches(1), Inches(1), Inches(1), Inches(1))
        child = group.shapes.add_textbox(Inches(2), Inches(1), Inches(1), Inches(1))
        assert (group.width, group.height) == (Inches(2), Inches(1))
        group.width = Inches(4)  # doubled: the second child now starts 2in right of the group's left edge
        acc: list = []
        _flatten(deck.slide.shapes, acc, 0, 0)
        assert [(top, left) for top, left, s in acc if s == child] == [(Inches(1), Inches(3))]

    def test_nested_group_combines_offsets_and_scales(self, deck):
        # Both groups start with chOff = off; moving and widening the outer group must carry the inner
        # group's children along: x = outer.off + (inner.x - outer.chOff) * 2 + (x - inner.chOff) * 2.
        outer = deck.slide.shapes.add_group_shape()
        inner = outer.shapes.add_group_shape()
        inner.shapes.add_textbox(Inches(1), Inches(1), Inches(1), Inches(1))
        child = inner.shapes.add_textbox(Inches(2), Inches(1), Inches(1), Inches(1))
        assert (outer.left, outer.width) == (Inches(1), Inches(2))
        outer.left, outer.width = Inches(4), Inches(4)
        acc: list = []
        _flatten(deck.slide.shapes, acc, 0, 0)
        assert [(top, left) for top, left, s in acc if s == child] == [(Inches(1), Inches(6))]


class TestClassify:
    def test_text_box_with_text(self, deck):
        assert _classify(textbox(deck.slide, "hello")) == "text"

    def test_empty_text_box_is_skipped(self, deck):
        assert _classify(textbox(deck.slide, "   ")) is None

    def test_picture(self, deck):
        pic = deck.slide.shapes.add_picture(BytesIO(png_bytes()), Inches(1), Inches(1))
        assert _classify(pic) == "picture"

    def test_table(self, deck):
        frame = deck.slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(1))
        assert _classify(frame) == "table"

    def test_chart_is_other(self):
        shape = SimpleNamespace(is_placeholder=False, has_table=False, has_chart=True)
        assert _classify(shape) == "other"

    def test_footer_placeholder_is_skipped(self):
        shape = SimpleNamespace(is_placeholder=True, placeholder_format=SimpleNamespace(type=PP_PLACEHOLDER.FOOTER))
        assert _classify(shape) is None


class TestTextFrameHelpers:
    def test_trimmed_paragraphs_drop_blank_edges_only(self, deck):
        tf = textbox(deck.slide).text_frame
        tf.text = ""
        for text in ("a", "", "b", ""):
            tf.add_paragraph().text = text
        assert [p.text for p in _trimmed_paragraphs(tf)] == ["a", "", "b"]

    def test_line_break_only_paragraph_is_not_empty(self, deck):
        tf = textbox(deck.slide).text_frame
        tf.paragraphs[0].add_line_break()
        assert len(_trimmed_paragraphs(tf)) == 1

    def test_first_run_rPr(self, deck):
        p = textbox(deck.slide).text_frame.paragraphs[0]
        run = p.add_run()
        run.text = "x"
        run.font.size = Pt(9)
        assert _first_run_rPr(p) is run._r.find(qn("a:rPr"))

    def test_first_run_rPr_falls_back_to_paragraph_end_properties(self, deck):
        p = textbox(deck.slide).text_frame.paragraphs[0]
        end = el("a:endParaRPr", sz="900")
        p._p.append(end)
        assert _first_run_rPr(p) is end

    def test_first_run_rPr_is_none_for_a_bare_paragraph(self, deck):
        assert _first_run_rPr(textbox(deck.slide).text_frame.paragraphs[0]) is None


class TestPPrChain:
    def test_body_placeholder_inherits_from_body_style(self, deck):
        slide = deck.prs.slides.add_slide(deck.prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "x"
        chain = _pPr_chain(body.text_frame.paragraphs[0], body, deck.ctx)
        assert deck.theme.tx_styles["bodyStyle"].find(qn("a:lvl1pPr")) in chain
        assert chain[-1] is deck.theme.default_text_style.find(qn("a:lvl1pPr"))

    def test_title_placeholder_inherits_from_title_style(self, deck):
        slide = deck.prs.slides.add_slide(deck.prs.slide_layouts[5])
        slide.shapes.title.text = "x"
        chain = _pPr_chain(slide.shapes.title.text_frame.paragraphs[0], slide.shapes.title, deck.ctx)
        assert deck.theme.tx_styles["titleStyle"].find(qn("a:lvl1pPr")) in chain
        assert deck.theme.tx_styles["bodyStyle"].find(qn("a:lvl1pPr")) not in chain

    def test_text_box_inherits_from_other_style(self, deck):
        shape = textbox(deck.slide, "x")
        chain = _pPr_chain(shape.text_frame.paragraphs[0], shape, deck.ctx)
        assert deck.theme.tx_styles["otherStyle"].find(qn("a:lvl1pPr")) in chain

    def test_paragraph_level_selects_the_matching_level(self, deck):
        slide = deck.prs.slides.add_slide(deck.prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "x"
        p = body.text_frame.paragraphs[0]
        p.level = 1
        chain = _pPr_chain(p, body, deck.ctx)
        assert deck.theme.tx_styles["bodyStyle"].find(qn("a:lvl2pPr")) in chain

    def test_paragraph_own_properties_come_first(self, deck):
        shape = textbox(deck.slide, "x")
        p = shape.text_frame.paragraphs[0]
        pPr = p._p.get_or_add_pPr()
        assert _pPr_chain(p, shape, deck.ctx)[0] is pPr


# ---------- paragraph emission ----------


class TestEmitParagraph:
    def test_hyperlink_is_flattened_with_the_url(self, deck):
        shape = textbox(deck.slide)
        p = shape.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = "docs"
        run.hyperlink.address = "https://example.com/x"
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert out.text == "docs (https://example.com/x)"
        assert any("Hyperlinks" in n for n in deck.ctx.report.notes)

    def test_hyperlink_split_over_runs_gets_the_url_once(self, deck):
        # PowerPoint often splits one link into several runs; the URL follows the whole link, not every piece.
        shape = textbox(deck.slide)
        p = shape.text_frame.paragraphs[0]
        p.add_run().text = "see "
        for piece in ("do", "cs"):
            run = p.add_run()
            run.text = piece
            run.hyperlink.address = "https://example.com/x"
        p.add_run().text = " now"
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert out.text == "see docs (https://example.com/x) now"

    def test_hyperlink_showing_its_own_url_is_not_repeated(self, deck):
        shape = textbox(deck.slide)
        p = shape.text_frame.paragraphs[0]
        for piece in ("www.example.com/", "x/"):
            run = p.add_run()
            run.text = piece
            run.hyperlink.address = "https://www.example.com/x"
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert out.text == "www.example.com/x/"
        assert any("Hyperlinks" in n for n in deck.ctx.report.notes)

    def test_adjacent_links_keep_their_own_urls(self, deck):
        shape = textbox(deck.slide)
        p = shape.text_frame.paragraphs[0]
        for piece, url in (("a", "https://a.example"), ("b", "https://b.example")):
            run = p.add_run()
            run.text = piece
            run.hyperlink.address = url
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert out.text == "a (https://a.example)b (https://b.example)"

    def test_line_break_becomes_a_word_break(self, deck):
        shape = textbox(deck.slide)
        p = shape.text_frame.paragraphs[0]
        p.add_run().text = "one"
        p.add_line_break()
        p.add_run().text = "two"
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert out.text == "one\ntwo"
        assert len(out._p.findall(".//" + wqn("w:br"))) == 1

    def test_field_text_is_emitted(self, deck):
        shape = textbox(deck.slide)
        p = shape.text_frame.paragraphs[0]
        # python-pptx's parser gives a:fld a .text property; a plain lxml element would not have one.
        fld = parse_xml(
            f'<a:fld {nsdecls("a")} id="{{00000000-0000-0000-0000-000000000001}}" type="slidenum"><a:t>7</a:t></a:fld>'
        )
        p._p.append(fld)
        assert _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx).text == "7"

    def test_bullet_level_is_clamped_to_word_list_styles(self, deck):
        slide = deck.prs.slides.add_slide(deck.prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "deep"
        p = body.text_frame.paragraphs[0]
        p.level = 4
        assert _emit_paragraph(deck.ctx.doc, p, body, deck.ctx).style.name == "List Bullet 3"

    def test_forced_style_overrides_bullets(self, deck):
        slide = deck.prs.slides.add_slide(deck.prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "x"
        out = _emit_paragraph(deck.ctx.doc, body.text_frame.paragraphs[0], body, deck.ctx, force_style="Heading 1")
        assert out.style.name == "Heading 1"

    def test_plain_paragraph_keeps_its_indents(self, deck):
        shape = textbox(deck.slide, "x")
        p = shape.text_frame.paragraphs[0]
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", "457200")
        pPr.set("indent", "-228600")
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert out.style.name == "Normal"
        assert out.paragraph_format.left_indent == Emu(457200)
        assert out.paragraph_format.first_line_indent == Emu(-228600)

    def test_bulleted_paragraph_leaves_indents_to_the_list_style(self, deck):
        shape = textbox(deck.slide, "x")
        p = shape.text_frame.paragraphs[0]
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", "457200")
        pPr.append(el("a:buChar", char="•"))
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert out.style.name == "List Bullet"
        assert out.paragraph_format.left_indent is None

    def test_alignment_and_spacing(self, deck):
        shape = textbox(deck.slide, "x")
        p = shape.text_frame.paragraphs[0]
        pPr = p._p.get_or_add_pPr()
        pPr.set("algn", "ctr")
        pPr.append(el("a:lnSpc", el("a:spcPct", val="150000")))
        pPr.append(el("a:spcBef", el("a:spcPts", val="600")))
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx, extra_before=Emu(12700))
        pf = out.paragraph_format
        assert out.alignment == WD_ALIGN_PARAGRAPH.CENTER
        assert pf.line_spacing == 1.5
        assert pf.space_before == Pt(6) + Emu(12700)

    def test_empty_paragraph_keeps_the_end_mark_size(self, deck):
        shape = textbox(deck.slide)
        p = shape.text_frame.paragraphs[0]
        p._p.append(el("a:endParaRPr", sz="900"))
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert out.text == "" and out.runs[0].font.size == Pt(9)

    def test_symbol_font_run_is_split(self, deck):
        shape = textbox(deck.slide)
        p = shape.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = " done"
        run._r.get_or_add_rPr().append(el("a:sym", typeface="Wingdings"))
        out = _emit_paragraph(deck.ctx.doc, p, shape, deck.ctx)
        assert [(r.text, r.font.name) for r in out.runs] == [(" ", "Wingdings"), ("done", deck.theme.minor_font)]


class TestEmitNotes:
    def test_slide_without_notes_adds_nothing(self, deck):
        _emit_notes(deck.slide, deck.ctx)
        assert deck.ctx.doc.paragraphs == []

    def test_notes_become_italic_lines_under_a_heading(self, deck):
        deck.slide.notes_slide.notes_text_frame.text = "first\n\n  second  "
        _emit_notes(deck.slide, deck.ctx)
        paras = deck.ctx.doc.paragraphs
        assert [(p.style.name, p.text) for p in paras] == [
            ("Heading 3", "Speaker notes"),
            ("Normal", "first"),
            ("Normal", "second"),
        ]
        assert all(r.italic for p in paras[1:] for r in p.runs)


# ---------- pictures ----------


class TestPreparePicture:
    def test_untouched_picture_passes_through_unchanged(self):
        blob = png_bytes()
        assert _prepare_picture(picture_shape(blob, 400, 300)) == (blob, 400, 300)

    def test_quarter_turn_swaps_the_box_and_rotates_the_pixels(self):
        data, w, h = _prepare_picture(picture_shape(png_bytes(120, 80), 400, 300, rotation=90.0))
        assert (w, h) == (300, 400)
        assert Image.open(BytesIO(data)).size == (80, 120)

    def test_crop_is_baked_into_the_pixels(self):
        data, w, h = _prepare_picture(picture_shape(png_bytes(120, 80), 400, 300, left=0.5, bottom=0.25))
        assert (w, h) == (400, 300)
        assert Image.open(BytesIO(data)).size == (60, 60)

    def test_arbitrary_angle_grows_the_box_height(self):
        # A 2:1 image turned by 45 degrees has a square bounding box: the width stays, the height grows to match.
        data, w, h = _prepare_picture(picture_shape(png_bytes(100, 50), 400, 200, rotation=45.0))
        assert w == 400 and h > 200
        assert Image.open(BytesIO(data)).width > 100

    def test_exif_orientation_is_applied(self):
        exif = Image.Exif()
        exif[0x0112] = 6  # rotate 90 CW on display
        out = BytesIO()
        Image.new("RGB", (120, 80), (1, 2, 3)).save(out, format="JPEG", exif=exif.tobytes())
        data, w, h = _prepare_picture(picture_shape(out.getvalue(), 400, 300))
        assert (w, h) == (400, 300)
        assert Image.open(BytesIO(data)).size == (80, 120)


class TestEmitPicture:
    def test_vector_picture_is_replaced_by_a_note(self):
        ctx = make_ctx()
        shape = SimpleNamespace(image=SimpleNamespace(ext="emf", blob=b""), width=100, height=100, rotation=0, name="v")
        _emit_picture(box(0, 0, 100, 100, kind="picture", shape=shape), ctx.doc, ctx)
        assert ctx.doc.paragraphs[-1].text == "[Image skipped: EMF pictures are not supported]"
        assert any("EMF" in w for w in ctx.report.warnings) and ctx.report.images == 0

    def test_svg_without_png_fallback_is_replaced_by_a_note(self, deck):
        # Office can embed an SVG with no PNG fallback; python-pptx then raises instead of returning an image.
        pic = deck.slide.shapes.add_picture(BytesIO(png_bytes()), Inches(1), Inches(1))
        blip = pic._element.find(".//" + qn("a:blip"))
        rId = blip.get(qn("r:embed"))
        del blip.attrib[qn("r:embed")]
        blip.append(
            parse_xml(
                f'<a:extLst {nsdecls("a", "r")}><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">'
                f'<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{rId}"/>'
                "</a:ext></a:extLst>"
            )
        )
        _emit_picture(box(0, 0, 100, 100, kind="picture", shape=pic), deck.ctx.doc, deck.ctx)
        assert deck.ctx.doc.paragraphs[-1].text == "[Image skipped: SVG pictures are not supported]"
        assert any("SVG" in w for w in deck.ctx.report.warnings) and deck.ctx.report.images == 0

    def test_linked_picture_is_replaced_by_a_note(self, deck):
        pic = deck.slide.shapes.add_picture(BytesIO(png_bytes()), Inches(1), Inches(1))
        blip = pic._element.find(".//" + qn("a:blip"))
        blip.set(qn("r:link"), blip.get(qn("r:embed")))
        del blip.attrib[qn("r:embed")]
        _emit_picture(box(0, 0, 100, 100, kind="picture", shape=pic), deck.ctx.doc, deck.ctx)
        assert deck.ctx.doc.paragraphs[-1].text == "[Image skipped: linked pictures are not supported]"
        assert any("linked" in w for w in deck.ctx.report.warnings) and deck.ctx.report.images == 0

    def test_raster_picture_is_inserted_at_the_planned_width(self):
        ctx = make_ctx()
        shape = picture_shape(png_bytes(), 1200000, 800000)
        b = box(0, 0, 800000, 1200000, kind="picture", shape=shape)
        b.pic_w, b.pic_h = 600000, 400000
        _emit_picture(b, ctx.doc, ctx)
        assert ctx.doc.inline_shapes[0].width == Emu(600000)
        assert ctx.report.images == 1

    def test_stretched_picture_keeps_the_frame_aspect(self):
        # A 3:2 image stretched into a square frame: Word must not fall back to the image's own aspect, or the
        # picture comes out a different height than the slide was planned with and can spill onto another page.
        ctx = make_ctx()
        shape = picture_shape(png_bytes(120, 80), 1200000, 1200000)
        b = box(0, 0, 1200000, 1200000, kind="picture", shape=shape)
        b.pic_w, b.pic_h = 600000, 600000
        _emit_picture(b, ctx.doc, ctx)
        assert (ctx.doc.inline_shapes[0].width, ctx.doc.inline_shapes[0].height) == (Emu(600000), Emu(600000))

    def test_undecodable_picture_is_skipped_with_a_warning(self):
        ctx = make_ctx()
        shape = picture_shape(b"not an image", 100, 100)
        _emit_picture(box(0, 0, 100, 100, kind="picture", shape=shape), ctx.doc, ctx)
        assert ctx.report.images == 0
        assert any("could not be decoded" in w for w in ctx.report.warnings)


# ---------- entry point ----------


def test_unreadable_file_raises_corrupt_file_error(tmp_path: Path):
    junk = tmp_path / "junk.pptx"
    junk.write_bytes(b"this is not a zip")
    with pytest.raises(CorruptFileError):
        convert_pptx(junk, tmp_path / "out.docx")
