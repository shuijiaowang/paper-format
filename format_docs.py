from __future__ import annotations

import argparse
import os
import re
from copy import deepcopy
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK, WD_LINE_SPACING, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph

OUTPUT_SUFFIX = "_f"
DEFAULT_INPUT_DIR = Path(__file__).with_name("process")

INTRO_TITLE_TEXT = "引言"
TOC_TITLE_TEXT = "目录"
CONCLUSION_TITLE_TEXT = "结语"
REFERENCES_TITLE_TEXT = "参考文献"
ACK_TITLE_TEXT = "致谢"

BODY_RULE = {
    "east_asia_font": "宋体",
    "size_pt": 12.0,
    "bold": False,
    "first_line_indent_chars": 2,
    "line_spacing_pt": 20.0,
    "alignment": "left",
}

H1_RULE = {
    "east_asia_font": "黑体",
    "size_pt": 15.0,
    "bold": False,
    "color_rgb": (0, 0, 0),
    "line_spacing": 1.5,
    "alignment": "center",
    "space_before_lines": 1,
    "space_after_lines": 1,
}

H2_RULE = {
    "east_asia_font": "黑体",
    "size_pt": 14.0,
    "bold": False,
    "color_rgb": (0, 0, 0),
    "line_spacing": 1.5,
    "alignment": "left",
}

H3_RULE = {
    "east_asia_font": "黑体",
    "size_pt": 12.0,
    "bold": False,
    "color_rgb": (0, 0, 0),
    "line_spacing": 1.5,
    "alignment": "left",
    "first_line_indent_chars": 2,
    "space_before_lines": 1,
}

REFERENCES_BODY_RULE = {
    "east_asia_font": "宋体",
    "size_pt": 10.5,
    "bold": False,
    "line_spacing": 1.5,
    "alignment": "left",
    "hanging_indent_chars": 2,
}

ACK_BODY_RULE = {
    "east_asia_font": "宋体",
    "size_pt": 10.5,
    "bold": False,
    "line_spacing_pt": 20.0,
    "alignment": "left",
    "first_line_indent_chars": 2,
}

FIGURE_CAPTION_RULE = {
    "east_asia_font": "黑体",
    "size_pt": 10.5,
    "bold": False,
    "line_spacing_pt": 20.0,
    "alignment": "center",
}

PAGE_NUMBER_RULE = {
    "east_asia_font": "宋体",
    "font_size_pt": 9.0,
    "alignment": "right",
    "start": 1,
}

CHINESE_H1_PREFIX_RE = r"[一二三四五六七八九十百千]+、"
CHINESE_H2_PREFIX_RE = r"[（(]\s*[一二三四五六七八九十百千]+\s*[)）]"
ARABIC_HEADING_PREFIX_RE = r"\d+(?:\.\d+)*(?:[.、．])?(?=\s|$)"
ARABIC_HEADING_STRIP_PREFIX_RE = r"(?:\d+(?:\.\d+)+(?:[.、．。])?|\d+[.、．。])"
HEADING_PREFIX_PATTERN = re.compile(
    rf"^(?:{CHINESE_H1_PREFIX_RE}|{CHINESE_H2_PREFIX_RE}(?:[.、．。])?|{ARABIC_HEADING_STRIP_PREFIX_RE})\s*"
)
FIGURE_PREFIX_PATTERN = re.compile(r"^图\s*\d+(?:\s*[-－]\s*\d+)?\s*")
DEFAULT_FIGURE_CAPTION_TEXT = "图1.1 题注"


def _debug_enabled() -> bool:
    return os.environ.get("PAPER_FORMAT_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}


def _debug_log(message: str) -> None:
    if _debug_enabled():
        print(f"[DEBUG] {message}")


def _resolve_alignment(alignment: str | None) -> WD_PARAGRAPH_ALIGNMENT | None:
    if not alignment:
        return None
    mapping = {
        "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
        "distributed": WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE,
        "distribute": WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE,
    }
    return mapping.get(alignment.lower())


def _set_run_font(run, east_asia_font: str, size_pt: float, bold: bool, color_rgb: tuple[int, int, int] | None = None) -> None:
    run.font.name = "Times New Roman"
    run.font.bold = bold
    run.font.size = Pt(size_pt)
    if color_rgb is not None:
        run.font.color.rgb = RGBColor(*color_rgb)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:cs"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), east_asia_font)


def _clear_indent(paragraph_format) -> None:
    paragraph_format.first_line_indent = None
    paragraph_format.left_indent = None
    paragraph_format.right_indent = None


def _clear_numbering(paragraph) -> None:
    p_pr = paragraph._p.find(qn("w:pPr"))
    if p_pr is None:
        return
    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is not None:
        p_pr.remove(num_pr)


def apply_paragraph_rule(paragraph, rule: dict) -> None:
    if paragraph._p is None:
        return
    _clear_numbering(paragraph)
    pf = paragraph.paragraph_format
    size_pt = float(rule.get("size_pt", 12.0))
    _clear_indent(pf)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    if "first_line_indent_chars" in rule:
        pf.first_line_indent = Pt(size_pt * float(rule["first_line_indent_chars"]))
    if "hanging_indent_chars" in rule:
        pf.first_line_indent = Pt(0)
        pf.left_indent = Pt(size_pt * float(rule["hanging_indent_chars"]))

    if "line_spacing_pt" in rule:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(float(rule["line_spacing_pt"]))
    elif "line_spacing" in rule:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = float(rule["line_spacing"])

    if "space_before_lines" in rule:
        pf.space_before = Pt(size_pt * float(rule["space_before_lines"]))
    if "space_after_lines" in rule:
        pf.space_after = Pt(size_pt * float(rule["space_after_lines"]))

    alignment = _resolve_alignment(rule.get("alignment"))
    if alignment is not None:
        paragraph.alignment = alignment

    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        _set_run_font(
            run,
            str(rule.get("east_asia_font", "宋体")),
            size_pt,
            bool(rule.get("bold", False)),
            tuple(rule["color_rgb"]) if "color_rgb" in rule else None,
        )


def _set_paragraph_text(paragraph, text: str) -> None:
    if not paragraph.runs:
        paragraph.add_run(text)
        return
    paragraph.runs[0].text = text
    for run in paragraph.runs[1:]:
        run.text = ""


def _prev_paragraph(paragraph) -> Paragraph | None:
    element = paragraph._p.getprevious()
    while element is not None and element.tag != qn("w:p"):
        element = element.getprevious()
    if element is None:
        return None
    return Paragraph(element, paragraph._parent)


def _next_paragraph(paragraph) -> Paragraph | None:
    element = paragraph._p.getnext()
    while element is not None and element.tag != qn("w:p"):
        element = element.getnext()
    if element is None:
        return None
    return Paragraph(element, paragraph._parent)


def _is_blank_paragraph(paragraph) -> bool:
    if paragraph._p is None:
        return True
    if paragraph._p.xpath(".//w:br[@w:type='page']"):
        return False
    if paragraph._p.xpath(".//w:drawing"):
        return False
    return not paragraph.text.strip()


def _insert_paragraph_before(paragraph) -> Paragraph:
    element = OxmlElement("w:p")
    paragraph._p.addprevious(element)
    return Paragraph(element, paragraph._parent)


def _insert_paragraph_after(paragraph) -> Paragraph:
    element = OxmlElement("w:p")
    paragraph._p.addnext(element)
    return Paragraph(element, paragraph._parent)


def _insert_blank_paragraph_after(paragraph) -> Paragraph:
    return _insert_paragraph_after(paragraph)


def _delete_paragraph(paragraph) -> None:
    parent = paragraph._element.getparent()
    if parent is not None:
        parent.remove(paragraph._element)


def _paragraph_has_page_break(paragraph) -> bool:
    if paragraph._p is None:
        return False
    return bool(paragraph._p.xpath(".//w:br[@w:type='page']"))


def _ensure_page_break_before(paragraph) -> None:
    prev = _prev_paragraph(paragraph)
    if prev is None or _paragraph_has_page_break(prev):
        return
    prev.add_run().add_break(WD_BREAK.PAGE)


def _make_blank_body_paragraph(paragraph) -> None:
    apply_paragraph_rule(paragraph, BODY_RULE)
    paragraph.paragraph_format.first_line_indent = Pt(0)


def _ensure_blank_after(paragraph) -> None:
    nxt = _next_paragraph(paragraph)
    if nxt is not None and _is_blank_paragraph(nxt):
        return
    blank = _insert_blank_paragraph_after(paragraph)
    _make_blank_body_paragraph(blank)


def _ensure_blank_before(paragraph) -> None:
    prev = _prev_paragraph(paragraph)
    if prev is not None and _is_blank_paragraph(prev):
        return
    blank = _insert_paragraph_before(paragraph)
    _make_blank_body_paragraph(blank)


def _remove_blanks_before(paragraph) -> None:
    while True:
        prev = _prev_paragraph(paragraph)
        if prev is None or not _is_blank_paragraph(prev):
            return
        _delete_paragraph(prev)


def _remove_blanks_after(paragraph) -> None:
    while True:
        nxt = _next_paragraph(paragraph)
        if nxt is None or not _is_blank_paragraph(nxt):
            return
        _delete_paragraph(nxt)


def _collapse_blanks_before(paragraph) -> None:
    while True:
        prev = _prev_paragraph(paragraph)
        if prev is None or not _is_blank_paragraph(prev):
            return
        prev_prev = _prev_paragraph(prev)
        if prev_prev is None or not _is_blank_paragraph(prev_prev):
            return
        _delete_paragraph(prev_prev)


def _collapse_blanks_after(paragraph) -> None:
    while True:
        nxt = _next_paragraph(paragraph)
        if nxt is None or not _is_blank_paragraph(nxt):
            return
        nxt_nxt = _next_paragraph(nxt)
        if nxt_nxt is None or not _is_blank_paragraph(nxt_nxt):
            return
        _delete_paragraph(nxt_nxt)


def _set_page_number_start(section, start: int) -> None:
    sect_pr = section._sectPr
    pg_num_type = sect_pr.find(qn("w:pgNumType"))
    if pg_num_type is None:
        pg_num_type = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num_type)
    pg_num_type.set(qn("w:start"), str(start))


def _add_page_number_run(paragraph) -> None:
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run = paragraph.add_run()
    run._element.append(fld_begin)
    run._element.append(instr_text)
    run._element.append(fld_end)


def _remove_footer_references(section) -> None:
    sect_pr = section._sectPr
    for node in list(sect_pr.findall(qn("w:footerReference"))):
        sect_pr.remove(node)


def _clear_footer(footer) -> None:
    for paragraph in footer.paragraphs:
        paragraph.text = ""


class Landmark(Enum):
    TOC = "toc"
    INTRO = "intro"
    CONCLUSION = "conclusion"
    REFERENCES = "references"
    ACK = "ack"


def _normalize_title_token(text: str) -> str:
    token = _strip_heading_prefix(text)
    token = re.sub(r"[\s·•\\-—_<>《》【】\\[\\]()（）:：,，。；;!！?？\"'`]+", "", token)
    return token


def _strip_heading_prefix(text: str) -> str:
    token = HEADING_PREFIX_PATTERN.sub("", text.strip()).strip()
    return re.sub(r"[，,、。．:：；;!！?？]+$", "", token).strip()


def _strip_figure_prefix(text: str) -> str:
    return FIGURE_PREFIX_PATTERN.sub("", text.strip()).strip()


def detect_landmark(paragraph) -> Landmark | None:
    text = paragraph.text.strip()
    if not text:
        return None
    normalized = _normalize_title_token(text)
    if normalized == TOC_TITLE_TEXT:
        return Landmark.TOC
    if normalized == INTRO_TITLE_TEXT:
        return Landmark.INTRO
    if normalized == CONCLUSION_TITLE_TEXT:
        return Landmark.CONCLUSION
    if normalized in {REFERENCES_TITLE_TEXT, "引用", "引用页"}:
        return Landmark.REFERENCES
    if normalized == ACK_TITLE_TEXT:
        return Landmark.ACK
    return None


@dataclass
class DocumentSegments:
    document: Document
    toc_title: Paragraph | None = None
    intro_title: Paragraph | None = None
    intro_body: list[Paragraph] = field(default_factory=list)
    main_body: list[Paragraph] = field(default_factory=list)
    conclusion_title: Paragraph | None = None
    conclusion_body: list[Paragraph] = field(default_factory=list)
    references_title: Paragraph | None = None
    references_body: list[Paragraph] = field(default_factory=list)
    ack_title: Paragraph | None = None
    ack_body: list[Paragraph] = field(default_factory=list)


@dataclass
class FormatReport:
    total_paragraphs: int = 0
    landmarks: list[str] = field(default_factory=list)
    body_titles: int = 0
    body_text: int = 0
    figure_captions: int = 0
    references_text: int = 0
    acknowledgment_text: int = 0
    section_titles: int = 0
    margins_updated: int = 0
    page_numbers_updated: int = 0
    toc_trimmed: bool = False
    missing_landmarks: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "total_paragraphs": self.total_paragraphs,
            "landmarks": self.landmarks,
            "body_titles": self.body_titles,
            "body_text": self.body_text,
            "figure_captions": self.figure_captions,
            "references_text": self.references_text,
            "acknowledgment_text": self.acknowledgment_text,
            "section_titles": self.section_titles,
            "margins_updated": self.margins_updated,
            "page_numbers_updated": self.page_numbers_updated,
            "toc_trimmed": self.toc_trimmed,
            "missing_landmarks": self.missing_landmarks,
        }


def scan_segments(document: Document) -> DocumentSegments:
    paragraphs = list(document.paragraphs)
    landmark_indexes: dict[Landmark, int] = {}
    for idx, paragraph in enumerate(paragraphs):
        landmark = detect_landmark(paragraph)
        if landmark is not None and landmark not in landmark_indexes:
            landmark_indexes[landmark] = idx
            _debug_log(f"landmark[{landmark.value}] idx={idx} text={paragraph.text!r}")

    intro_idx = landmark_indexes.get(Landmark.INTRO)
    conclusion_idx = landmark_indexes.get(Landmark.CONCLUSION)
    refs_idx = landmark_indexes.get(Landmark.REFERENCES)
    ack_idx = landmark_indexes.get(Landmark.ACK)
    toc_idx = landmark_indexes.get(Landmark.TOC)

    segments = DocumentSegments(document=document)
    if toc_idx is not None:
        segments.toc_title = paragraphs[toc_idx]
    if intro_idx is not None:
        segments.intro_title = paragraphs[intro_idx]
    if conclusion_idx is not None:
        segments.conclusion_title = paragraphs[conclusion_idx]
    if refs_idx is not None:
        segments.references_title = paragraphs[refs_idx]
    if ack_idx is not None:
        segments.ack_title = paragraphs[ack_idx]

    if intro_idx is None:
        return segments

    intro_end = min(
        (i for i in (conclusion_idx, refs_idx, ack_idx) if i is not None and i > intro_idx),
        default=len(paragraphs),
    )
    segments.intro_body = paragraphs[intro_idx + 1 : intro_end]

    if conclusion_idx is not None and conclusion_idx > intro_idx:
        segments.main_body = paragraphs[intro_idx + 1 : conclusion_idx]
        conclusion_end = min(
            (i for i in (refs_idx, ack_idx) if i is not None and i > conclusion_idx),
            default=len(paragraphs),
        )
        segments.conclusion_body = paragraphs[conclusion_idx + 1 : conclusion_end]
    else:
        main_end = min((i for i in (refs_idx, ack_idx) if i is not None and i > intro_idx), default=len(paragraphs))
        segments.main_body = paragraphs[intro_idx + 1 : main_end]

    if refs_idx is not None:
        refs_end = ack_idx if ack_idx is not None and ack_idx > refs_idx else len(paragraphs)
        segments.references_body = paragraphs[refs_idx + 1 : refs_end]

    if ack_idx is not None:
        segments.ack_body = paragraphs[ack_idx + 1 :]

    return segments


def _heading_level_by_prefix(text: str) -> int | None:
    candidate = text.strip()
    if re.match(rf"^{CHINESE_H1_PREFIX_RE}", candidate):
        return 1
    if re.match(rf"^{CHINESE_H2_PREFIX_RE}", candidate):
        return 2
    arabic_match = re.match(rf"^({ARABIC_HEADING_PREFIX_RE})", candidate)
    if arabic_match:
        depth = arabic_match.group(1).rstrip(".、．").count(".") + 1
        return min(depth, 3)
    return None


def _resolve_style_key(paragraph) -> str:
    style_name = (paragraph.style.name or "").lower()
    if "heading 1" in style_name or "标题 1" in style_name:
        return "h1"
    if "heading 2" in style_name or "标题 2" in style_name:
        return "h2"
    if "heading 3" in style_name or "标题 3" in style_name:
        return "h3"
    return "body"


def _looks_like_heading(paragraph) -> bool:
    return _resolve_style_key(paragraph) in {"h1", "h2", "h3"}


def _resolve_heading_level(paragraph) -> int:
    style_key = _resolve_style_key(paragraph)
    if style_key == "h2":
        return 2
    if style_key == "h3":
        return 3
    return 1


def _is_figure_caption(text: str) -> bool:
    stripped = text.strip()
    return bool(re.match(r"^图\s*\d+", stripped)) or bool(re.match(r"^\d+(?:\.\d+)+\s*题注", stripped))


def _apply_figure_caption(paragraph, text: str) -> None:
    _set_paragraph_text(paragraph, text.strip())
    apply_paragraph_rule(paragraph, FIGURE_CAPTION_RULE)
    paragraph.paragraph_format.first_line_indent = Pt(0)


def _apply_page_title(paragraph, title_text: str) -> None:
    _set_paragraph_text(paragraph, title_text)
    apply_paragraph_rule(paragraph, H1_RULE)
    paragraph.paragraph_format.space_before = Pt(0)
    _ensure_page_break_before(paragraph)
    _remove_blanks_before(paragraph)
    _ensure_blank_after(paragraph)
    _collapse_blanks_after(paragraph)


def _format_body_paragraph(paragraph) -> None:
    apply_paragraph_rule(paragraph, BODY_RULE)


def _format_h1(paragraph, prefix: str, plain_text: str) -> None:
    _set_paragraph_text(paragraph, f"{prefix}{plain_text}")
    apply_paragraph_rule(paragraph, H1_RULE)
    _ensure_blank_before(paragraph)
    _collapse_blanks_before(paragraph)
    _ensure_blank_after(paragraph)
    _collapse_blanks_after(paragraph)


def _format_h2(paragraph, prefix: str, plain_text: str) -> None:
    _set_paragraph_text(paragraph, f"{prefix}{plain_text}")
    apply_paragraph_rule(paragraph, H2_RULE)
    _remove_blanks_before(paragraph)
    _remove_blanks_after(paragraph)


def _format_h3(paragraph, prefix: str, plain_text: str) -> None:
    _set_paragraph_text(paragraph, f"{prefix}{plain_text}")
    apply_paragraph_rule(paragraph, H3_RULE)
    _ensure_blank_before(paragraph)
    _collapse_blanks_before(paragraph)
    _remove_blanks_after(paragraph)


def _format_references_body(paragraph) -> None:
    apply_paragraph_rule(paragraph, REFERENCES_BODY_RULE)


def _format_ack_body(paragraph) -> None:
    apply_paragraph_rule(paragraph, ACK_BODY_RULE)


def _chinese_numeral(value: int) -> str:
    numerals = "零一二三四五六七八九"
    if value <= 10:
        return "十" if value == 10 else numerals[value]
    if value < 20:
        return f"十{numerals[value % 10]}"
    if value < 100:
        tens, ones = divmod(value, 10)
        prefix = f"{numerals[tens]}十"
        return prefix if ones == 0 else f"{prefix}{numerals[ones]}"
    return str(value)


def _section_index_for_paragraph(document: Document, paragraph) -> int:
    target_p = paragraph._p
    section_idx = 0
    body = document._body._element
    for child in body.iterchildren():
        if child is target_p:
            return min(section_idx, len(document.sections) - 1)
        if child.tag != qn("w:p"):
            continue
        p_pr = child.find(qn("w:pPr"))
        sect_pr = p_pr.find(qn("w:sectPr")) if p_pr is not None else None
        if sect_pr is not None:
            section_idx += 1
    return min(section_idx, len(document.sections) - 1)


def _ensure_section_break_at(document: Document, paragraph):
    prev = _prev_paragraph(paragraph)
    sections = list(document.sections)
    if prev is None:
        return sections[0]

    p_pr = prev._p.find(qn("w:pPr"))
    existing_sect_pr = p_pr.find(qn("w:sectPr")) if p_pr is not None else None
    if existing_sect_pr is not None:
        for idx, section in enumerate(sections):
            if section._sectPr is existing_sect_pr:
                return sections[min(idx + 1, len(sections) - 1)]
        return sections[-1]

    body_sect_pr = document._body._element.sectPr
    if body_sect_pr is None:
        return sections[0]
    p_pr = prev._p.get_or_add_pPr()
    p_pr.append(deepcopy(body_sect_pr))
    return document.sections[_section_index_for_paragraph(document, paragraph)]


def trim_toc_between_catalog_and_intro(segments: DocumentSegments, report: FormatReport) -> None:
    if segments.toc_title is None or segments.intro_title is None:
        return
    toc = segments.toc_title
    intro = segments.intro_title
    paragraphs = list(segments.document.paragraphs)
    toc_idx = next((idx for idx, paragraph in enumerate(paragraphs) if paragraph._p is toc._p), None)
    intro_idx = next((idx for idx, paragraph in enumerate(paragraphs) if paragraph._p is intro._p), None)
    if toc_idx is None or intro_idx is None or toc_idx >= intro_idx:
        return
    current = toc
    removed_any = False
    while current is not None and current._p is not intro._p:
        nxt = _next_paragraph(current)
        _delete_paragraph(current)
        removed_any = True
        current = nxt
    report.toc_trimmed = removed_any


def _next_figure_label(chapter_figure_counts: dict[int, int], chapter_idx: int) -> str:
    safe_chapter = max(chapter_idx, 1)
    chapter_figure_counts[safe_chapter] = chapter_figure_counts.get(safe_chapter, 0) + 1
    return f"图{safe_chapter}-{chapter_figure_counts[safe_chapter]}"


def _try_apply_image_wrap(paragraph) -> None:
    if paragraph._p is None:
        return

    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    image_runs = [run for run in paragraph.runs if run._element.xpath(".//w:drawing")]
    if len(image_runs) > 1:
        for run in image_runs[:-1]:
            if not run._element.xpath("./w:br"):
                run.add_break(WD_BREAK.LINE)

    def _child_by_local_name(parent, local_name: str):
        for child in parent:
            if child.tag.rsplit("}", 1)[-1] == local_name:
                return child
        return None

    def _build_top_bottom_anchor(inline):
        anchor = OxmlElement("wp:anchor")
        for key, value in {
            "distT": "0",
            "distB": "0",
            "distL": "0",
            "distR": "0",
            "simplePos": "0",
            "relativeHeight": "0",
            "behindDoc": "0",
            "locked": "0",
            "layoutInCell": "1",
            "allowOverlap": "0",
        }.items():
            anchor.set(key, value)

        simple_pos = OxmlElement("wp:simplePos")
        simple_pos.set("x", "0")
        simple_pos.set("y", "0")
        anchor.append(simple_pos)

        position_h = OxmlElement("wp:positionH")
        position_h.set("relativeFrom", "column")
        align = OxmlElement("wp:align")
        align.text = "center"
        position_h.append(align)
        anchor.append(position_h)

        position_v = OxmlElement("wp:positionV")
        position_v.set("relativeFrom", "paragraph")
        pos_offset = OxmlElement("wp:posOffset")
        pos_offset.text = "0"
        position_v.append(pos_offset)
        anchor.append(position_v)

        for local_name in ("extent", "effectExtent"):
            child = _child_by_local_name(inline, local_name)
            if child is not None:
                anchor.append(deepcopy(child))

        anchor.append(OxmlElement("wp:wrapTopAndBottom"))

        for local_name in ("docPr", "cNvGraphicFramePr", "graphic"):
            child = _child_by_local_name(inline, local_name)
            if child is not None:
                anchor.append(deepcopy(child))

        return anchor

    for inline in paragraph._p.xpath(".//wp:inline"):
        parent = inline.getparent()
        if parent is not None:
            parent.replace(inline, _build_top_bottom_anchor(inline))

    for anchor in paragraph._p.xpath(".//wp:anchor"):
        for key, value in {
            "distT": "0",
            "distB": "0",
            "distL": "0",
            "distR": "0",
            "behindDoc": "0",
            "locked": "0",
            "layoutInCell": "1",
            "allowOverlap": "0",
        }.items():
            anchor.set(key, value)

        wrap = _child_by_local_name(anchor, "wrapTopAndBottom")
        if wrap is None:
            wrap = OxmlElement("wp:wrapTopAndBottom")
            insert_at = next(
                (
                    idx
                    for idx, child in enumerate(anchor)
                    if child.tag.rsplit("}", 1)[-1] in {"docPr", "cNvGraphicFramePr", "graphic"}
                ),
                len(anchor),
            )
            anchor.insert(insert_at, wrap)

        position_h = _child_by_local_name(anchor, "positionH")
        if position_h is None:
            position_h = OxmlElement("wp:positionH")
            position_h.set("relativeFrom", "column")
            anchor.insert(1, position_h)
        else:
            position_h.set("relativeFrom", "column")
            for child in list(position_h):
                position_h.remove(child)
        align = OxmlElement("wp:align")
        align.text = "center"
        position_h.append(align)


def format_intro_and_main_body(segments: DocumentSegments, report: FormatReport) -> None:
    if segments.intro_title is None:
        report.missing_landmarks.append(INTRO_TITLE_TEXT)
        return

    _apply_page_title(segments.intro_title, INTRO_TITLE_TEXT)
    report.section_titles += 1

    chapter_idx = 0
    section_idx = 0
    sub_idx = 0
    pending_figure_anchor: Paragraph | None = None

    def flush_pending_figure_caption() -> None:
        nonlocal pending_figure_anchor
        if pending_figure_anchor is None:
            return
        caption = _insert_paragraph_after(pending_figure_anchor)
        _apply_figure_caption(caption, DEFAULT_FIGURE_CAPTION_TEXT)
        report.figure_captions += 1
        pending_figure_anchor = None

    def handle_content(paragraph, allow_headings: bool, body_rule_kind: str = "body") -> None:
        nonlocal chapter_idx, section_idx, sub_idx, pending_figure_anchor
        if paragraph._p is None:
            return
        text = paragraph.text.strip()

        if paragraph._p.xpath(".//w:footnoteReference"):
            return
        if paragraph._p.xpath(".//w:drawing"):
            _try_apply_image_wrap(paragraph)
            pending_figure_anchor = paragraph
            return
        if not text:
            return

        if pending_figure_anchor is not None:
            if _is_figure_caption(text):
                _apply_figure_caption(paragraph, text)
                report.figure_captions += 1
                pending_figure_anchor = None
                return
            flush_pending_figure_caption()

        if _is_figure_caption(text):
            _apply_figure_caption(paragraph, text)
            report.figure_captions += 1
            return

        if allow_headings and _looks_like_heading(paragraph):
            level = _resolve_heading_level(paragraph)
            plain_text = _strip_heading_prefix(text)
            if level == 1:
                chapter_idx += 1
                section_idx = 0
                _format_h1(paragraph, f"{_chinese_numeral(chapter_idx)}、", plain_text)
            elif level == 2:
                if chapter_idx == 0:
                    chapter_idx = 1
                section_idx += 1
                _format_h2(paragraph, f"（{_chinese_numeral(section_idx)}）", plain_text)
            else:
                sub_idx += 1
                _format_h3(paragraph, f"{sub_idx}. ", plain_text)
            report.body_titles += 1
            return

        if body_rule_kind == "ack":
            _format_ack_body(paragraph)
            report.acknowledgment_text += 1
        else:
            _format_body_paragraph(paragraph)
            report.body_text += 1

    for paragraph in segments.main_body:
        handle_content(paragraph, allow_headings=True)
    flush_pending_figure_caption()

    if segments.conclusion_title is not None:
        _apply_page_title(segments.conclusion_title, CONCLUSION_TITLE_TEXT)
        report.section_titles += 1
    for paragraph in segments.conclusion_body:
        handle_content(paragraph, allow_headings=False)
    flush_pending_figure_caption()


def format_references(segments: DocumentSegments, report: FormatReport) -> None:
    if segments.references_title is None:
        report.missing_landmarks.append(REFERENCES_TITLE_TEXT)
        return
    _apply_page_title(segments.references_title, REFERENCES_TITLE_TEXT)
    report.section_titles += 1
    for paragraph in segments.references_body:
        if _is_blank_paragraph(paragraph):
            continue
        _format_references_body(paragraph)
        report.references_text += 1


def format_acknowledgment(segments: DocumentSegments, report: FormatReport) -> None:
    if segments.ack_title is None:
        report.missing_landmarks.append(ACK_TITLE_TEXT)
        return
    _apply_page_title(segments.ack_title, ACK_TITLE_TEXT)
    report.section_titles += 1
    for paragraph in segments.ack_body:
        if _is_blank_paragraph(paragraph) or paragraph._p.xpath(".//w:footnoteReference"):
            continue
        _format_ack_body(paragraph)
        report.acknowledgment_text += 1


def apply_document_setup(document: Document, intro_title: Paragraph | None, report: FormatReport) -> None:
    for section in document.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
    report.margins_updated = len(document.sections)

    if intro_title is None:
        return

    intro_section = _ensure_section_break_at(document, intro_title)
    intro_sect_pr = intro_section._sectPr
    sections = list(document.sections)
    intro_section_idx = 0
    for idx, section in enumerate(sections):
        if section._sectPr is intro_sect_pr:
            intro_section_idx = idx
            break

    for idx, section in enumerate(sections):
        section.different_first_page_header_footer = False
        section.odd_and_even_pages_header_footer = False
        if idx < intro_section_idx:
            _remove_footer_references(section)
            continue
        footer = section.footer
        footer.is_linked_to_previous = False
        _clear_footer(footer)

    _set_page_number_start(intro_section, int(PAGE_NUMBER_RULE["start"]))

    intro_reached = False
    for section in sections:
        if section._sectPr is intro_sect_pr:
            intro_reached = True
        if not intro_reached:
            continue
        footer = section.footer
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_para.text = ""
        footer_para.alignment = _resolve_alignment(str(PAGE_NUMBER_RULE["alignment"])) or WD_PARAGRAPH_ALIGNMENT.RIGHT
        _add_page_number_run(footer_para)
        for run in footer_para.runs:
            _set_run_font(
                run,
                str(PAGE_NUMBER_RULE["east_asia_font"]),
                float(PAGE_NUMBER_RULE["font_size_pt"]),
                False,
            )
    report.page_numbers_updated = max(len(sections) - intro_section_idx, 0)


def apply_mvp_format(doc_path: Path, output_path: Path) -> FormatReport:
    document = Document(str(doc_path))
    report = FormatReport(total_paragraphs=len(document.paragraphs))
    segments = scan_segments(document)

    landmark_names: list[str] = []
    if segments.toc_title is not None:
        landmark_names.append(TOC_TITLE_TEXT)
    if segments.intro_title is not None:
        landmark_names.append(INTRO_TITLE_TEXT)
    if segments.conclusion_title is not None:
        landmark_names.append(CONCLUSION_TITLE_TEXT)
    if segments.references_title is not None:
        landmark_names.append(REFERENCES_TITLE_TEXT)
    if segments.ack_title is not None:
        landmark_names.append(ACK_TITLE_TEXT)
    report.landmarks = landmark_names

    trim_toc_between_catalog_and_intro(segments, report)
    segments = scan_segments(document)
    format_intro_and_main_body(segments, report)
    format_references(segments, report)
    format_acknowledgment(segments, report)
    apply_document_setup(document, segments.intro_title, report)
    document.save(str(output_path))
    return report


def iter_docx_files(root_dir: Path) -> list[Path]:
    return [
        path
        for path in root_dir.rglob("*.docx")
        if path.is_file() and not path.name.startswith("~$") and not path.stem.endswith(OUTPUT_SUFFIX)
    ]


def build_output_path(src: Path) -> Path:
    return src.with_name(f"{src.stem}{OUTPUT_SUFFIX}{src.suffix}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Format .docx files in a directory and generate copies.")
    parser.add_argument(
        "directory",
        nargs="?",
        type=Path,
        help=f"Target directory containing .docx files (default: {DEFAULT_INPUT_DIR})",
    )
    parser.add_argument("--debug", action="store_true", help="Enable debug logs")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.debug:
        os.environ["PAPER_FORMAT_DEBUG"] = "1"

    target_dir = (args.directory or DEFAULT_INPUT_DIR).resolve()
    if not target_dir.exists() or not target_dir.is_dir():
        print(f"[ERROR] Directory not found: {target_dir}")
        return 1

    docx_files = iter_docx_files(target_dir)
    if not docx_files:
        print(f"[INFO] No .docx files found in: {target_dir}")
        return 0

    for src in docx_files:
        out = build_output_path(src)
        apply_mvp_format(src, out)
        print(f"[OK] {src} -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
