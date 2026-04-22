from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml import parse_xml
from docx.parts.numbering import NumberingPart
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph

OUTPUT_SUFFIX = "_formatted"
DEFAULT_CONFIG_PATH = Path(__file__).with_name("config.json")
STYLE_BODY = "body"
STYLE_H1 = "h1"
STYLE_H2 = "h2"
STYLE_H3 = "h3"
SPECIAL_H1_TITLES = {"引言", "结语", "参考文献", "致谢"}
H1_PREFIXES = "一二三四五六七八九十"


def _resolve_alignment(alignment: str | None) -> WD_PARAGRAPH_ALIGNMENT | None:
    if not alignment:
        return None
    mapping = {
        "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
    }
    return mapping.get(alignment.lower())


def _set_run_font(run, east_asia_font: str, size_pt: float, bold: bool) -> None:
    run.font.name = "Times New Roman"
    run.font.bold = bold
    run.font.size = Pt(size_pt)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:cs"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), east_asia_font)


def _clear_indent(pf) -> None:
    pf.first_line_indent = None
    pf.left_indent = None


def apply_paragraph_rule(paragraph, rule: dict[str, float | int | str | bool]) -> None:
    pf = paragraph.paragraph_format
    size_pt = float(rule.get("size_pt", 12))
    _clear_indent(pf)

    if "first_line_indent_chars" in rule:
        pf.first_line_indent = Pt(size_pt * int(rule["first_line_indent_chars"]))
    if "hanging_indent_chars" in rule:
        pf.first_line_indent = Pt(0)
        pf.left_indent = Pt(size_pt * int(rule["hanging_indent_chars"]))

    if "line_spacing_pt" in rule:
        pf.line_spacing = Pt(float(rule["line_spacing_pt"]))
    elif "line_spacing" in rule:
        pf.line_spacing = float(rule["line_spacing"])

    if "space_before_lines" in rule:
        pf.space_before = Pt(size_pt * float(rule["space_before_lines"]))
    if "space_after_lines" in rule:
        pf.space_after = Pt(size_pt * float(rule["space_after_lines"]))

    para_alignment = _resolve_alignment(rule.get("alignment"))
    if para_alignment is not None:
        paragraph.alignment = para_alignment

    for run in paragraph.runs:
        _set_run_font(run, str(rule.get("east_asia_font", "宋体")), size_pt, bool(rule.get("bold", False)))


def load_config(config_path: Path) -> dict:
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")
    return json.loads(config_path.read_text(encoding="utf-8"))


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


def _clear_footer(footer) -> None:
    for paragraph in footer.paragraphs:
        paragraph.text = ""


def _insert_blank_paragraph_before(paragraph):
    p = OxmlElement("w:p")
    paragraph._p.addprevious(p)
    return Paragraph(p, paragraph._parent)


def _insert_blank_paragraph_after(paragraph):
    p = OxmlElement("w:p")
    paragraph._p.addnext(p)
    return Paragraph(p, paragraph._parent)


def _is_blank_paragraph(paragraph) -> bool:
    # A paragraph that only contains a page break has empty text,
    # but it is structural content and must not be treated as blank.
    if paragraph._p.xpath(".//w:br[@w:type='page']"):
        return False
    return not paragraph.text.strip()


def _delete_paragraph(paragraph) -> None:
    p = paragraph._element
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)
    paragraph._p = paragraph._element = None


def _ensure_margins(document: Document) -> None:
    for section in document.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)


def _find_start_index(document: Document) -> int:
    for idx, paragraph in enumerate(document.paragraphs):
        if _is_intro_title(paragraph.text):
            return idx

    for idx, paragraph in enumerate(document.paragraphs):
        if paragraph.text.strip():
            return idx
    return 0


def _is_figure_caption(text: str) -> bool:
    return bool(re.match(r"^图\d+(-\d+)?", text))


def _normalize_title_token(text: str) -> str:
    token = text.strip()
    token = re.sub(r"^[（(]\s*[一二三四五六七八九十]+\s*[)）]", "", token)
    token = re.sub(r"^[一二三四五六七八九十]+、", "", token)
    token = re.sub(r"^\d+(?:\.\d+)*\.", "", token)
    token = re.sub(r"[\s·•\\-—_<>《》【】\\[\\]()（）:：,，。；;!！?？\"'`]+", "", token)
    return token


def _reference_title_tokens(config: dict) -> set[str]:
    structure_config = config.get("structure", {})
    configured = _normalize_title_token(structure_config.get("references_title", "参考文献"))
    tokens = {"参考文献", "引用", "引用页"}
    if configured:
        tokens.add(configured)
    return {_normalize_title_token(token) for token in tokens if token}


def _is_reference_title(text: str, config: dict) -> bool:
    return _normalize_title_token(text) in _reference_title_tokens(config)


def _strip_number_prefix(text: str) -> str:
    candidate = text.strip()
    candidate = re.sub(r"^[一二三四五六七八九十]+、", "", candidate)
    candidate = re.sub(r"^（[一二三四五六七八九十]+）", "", candidate)
    candidate = re.sub(r"^\(\s*[一二三四五六七八九十]+\s*\)", "", candidate)
    candidate = re.sub(r"^\d+(?:\.\d+)*\.", "", candidate)
    return candidate.strip()


def _is_intro_title(text: str) -> bool:
    return _normalize_title_token(text) == "引言"


def _heading_level_by_prefix(text: str) -> int | None:
    t = text.strip()
    if re.match(r"^[一二三四五六七八九十]+、", t):
        return 1
    if re.match(r"^(（[一二三四五六七八九十]+）|\(\s*[一二三四五六七八九十]+\s*\))", t):
        return 2
    if re.match(r"^\d+(?:\.\d+)*\.", t):
        return 3
    return None


def _resolve_rule_key(paragraph) -> str:
    style_name = (paragraph.style.name or "").lower()
    if "heading 1" in style_name or "标题 1" in style_name:
        return STYLE_H1
    if "heading 2" in style_name or "标题 2" in style_name:
        return STYLE_H2
    if "heading 3" in style_name or "标题 3" in style_name:
        return STYLE_H3
    return STYLE_BODY


def _looks_like_heading(paragraph) -> bool:
    text = paragraph.text.strip()
    if not text:
        return False
    style_key = _resolve_rule_key(paragraph)
    if style_key in {STYLE_H1, STYLE_H2, STYLE_H3}:
        return True
    if text in SPECIAL_H1_TITLES:
        return True
    if _heading_level_by_prefix(text) is not None:
        return True
    if len(text) <= 24 and not any(ch in text for ch in "，。；：！？,.?!:;"):
        return True
    return False


def _set_paragraph_text(paragraph, text: str) -> None:
    if not paragraph.runs:
        paragraph.add_run(text)
        return
    paragraph.runs[0].text = text
    for run in paragraph.runs[1:]:
        run.text = ""


def _get_or_create_numbering_part(document: Document) -> NumberingPart:
    try:
        return document.part.numbering_part
    except NotImplementedError:
        package = document.part.package
        numbering_part = NumberingPart(
            PackURI("/word/numbering.xml"),
            CT.WML_NUMBERING,
            parse_xml(
                '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
            ),
            package,
        )
        document.part.relate_to(numbering_part, RT.NUMBERING)
        return numbering_part


def _create_heading_numbering(document: Document) -> int:
    numbering_part = _get_or_create_numbering_part(document)
    numbering = numbering_part.numbering_definitions._numbering
    abstract_nums = numbering.xpath("./w:abstractNum")
    nums = numbering.xpath("./w:num")
    next_abstract_id = (
        max((int(node.get(qn("w:abstractNumId"))) for node in abstract_nums), default=-1) + 1
    )
    next_num_id = max((int(node.get(qn("w:numId"))) for node in nums), default=0) + 1

    abstract_num = OxmlElement("w:abstractNum")
    abstract_num.set(qn("w:abstractNumId"), str(next_abstract_id))

    def _append_numbering_rpr(lvl, size_pt: float) -> None:
        r_pr = OxmlElement("w:rPr")
        r_fonts = OxmlElement("w:rFonts")
        r_fonts.set(qn("w:ascii"), "Times New Roman")
        r_fonts.set(qn("w:hAnsi"), "Times New Roman")
        r_fonts.set(qn("w:cs"), "Times New Roman")
        r_fonts.set(qn("w:eastAsia"), "Times New Roman")
        r_pr.append(r_fonts)

        size_half_points = str(int(round(size_pt * 2)))
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), size_half_points)
        sz_cs = OxmlElement("w:szCs")
        sz_cs.set(qn("w:val"), size_half_points)
        r_pr.extend([sz, sz_cs])
        lvl.append(r_pr)

    # 一级：一、二、三、
    lvl0 = OxmlElement("w:lvl")
    lvl0.set(qn("w:ilvl"), "0")
    start0 = OxmlElement("w:start")
    start0.set(qn("w:val"), "1")
    numfmt0 = OxmlElement("w:numFmt")
    numfmt0.set(qn("w:val"), "chineseCounting")
    lvltext0 = OxmlElement("w:lvlText")
    lvltext0.set(qn("w:val"), "%1、")
    lvl0.extend([start0, numfmt0, lvltext0])
    _append_numbering_rpr(lvl0, 15.0)

    # 二级：（一）（二）（三）
    lvl1 = OxmlElement("w:lvl")
    lvl1.set(qn("w:ilvl"), "1")
    start1 = OxmlElement("w:start")
    start1.set(qn("w:val"), "1")
    numfmt1 = OxmlElement("w:numFmt")
    numfmt1.set(qn("w:val"), "chineseCounting")
    lvltext1 = OxmlElement("w:lvlText")
    lvltext1.set(qn("w:val"), "（%2）")
    lvl1.extend([start1, numfmt1, lvltext1])
    _append_numbering_rpr(lvl1, 14.0)

    # 三级：1. 2. 3.
    lvl2 = OxmlElement("w:lvl")
    lvl2.set(qn("w:ilvl"), "2")
    start2 = OxmlElement("w:start")
    start2.set(qn("w:val"), "1")
    numfmt2 = OxmlElement("w:numFmt")
    numfmt2.set(qn("w:val"), "decimal")
    lvltext2 = OxmlElement("w:lvlText")
    lvltext2.set(qn("w:val"), "%3.")
    lvl2.extend([start2, numfmt2, lvltext2])
    _append_numbering_rpr(lvl2, 12.0)

    abstract_num.extend([lvl0, lvl1, lvl2])
    numbering.append(abstract_num)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(next_num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(next_abstract_id))
    num.append(abstract_ref)
    numbering.append(num)
    return next_num_id


def _set_heading_numbering(paragraph, num_id: int, level: int) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    num_pr = p_pr.numPr
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        p_pr.append(num_pr)

    ilvl = num_pr.find(qn("w:ilvl"))
    if ilvl is None:
        ilvl = OxmlElement("w:ilvl")
        num_pr.append(ilvl)
    ilvl.set(qn("w:val"), str(level))

    num_id_node = num_pr.find(qn("w:numId"))
    if num_id_node is None:
        num_id_node = OxmlElement("w:numId")
        num_pr.append(num_id_node)
    num_id_node.set(qn("w:val"), str(num_id))


def _format_heading_text(text: str, level: int, i1: int, i2: int, i3: int) -> str:
    core = _strip_number_prefix(text)
    if level == 1:
        if not (1 <= i1 <= len(H1_PREFIXES)):
            prefix = f"{i1}、"
        else:
            prefix = f"{H1_PREFIXES[i1 - 1]}、"
        return f"{prefix}{core}"
    if level == 2:
        if not (1 <= i2 <= len(H1_PREFIXES)):
            prefix = f"({i2})"
        else:
            prefix = f"（{H1_PREFIXES[i2 - 1]}）"
        return f"{prefix}{core}"
    return f"{i3}.{core}"


def _find_section_by_sectpr(document: Document, sect_pr):
    for section in document.sections:
        if section._sectPr is sect_pr:
            return section
    return document.sections[-1]


def _insert_intro_section_break(document: Document, start_idx: int):
    if start_idx <= 0:
        return document.sections[0]
    p = document.paragraphs[start_idx - 1]
    sec_pr = p._p.pPr.sectPr if p._p.pPr is not None else None
    if sec_pr is not None:
        return _find_section_by_sectpr(document, sec_pr)
    new_section = document.add_section(WD_SECTION_START.NEW_PAGE)
    moved = new_section._sectPr
    body = document._body._element
    body.remove(moved)
    p._p.addnext(moved)
    return _find_section_by_sectpr(document, moved)


def apply_page_number_from_intro(document: Document, start_idx: int, page_number_config: dict) -> None:
    intro_section = _insert_intro_section_break(document, start_idx)
    _ensure_margins(document)

    for section in document.sections:
        section.different_first_page_header_footer = False
        section.odd_and_even_pages_header_footer = False
        footer = section.footer
        footer.is_linked_to_previous = False
        _clear_footer(footer)

    _set_page_number_start(intro_section, int(page_number_config.get("start", 1)))
    intro_found = False
    for section in document.sections:
        if section is intro_section:
            intro_found = True
        if not intro_found:
            continue
        footer = section.footer
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_para.text = ""
        footer_para.alignment = _resolve_alignment(page_number_config.get("alignment")) or WD_PARAGRAPH_ALIGNMENT.RIGHT
        _add_page_number_run(footer_para)
        for run in footer_para.runs:
            _set_run_font(
                run,
                str(page_number_config.get("east_asia_font", "宋体")),
                float(page_number_config.get("font_size_pt", 9)),
                False,
            )


def _enforce_structure(document: Document, config: dict, start_idx: int) -> None:
    style_config = config["styles"]
    page_only_titles = _reference_title_tokens(config) | {"引言", "致谢"}

    idx = start_idx
    while idx < len(document.paragraphs):
        paragraph = document.paragraphs[idx]
        text = paragraph.text.strip()
        if not text:
            idx += 1
            continue

        normalized = _normalize_title_token(text)
        plain_title = _strip_number_prefix(text).strip()
        level = _heading_level_by_prefix(text)
        is_intro_title = _is_intro_title(text)
        is_conclusion_title = normalized == "结语"
        is_page_only_title = normalized in page_only_titles
        is_h1 = level == 1
        is_h3 = level == 3

        # 只对标题执行结构空行规则，避免清理正文中的手动换行。
        if not (is_intro_title or is_page_only_title or is_conclusion_title or is_h1 or is_h3 or level == 2):
            idx += 1
            continue

        # 空行规则：
        # - 引言：前换页，后空一行
        # - 正文一级标题/结语：上下一行
        # - 二级标题：上下不空行
        # - 三级标题：上空一下空零
        # - 参考文献/致谢：前换页，后空一行
        needs_blank_before = is_h1 or is_h3 or is_conclusion_title
        if is_intro_title or is_page_only_title:
            needs_blank_before = False
        needs_blank_after = is_h1 or is_intro_title or is_page_only_title or is_conclusion_title

        # 引言/参考文献/致谢前不空行，并强制换页。
        if is_page_only_title and idx > 0:
            prev_para = document.paragraphs[idx - 1]
            prev_para.add_run().add_break(WD_BREAK.PAGE)
            while idx > 0 and _is_blank_paragraph(document.paragraphs[idx - 1]):
                _delete_paragraph(document.paragraphs[idx - 1])
                idx -= 1
            paragraph = document.paragraphs[idx]
            paragraph.paragraph_format.space_before = Pt(0)
            apply_paragraph_rule(paragraph, style_config[STYLE_H1])
        elif is_intro_title:
            while idx > 0 and _is_blank_paragraph(document.paragraphs[idx - 1]):
                _delete_paragraph(document.paragraphs[idx - 1])
                idx -= 1
            paragraph = document.paragraphs[idx]
            paragraph.paragraph_format.space_before = Pt(0)

        if needs_blank_before:
            if idx == 0 or _is_blank_paragraph(document.paragraphs[idx - 1]):
                while idx > 1 and _is_blank_paragraph(document.paragraphs[idx - 1]) and _is_blank_paragraph(document.paragraphs[idx - 2]):
                    _delete_paragraph(document.paragraphs[idx - 2])
                    idx -= 1
            else:
                blank_before = _insert_blank_paragraph_before(paragraph)
                apply_paragraph_rule(blank_before, style_config[STYLE_BODY])
                blank_before.paragraph_format.first_line_indent = Pt(0)
                idx += 1
                paragraph = document.paragraphs[idx]
        else:
            while idx > 0 and _is_blank_paragraph(document.paragraphs[idx - 1]):
                _delete_paragraph(document.paragraphs[idx - 1])
                idx -= 1
                paragraph = document.paragraphs[idx]

        if needs_blank_after:
            if idx + 1 >= len(document.paragraphs):
                blank_after = _insert_blank_paragraph_after(paragraph)
                apply_paragraph_rule(blank_after, style_config[STYLE_BODY])
                blank_after.paragraph_format.first_line_indent = Pt(0)
            elif not _is_blank_paragraph(document.paragraphs[idx + 1]):
                blank_after = _insert_blank_paragraph_after(paragraph)
                apply_paragraph_rule(blank_after, style_config[STYLE_BODY])
                blank_after.paragraph_format.first_line_indent = Pt(0)
            while idx + 2 < len(document.paragraphs) and _is_blank_paragraph(document.paragraphs[idx + 1]) and _is_blank_paragraph(document.paragraphs[idx + 2]):
                _delete_paragraph(document.paragraphs[idx + 2])
        else:
            while idx + 1 < len(document.paragraphs) and _is_blank_paragraph(document.paragraphs[idx + 1]):
                _delete_paragraph(document.paragraphs[idx + 1])

        idx += 1


def apply_required_format(document: Document, config: dict) -> int:
    style_config = config["styles"]
    start_idx = _find_start_index(document)

    chapter_idx = 0
    section_idx = 0
    sub_idx = 0
    in_references = False
    in_ack = False
    heading_num_id = _create_heading_numbering(document)

    for idx, paragraph in enumerate(document.paragraphs):
        if idx < start_idx:
            continue
        text = paragraph.text.strip()
        if not text:
            continue
        if paragraph._p.xpath(".//w:footnoteReference"):
            continue
        if paragraph._p.xpath(".//w:drawing"):
            continue
        if _is_figure_caption(text):
            apply_paragraph_rule(paragraph, style_config[STYLE_BODY])
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            paragraph.paragraph_format.first_line_indent = Pt(0)
            for run in paragraph.runs:
                _set_run_font(run, "黑体", 10.5, False)
            continue

        stripped = _strip_number_prefix(text)
        normalized = _normalize_title_token(text)
        if _is_reference_title(text, config):
            in_references = True
            in_ack = False
            _set_paragraph_text(paragraph, "参考文献")
            apply_paragraph_rule(paragraph, style_config[STYLE_H1])
            paragraph.paragraph_format.space_before = Pt(0)
            continue
        if normalized == "致谢":
            in_ack = True
            in_references = False
            _set_paragraph_text(paragraph, "致谢")
            apply_paragraph_rule(paragraph, style_config[STYLE_H1])
            continue
        if _is_intro_title(text):
            _set_paragraph_text(paragraph, "引言")
            apply_paragraph_rule(paragraph, style_config[STYLE_H1])
            paragraph.paragraph_format.space_before = Pt(0)
            chapter_idx = 0
            section_idx = 0
            sub_idx = 0
            continue
        if normalized == "结语":
            _set_paragraph_text(paragraph, "结语")
            apply_paragraph_rule(paragraph, style_config[STYLE_H1])
            continue

        if in_references:
            apply_paragraph_rule(
                paragraph,
                {
                    "east_asia_font": "宋体",
                    "size_pt": 10.5,
                    "bold": False,
                    "line_spacing": 1.5,
                    "alignment": "left",
                    "hanging_indent_chars": 2,
                },
            )
            continue
        if in_ack:
            apply_paragraph_rule(paragraph, style_config[STYLE_BODY])
            continue

        if _looks_like_heading(paragraph):
            level = _heading_level_by_prefix(text)
            if level is None:
                style_key = _resolve_rule_key(paragraph)
                if style_key == STYLE_H2:
                    level = 2
                elif style_key == STYLE_H3:
                    level = 3
                else:
                    level = 1

            if level == 1:
                chapter_idx += 1
                section_idx = 0
                sub_idx = 0
                _set_paragraph_text(paragraph, _strip_number_prefix(text))
                _set_heading_numbering(paragraph, heading_num_id, 0)
                apply_paragraph_rule(paragraph, style_config[STYLE_H1])
            elif level == 2:
                if chapter_idx == 0:
                    chapter_idx = 1
                section_idx += 1
                sub_idx = 0
                _set_paragraph_text(paragraph, _strip_number_prefix(text))
                _set_heading_numbering(paragraph, heading_num_id, 1)
                apply_paragraph_rule(paragraph, style_config[STYLE_H2])
            else:
                if chapter_idx == 0:
                    chapter_idx = 1
                if section_idx == 0:
                    section_idx = 1
                sub_idx += 1
                _set_paragraph_text(paragraph, _strip_number_prefix(text))
                _set_heading_numbering(paragraph, heading_num_id, 2)
                apply_paragraph_rule(paragraph, style_config[STYLE_H3])
        else:
            apply_paragraph_rule(paragraph, style_config[STYLE_BODY])

    return start_idx


def apply_mvp_format(doc_path: Path, output_path: Path, config: dict) -> None:
    document = Document(str(doc_path))
    start_idx = apply_required_format(document, config)
    _enforce_structure(document, config, start_idx)
    apply_page_number_from_intro(document, start_idx, config.get("page_number", {}))
    document.save(str(output_path))


def iter_docx_files(root_dir: Path) -> list[Path]:
    return [
        p
        for p in root_dir.rglob("*.docx")
        if p.is_file() and not p.name.startswith("~$") and not p.stem.endswith(OUTPUT_SUFFIX)
    ]


def build_output_path(src: Path) -> Path:
    """Create an output path in the same directory with suffix."""
    return src.with_name(f"{src.stem}{OUTPUT_SUFFIX}{src.suffix}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Format .docx files in a directory and generate copies."
    )
    parser.add_argument(
        "directory",
        nargs="?",
        type=Path,
        help="Target directory containing .docx files (optional, default from config)",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=DEFAULT_CONFIG_PATH,
        help=f"Path to config json (default: {DEFAULT_CONFIG_PATH.name})",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        config = load_config(args.config)
    except Exception as exc:
        print(f"[ERROR] Failed to load config: {exc}")
        return 1

    config_base_dir = args.config.resolve().parent
    configured_directory = Path(config.get("directory", "."))
    if not configured_directory.is_absolute():
        configured_directory = (config_base_dir / configured_directory).resolve()
    target_dir = args.directory or configured_directory
    target_dir = target_dir.resolve()

    if not target_dir.exists() or not target_dir.is_dir():
        print(f"[ERROR] Directory not found: {target_dir}")
        return 1

    docx_files = iter_docx_files(target_dir)
    if not docx_files:
        print(f"[INFO] No .docx files found in: {target_dir}")
        return 0

    for src in docx_files:
        out = build_output_path(src)
        apply_mvp_format(src, out, config)
        print(f"[OK] {src} -> {out}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
