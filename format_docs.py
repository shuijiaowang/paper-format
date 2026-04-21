from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

OUTPUT_SUFFIX = "_formatted"
DEFAULT_CONFIG_PATH = Path(__file__).with_name("config.json")
STYLE_BODY = "body"
STYLE_H1 = "h1"
STYLE_H2 = "h2"
STYLE_H3 = "h3"


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


def apply_paragraph_rule(paragraph, rule: dict[str, float | int | str | bool]) -> None:
    pf = paragraph.paragraph_format
    size_pt = float(rule.get("size_pt", 12))

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


def apply_page_number(document: Document, page_number_config: dict) -> None:
    for section in document.sections:
        section.different_first_page_header_footer = False
        section.odd_and_even_pages_header_footer = False
        _set_page_number_start(section, int(page_number_config.get("start", 1)))

        footer = section.footer
        footer.is_linked_to_previous = False
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


def _contains_page_break(paragraph) -> bool:
    for run in paragraph.runs:
        if "<w:br" in run._element.xml and 'w:type="page"' in run._element.xml:
            return True
    return False


def _build_page_map(document: Document) -> list[int]:
    page_no = 1
    page_map: list[int] = []
    for paragraph in document.paragraphs:
        page_map.append(page_no)
        if _contains_page_break(paragraph):
            page_no += 1
    return page_map


def _find_start_index(document: Document, page_map: list[int], processing_cfg: dict) -> int:
    intro_texts = {"引言", "引  言"}
    start_page = int(processing_cfg.get("start_page", 5))
    probe_pages = int(processing_cfg.get("intro_probe_pages", 3))
    max_probe_page = start_page + max(probe_pages - 1, 0)

    for idx, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if text in intro_texts and start_page <= page_map[idx] <= max_probe_page:
            return idx

    for idx, paragraph in enumerate(document.paragraphs):
        if paragraph.text.strip() in intro_texts:
            return idx

    for idx, page in enumerate(page_map):
        if page >= start_page and document.paragraphs[idx].text.strip():
            return idx
    return 0


def _is_figure_caption(text: str) -> bool:
    return bool(re.match(r"^图\d+", text))


def _resolve_rule_key(paragraph) -> str:
    style_name = (paragraph.style.name or "").lower()
    if "heading 1" in style_name or "标题 1" in style_name:
        return STYLE_H1
    if "heading 2" in style_name or "标题 2" in style_name:
        return STYLE_H2
    if "heading 3" in style_name or "标题 3" in style_name:
        return STYLE_H3
    return STYLE_BODY


def apply_required_format(document: Document, config: dict) -> None:
    style_config = config["styles"]
    page_map = _build_page_map(document)
    start_idx = _find_start_index(document, page_map, config.get("processing", {}))
    last_page = max(page_map) if page_map else 1

    in_references = False
    for idx, paragraph in enumerate(document.paragraphs):
        if idx < start_idx:
            continue
        text = paragraph.text.strip()
        if not text:
            continue
        if page_map[idx] >= last_page:
            continue
        if text.startswith("参考文献"):
            in_references = True
            continue
        if in_references:
            continue
        if _is_figure_caption(text):
            continue
        if paragraph._p.xpath(".//w:drawing"):
            continue
        apply_paragraph_rule(paragraph, style_config[_resolve_rule_key(paragraph)])


def apply_mvp_format(doc_path: Path, output_path: Path, config: dict) -> None:
    document = Document(str(doc_path))
    apply_required_format(document, config)
    apply_page_number(document, config.get("page_number", {}))
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
