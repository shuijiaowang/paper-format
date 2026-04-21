from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


STYLE_BODY = "Normal"
STYLE_H1 = "Heading 1"
STYLE_H2 = "Heading 2"
STYLE_H3 = "Heading 3"

OUTPUT_SUFFIX = "_formatted"
DEFAULT_CONFIG_PATH = Path(__file__).with_name("config.json")


def ensure_paragraph_style(document: Document, style_name: str) -> None:
    """Create paragraph style if missing."""
    if style_name in document.styles:
        return
    document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)


def set_style_font(
    document: Document,
    style_name: str,
    *,
    east_asia_font: str,
    latin_font: str = "Times New Roman",
    size_pt: float,
    bold: bool = False,
    first_line_indent_chars: int = 0,
    line_spacing: float | None = None,
) -> None:
    """Set font attributes for a paragraph style."""
    ensure_paragraph_style(document, style_name)
    style = document.styles[style_name]
    paragraph_format = style.paragraph_format
    paragraph_format.first_line_indent = Pt(size_pt * first_line_indent_chars)
    if line_spacing is not None:
        paragraph_format.line_spacing = line_spacing

    font = style.font
    font.name = latin_font
    font.bold = bold
    font.size = Pt(size_pt)
    font.element.rPr.rFonts.set(qn("w:eastAsia"), east_asia_font)


def apply_run_fonts(document: Document, style_config: dict) -> None:
    """Apply run-level fonts so English/digits stay Times New Roman."""
    style_font_map = {
        STYLE_BODY: style_config["body"]["east_asia_font"],
        STYLE_H1: style_config["h1"]["east_asia_font"],
        STYLE_H2: style_config["h2"]["east_asia_font"],
        STYLE_H3: style_config["h3"]["east_asia_font"],
    }

    for paragraph in document.paragraphs:
        east_asia_font = style_font_map.get(paragraph.style.name, style_config["body"]["east_asia_font"])
        for run in paragraph.runs:
            run.font.name = "Times New Roman"
            if run._element.rPr is None:
                run._element.get_or_add_rPr()
            r_fonts = run._element.rPr.rFonts
            if r_fonts is None:
                r_fonts = OxmlElement("w:rFonts")
                run._element.rPr.append(r_fonts)
            r_fonts.set(qn("w:ascii"), "Times New Roman")
            r_fonts.set(qn("w:hAnsi"), "Times New Roman")
            r_fonts.set(qn("w:cs"), "Times New Roman")
            r_fonts.set(qn("w:eastAsia"), east_asia_font)


def load_config(config_path: Path) -> dict:
    """Load formatter config from json."""
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")
    return json.loads(config_path.read_text(encoding="utf-8"))


def apply_page_config(document: Document, page_config: dict) -> None:
    """Apply simple page-level config to all sections."""
    for section in document.sections:
        if "top_margin_cm" in page_config:
            section.top_margin = Cm(page_config["top_margin_cm"])
        elif "top_margin_pt" in page_config:
            section.top_margin = Pt(page_config["top_margin_pt"])

        if "bottom_margin_cm" in page_config:
            section.bottom_margin = Cm(page_config["bottom_margin_cm"])
        elif "bottom_margin_pt" in page_config:
            section.bottom_margin = Pt(page_config["bottom_margin_pt"])

        if "left_margin_cm" in page_config:
            section.left_margin = Cm(page_config["left_margin_cm"])
        elif "left_margin_pt" in page_config:
            section.left_margin = Pt(page_config["left_margin_pt"])

        if "right_margin_cm" in page_config:
            section.right_margin = Cm(page_config["right_margin_cm"])
        elif "right_margin_pt" in page_config:
            section.right_margin = Pt(page_config["right_margin_pt"])

        if "header_distance_cm" in page_config:
            section.header_distance = Cm(page_config["header_distance_cm"])
        if "footer_distance_cm" in page_config:
            section.footer_distance = Cm(page_config["footer_distance_cm"])


def _set_page_number_start(section, start: int) -> None:
    """Set page numbering start in section properties."""
    sect_pr = section._sectPr
    pg_num_type = sect_pr.find(qn("w:pgNumType"))
    if pg_num_type is None:
        pg_num_type = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num_type)
    pg_num_type.set(qn("w:start"), str(start))


def _add_page_number_run(paragraph) -> None:
    """Insert PAGE field into a paragraph."""
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


def _mark_update_fields_on_open(document: Document) -> None:
    """Ask Word to update fields (including TOC) when opening."""
    settings = document.settings.element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")


def apply_header_footer_config(document: Document, page_config: dict) -> None:
    """Apply header/footer text and page number display."""
    header_text = page_config.get("header_text", "")

    for section in document.sections:
        section.different_first_page_header_footer = True
        section.odd_and_even_pages_header_footer = False
        _set_page_number_start(section, page_config.get("page_number_start", 0))

        header = section.header
        header.is_linked_to_previous = False
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.text = header_text
        header_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        for run in header_para.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(page_config.get("header_font_size_pt", 10.5))
            r_pr = run._element.get_or_add_rPr()
            r_fonts = r_pr.rFonts
            if r_fonts is None:
                r_fonts = OxmlElement("w:rFonts")
                r_pr.append(r_fonts)
            r_fonts.set(qn("w:eastAsia"), page_config.get("header_east_asia_font", "宋体"))

        first_page_header = section.first_page_header
        first_page_header.is_linked_to_previous = False
        if first_page_header.paragraphs:
            first_page_header.paragraphs[0].text = ""

        footer = section.footer
        footer.is_linked_to_previous = False
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_para.text = ""
        footer_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        _add_page_number_run(footer_para)
        for run in footer_para.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(page_config.get("page_number_font_size_pt", 10.5))
            r_pr = run._element.get_or_add_rPr()
            r_fonts = r_pr.rFonts
            if r_fonts is None:
                r_fonts = OxmlElement("w:rFonts")
                r_pr.append(r_fonts)
            r_fonts.set(qn("w:eastAsia"), page_config.get("page_number_east_asia_font", "宋体"))

        first_page_footer = section.first_page_footer
        first_page_footer.is_linked_to_previous = False
        if first_page_footer.paragraphs:
            first_page_footer.paragraphs[0].text = ""


def ensure_toc_field(document: Document) -> None:
    """Ensure a TOC field exists after a '目录' paragraph if present."""
    for paragraph in document.paragraphs:
        xml = paragraph._element.xml
        if "TOC" in xml and "fldCharType" in xml:
            return

    for paragraph in document.paragraphs:
        if paragraph.text.strip() == "目录":
            toc_para = paragraph.insert_paragraph_before("")
            fld_begin = OxmlElement("w:fldChar")
            fld_begin.set(qn("w:fldCharType"), "begin")

            instr_text = OxmlElement("w:instrText")
            instr_text.set(qn("xml:space"), "preserve")
            instr_text.text = r'TOC \o "1-3" \h \z \u'

            fld_separate = OxmlElement("w:fldChar")
            fld_separate.set(qn("w:fldCharType"), "separate")

            fld_end = OxmlElement("w:fldChar")
            fld_end.set(qn("w:fldCharType"), "end")

            run = toc_para.add_run()
            run._element.append(fld_begin)
            run._element.append(instr_text)
            run._element.append(fld_separate)
            run._element.append(fld_end)
            return


def apply_mvp_format(doc_path: Path, output_path: Path, config: dict) -> None:
    """Apply minimal formatting rules and save a copy."""
    document = Document(str(doc_path))
    style_config = config["styles"]
    page_config = config.get("page", {})

    apply_page_config(document, page_config)

    # 正文：小四（12pt），首行缩进两个字符
    set_style_font(
        document,
        STYLE_BODY,
        east_asia_font=style_config["body"]["east_asia_font"],
        size_pt=style_config["body"]["size_pt"],
        bold=style_config["body"]["bold"],
        first_line_indent_chars=style_config["body"]["first_line_indent_chars"],
        line_spacing=style_config["body"].get("line_spacing", 1.5),
    )

    # 一级标题：小三（15pt）黑体
    set_style_font(
        document,
        STYLE_H1,
        east_asia_font=style_config["h1"]["east_asia_font"],
        size_pt=style_config["h1"]["size_pt"],
        bold=style_config["h1"]["bold"],
    )

    # 二级标题：四号（14pt）黑体
    set_style_font(
        document,
        STYLE_H2,
        east_asia_font=style_config["h2"]["east_asia_font"],
        size_pt=style_config["h2"]["size_pt"],
        bold=style_config["h2"]["bold"],
    )

    # 三级标题：小四（12pt）宋体加粗
    set_style_font(
        document,
        STYLE_H3,
        east_asia_font=style_config["h3"]["east_asia_font"],
        size_pt=style_config["h3"]["size_pt"],
        bold=style_config["h3"]["bold"],
    )

    apply_run_fonts(document, style_config)
    apply_header_footer_config(document, page_config)
    ensure_toc_field(document)
    _mark_update_fields_on_open(document)

    document.save(str(output_path))


def iter_docx_files(root_dir: Path) -> list[Path]:
    """Find .docx files recursively and skip temp files."""
    return [
        p
        for p in root_dir.rglob("*.docx")
        if p.is_file() and not p.name.startswith("~$")
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

    configured_directory = Path(config.get("directory", "."))
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
