from __future__ import annotations

import argparse
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def read_xml_from_docx(docx_path: Path, inner_path: str) -> ET.Element | None:
    with ZipFile(docx_path, "r") as zf:
        try:
            data = zf.read(inner_path)
        except KeyError:
            return None
    return ET.fromstring(data)


def parse_comments(docx_path: Path) -> dict[str, str]:
    comments_root = read_xml_from_docx(docx_path, "word/comments.xml")
    if comments_root is None:
        return {}

    comments: dict[str, str] = {}
    for comment in comments_root.findall("w:comment", NS):
        comment_id = comment.attrib.get(w_tag("id"))
        if comment_id is None:
            continue

        parts: list[str] = []
        for paragraph in comment.findall(".//w:p", NS):
            text = extract_text_from_paragraph(paragraph).strip()
            if text:
                parts.append(text)
        comments[comment_id] = "\n".join(parts).strip()
    return comments


def extract_text_from_paragraph(paragraph: ET.Element) -> str:
    chunks: list[str] = []

    for node in paragraph.iter():
        if node.tag == w_tag("t"):
            chunks.append(node.text or "")
        elif node.tag == w_tag("tab"):
            chunks.append("\t")
        elif node.tag in (w_tag("br"), w_tag("cr")):
            chunks.append("\n")

    return "".join(chunks).strip()


def get_heading_level(paragraph: ET.Element) -> int | None:
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        return None

    p_style = p_pr.find("w:pStyle", NS)
    if p_style is None:
        return None

    style_val = p_style.attrib.get(w_tag("val"), "")
    lowered = style_val.lower()
    if lowered.startswith("heading"):
        suffix = lowered.replace("heading", "").strip()
        if suffix.isdigit():
            level = int(suffix)
            if 1 <= level <= 6:
                return level

    # Chinese localized style IDs sometimes look like "标题1"
    if style_val.startswith("标题"):
        suffix = style_val.replace("标题", "").strip()
        if suffix.isdigit():
            level = int(suffix)
            if 1 <= level <= 6:
                return level

    return None


def parse_document_blocks(docx_path: Path, comments_by_id: dict[str, str]) -> list[str]:
    document_root = read_xml_from_docx(docx_path, "word/document.xml")
    if document_root is None:
        raise ValueError("Invalid .docx: missing word/document.xml")

    body = document_root.find("w:body", NS)
    if body is None:
        return []

    blocks: list[str] = []
    paragraph_comments: dict[int, list[str]] = defaultdict(list)

    for index, paragraph in enumerate(body.findall("w:p", NS)):
        text = extract_text_from_paragraph(paragraph)
        if not text:
            continue

        heading_level = get_heading_level(paragraph)
        if heading_level is not None:
            blocks.append(f'{"#" * heading_level} {text}')
        else:
            blocks.append(text)

        comment_ids: list[str] = []
        for ref in paragraph.findall(".//w:commentReference", NS):
            comment_id = ref.attrib.get(w_tag("id"))
            if comment_id:
                comment_ids.append(comment_id)
        for start in paragraph.findall(".//w:commentRangeStart", NS):
            comment_id = start.attrib.get(w_tag("id"))
            if comment_id:
                comment_ids.append(comment_id)

        seen = set()
        unique_ids = []
        for comment_id in comment_ids:
            if comment_id in seen:
                continue
            seen.add(comment_id)
            unique_ids.append(comment_id)

        for comment_id in unique_ids:
            comment_text = comments_by_id.get(comment_id)
            if comment_text:
                paragraph_comments[index].append(comment_text)

        if paragraph_comments.get(index):
            for i, comment_text in enumerate(paragraph_comments[index], start=1):
                blocks.append(f"> 批注{i}: {comment_text}")

        blocks.append("")

    return blocks


def convert_docx_to_markdown(docx_path: Path, output_path: Path) -> None:
    comments_by_id = parse_comments(docx_path)
    lines = parse_document_blocks(docx_path, comments_by_id)
    markdown = "\n".join(lines).rstrip() + "\n"
    output_path.write_text(markdown, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert .docx to .md and keep Word comments."
    )
    parser.add_argument("input_docx", type=Path, help="Input .docx file path")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output .md path (default: same name as input)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_docx = args.input_docx.resolve()
    if not input_docx.exists() or input_docx.suffix.lower() != ".docx":
        print(f"[ERROR] Invalid input .docx: {input_docx}")
        return 1

    output_path = args.output.resolve() if args.output else input_docx.with_suffix(".md")
    convert_docx_to_markdown(input_docx, output_path)
    print(f"[OK] {input_docx} -> {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
