"""Unpack PPTX files for editing.

Extracts the ZIP archive and pretty-prints XML files.

Usage:
    python 1_unpack_pptx.py <pptx_file> <output_dir>

Example:
    python 1_unpack_pptx.py presentation.pptx unpacked/
"""

import argparse
import sys
import zipfile
from pathlib import Path

import defusedxml.minidom


SMART_QUOTE_REPLACEMENTS = {
    "\u201c": "&#x201C;",
    "\u201d": "&#x201D;",
    "\u2018": "&#x2018;",
    "\u2019": "&#x2019;",
}


def unpack(
    input_file: str,
    output_directory: str,
    allow_already_translated: bool = False,
    direction: str = "en2ja",
) -> tuple[None, str, bool]:
    """Unpack a PPTX. Returns (None, message, is_error)."""
    input_path = Path(input_file)
    output_path = Path(output_directory)

    if not input_path.exists():
        return None, f"Error: {input_file} does not exist", True

    if input_path.suffix.lower() != ".pptx":
        return None, f"Error: {input_file} must be a .pptx file", True

    # Direction-aware suffix guard:
    #   en2ja: refuse _JA (already Japanese) and _EN (English output of ja2en)
    #          because double-translating either is almost certainly a mistake.
    #   ja2en: refuse only _EN (already English). _JA inputs are the normal
    #          Japanese source for an English translation.
    if not allow_already_translated:
        suffix_blocked: tuple[str, ...]
        if direction == "ja2en":
            suffix_blocked = ("_EN",)
        else:
            suffix_blocked = ("_JA", "_EN")
        if input_path.stem.endswith(suffix_blocked):
            return None, (
                f"Error: {input_path.name} appears to already be a translation "
                f"output (suffix '{input_path.stem[-3:]}') for direction "
                f"{direction!r}. Refusing to translate to avoid double "
                "translation. Pass --force-already-translated to override."
            ), True

    try:
        output_path.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(output_path)

        xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
        for xml_file in xml_files:
            _pretty_print_xml(xml_file)

        for xml_file in xml_files:
            _escape_smart_quotes(xml_file)

        return None, f"Unpacked {input_file} ({len(xml_files)} XML files)", False

    except zipfile.BadZipFile:
        return None, f"Error: {input_file} is not a valid PPTX file", True
    except Exception as e:
        return None, f"Error unpacking: {e}", True


def _pretty_print_xml(xml_file: Path) -> None:
    try:
        content = xml_file.read_text(encoding="utf-8")
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
    except Exception:
        pass


def _escape_smart_quotes(xml_file: Path) -> None:
    try:
        content = xml_file.read_text(encoding="utf-8")
        for char, entity in SMART_QUOTE_REPLACEMENTS.items():
            content = content.replace(char, entity)
        xml_file.write_text(content, encoding="utf-8")
    except Exception:
        pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Unpack a PPTX file for editing")
    parser.add_argument("input_file", help="PPTX file to unpack")
    parser.add_argument("output_directory", help="Output directory")
    parser.add_argument(
        "--direction",
        choices=("en2ja", "ja2en"),
        default="en2ja",
        help=(
            "Translation direction this unpack is preparing for. en2ja "
            "(default) rejects _JA and _EN inputs; ja2en rejects only _EN."
        ),
    )
    parser.add_argument(
        "--force-already-translated",
        action="store_true",
        help=(
            "Allow unpacking files whose basename ends with _JA or _EN. "
            "By default such files are rejected to avoid double translation."
        ),
    )
    args = parser.parse_args()

    _, message, is_error = unpack(
        args.input_file,
        args.output_directory,
        allow_already_translated=args.force_already_translated,
        direction=args.direction,
    )
    print(message, file=sys.stderr if is_error else sys.stdout)

    if is_error:
        sys.exit(1)
