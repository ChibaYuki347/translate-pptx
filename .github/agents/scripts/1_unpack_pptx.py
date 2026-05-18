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
) -> tuple[None, str]:
    input_path = Path(input_file)
    output_path = Path(output_directory)

    if not input_path.exists():
        return None, f"Error: {input_file} does not exist"

    if input_path.suffix.lower() != ".pptx":
        return None, f"Error: {input_file} must be a .pptx file"

    if not allow_already_translated and input_path.stem.endswith(("_JA", "_EN")):
        return None, (
            f"Error: {input_path.name} appears to already be a translation output "
            f"(suffix '{input_path.stem[-3:]}'). Refusing to translate to avoid "
            "double translation. Pass --force-already-translated to override."
        )

    try:
        output_path.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(output_path)

        xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
        for xml_file in xml_files:
            _pretty_print_xml(xml_file)

        for xml_file in xml_files:
            _escape_smart_quotes(xml_file)

        return None, f"Unpacked {input_file} ({len(xml_files)} XML files)"

    except zipfile.BadZipFile:
        return None, f"Error: {input_file} is not a valid PPTX file"
    except Exception as e:
        return None, f"Error unpacking: {e}"


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
        "--force-already-translated",
        action="store_true",
        help=(
            "Allow unpacking files whose basename ends with _JA or _EN. "
            "By default such files are rejected to avoid double translation."
        ),
    )
    args = parser.parse_args()

    _, message = unpack(
        args.input_file,
        args.output_directory,
        allow_already_translated=args.force_already_translated,
    )
    print(message)

    if "Error" in message:
        sys.exit(1)
