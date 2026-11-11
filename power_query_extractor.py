import argparse
import base64
import binascii
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path

# Support standard library xml.etree with optional lxml acceleration
try:
    from lxml import etree

    HAS_LXML = True
    XML_PARSE_ERRORS: tuple[type[BaseException], ...] = (
        etree.XMLSyntaxError,
        ValueError,
        TypeError,
    )
except ImportError:
    import xml.etree.ElementTree as etree  # nosec B405 # type: ignore

    HAS_LXML = False
    XML_PARSE_ERRORS = (
        etree.ParseError,
        ValueError,
        TypeError,
    )

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb"}
DATA_MASHUP_NAMESPACE = {"d": "http://schemas.microsoft.com/DataMashup"}
WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{num}" for num in range(1, 10)),
    *(f"LPT{num}" for num in range(1, 10)),
}


def _split_query_definitions(section_content: str) -> list[tuple[str, str]]:
    """Split a Power Query section into named query definitions.

    Handles comments, nested brackets/records, string literals, section
    attributes, parameters, and quoted or unquoted query identifiers.
    """
    segments: list[str] = []
    segment_start = 0
    depth = 0
    quote = False
    line_comment = False
    block_comment = False
    index = 0

    while index < len(section_content):
        char = section_content[index]
        next_char = (
            section_content[index + 1] if index + 1 < len(section_content) else ""
        )

        if line_comment:
            if char in "\r\n":
                line_comment = False
        elif block_comment:
            if char == "*" and next_char == "/":
                block_comment = False
                index += 1
        elif quote:
            if char == '"':
                if next_char == '"':
                    index += 1
                else:
                    quote = False
        elif char == "/" and next_char == "/":
            line_comment = True
            index += 1
        elif char == "/" and next_char == "*":
            block_comment = True
            index += 1
        elif char == '"':
            quote = True
        elif char in "([{":
            depth += 1
        elif char in ")]}":
            depth = max(0, depth - 1)
        elif char == ";" and depth == 0:
            segments.append(section_content[segment_start:index].strip())
            segment_start = index + 1

        index += 1

    trailing_segment = section_content[segment_start:].strip()
    if trailing_segment:
        segments.append(trailing_segment)

    queries: list[tuple[str, str]] = []
    for segment in segments:
        if not segment:
            continue

        # Find the assignment '=' at top-level depth 0 (outside quotes and comments)
        eq_pos = -1
        seg_depth = 0
        seg_quote = False
        seg_line_comment = False
        seg_block_comment = False
        i = 0

        while i < len(segment):
            ch = segment[i]
            nxt_ch = segment[i + 1] if i + 1 < len(segment) else ""

            if seg_line_comment:
                if ch in "\r\n":
                    seg_line_comment = False
            elif seg_block_comment:
                if ch == "*" and nxt_ch == "/":
                    seg_block_comment = False
                    i += 1
            elif seg_quote:
                if ch == '"':
                    if nxt_ch == '"':
                        i += 1
                    else:
                        seg_quote = False
            elif ch == "/" and nxt_ch == "/":
                seg_line_comment = True
                i += 1
            elif ch == "/" and nxt_ch == "*":
                seg_block_comment = True
                i += 1
            elif ch == '"':
                seg_quote = True
            elif ch in "([{":
                seg_depth += 1
            elif ch in ")]}":
                seg_depth = max(0, seg_depth - 1)
            elif ch == "=" and seg_depth == 0:
                eq_pos = i
                break

            i += 1

        # A section declaration like 'section Section1;' has no '=' assignment
        if eq_pos == -1:
            continue

        header = segment[:eq_pos].strip()
        query_content = segment[eq_pos + 1 :].strip()

        # Extract query identifier right before '='
        # Identifiers can be #"quoted identifier" or standard identifier
        match = re.search(r'(#"(?:[^"]|"")*"|[^\W\d_]\w*)\s*$', header)
        if not match:
            continue

        raw_name = match.group(1)
        if raw_name.startswith('#"'):
            query_name = raw_name[2:-1].replace('""', '"')
        else:
            query_name = raw_name

        queries.append((query_name, query_content))

    return queries


def _extract_mashup_package(decoded_content: bytes) -> bytes | None:
    """Extract the embedded Package ZIP bytes from decoded DataMashup content."""
    # DataMashup format specifies a 4-byte little-endian PackageLength at offset 0
    if len(decoded_content) >= 8:
        pkg_len = int.from_bytes(decoded_content[:4], "little")
        if (
            decoded_content[4:8] == b"PK\x03\x04"
            and len(decoded_content) >= 4 + pkg_len
        ):
            return decoded_content[4 : 4 + pkg_len]

    # Fallback: scan for standard ZIP signatures
    zip_start = decoded_content.find(b"PK\x03\x04")
    if zip_start == -1:
        return None

    # Find the End of Central Directory (EOCD) signature
    zip_end = decoded_content.rfind(b"PK\x05\x06", zip_start)
    if zip_end == -1:
        return None

    # EOCD is 22 bytes + comment length (2 bytes at offset 20)
    comment_len = 0
    if len(decoded_content) >= zip_end + 22:
        comment_len = int.from_bytes(
            decoded_content[zip_end + 20 : zip_end + 22], "little"
        )

    end_offset = zip_end + 22 + comment_len
    return decoded_content[zip_start:end_offset]


def _extract_mashup_base64_from_xml(xml_bytes: bytes) -> str | None:
    """Parse XML and extract the DataMashup base64 content text."""
    try:
        root = etree.fromstring(xml_bytes)  # nosec B314
    except XML_PARSE_ERRORS:
        return None

    if HAS_LXML:
        elements = root.xpath(  # type: ignore
            "//d:DataMashup", namespaces=DATA_MASHUP_NAMESPACE
        )
    else:
        elements = root.findall(  # type: ignore
            ".//d:DataMashup", namespaces=DATA_MASHUP_NAMESPACE
        )

    if elements and elements[0].text:
        return elements[0].text.strip()
    return None


def find_power_query_files(excel_path: str | Path) -> list[tuple[str, str]]:
    """Return each Power Query name and its M code from an Excel workbook."""
    workbook_path = Path(excel_path)
    if not workbook_path.is_file():
        raise FileNotFoundError(f"File not found: {workbook_path}")

    if workbook_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        supported_formats = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(
            f"Unsupported file format '{workbook_path.suffix}'. Use {supported_formats}."
        )

    power_query_files: list[tuple[str, str]] = []
    try:
        with zipfile.ZipFile(workbook_path, "r") as workbook:
            for file_info in workbook.infolist():
                file_name = file_info.filename
                if not (
                    file_name.startswith("customXml/")
                    and file_name.lower().endswith(".xml")
                ):
                    continue

                with workbook.open(file_info) as xml_file:
                    mashup_b64 = _extract_mashup_base64_from_xml(xml_file.read())
                    if not mashup_b64:
                        continue

                try:
                    decoded_content = base64.b64decode(mashup_b64)
                except (binascii.Error, ValueError) as error:
                    raise ValueError(
                        f"Invalid DataMashup base64 data in {file_name}."
                    ) from error

                package_bytes = _extract_mashup_package(decoded_content)
                if not package_bytes:
                    continue

                try:
                    with zipfile.ZipFile(BytesIO(package_bytes), "r") as mashup:
                        if "Formulas/Section1.m" in mashup.namelist():
                            section_content = mashup.read(
                                "Formulas/Section1.m"
                            ).decode("utf-8-sig")
                            power_query_files.extend(
                                _split_query_definitions(section_content)
                            )
                except zipfile.BadZipFile:
                    continue
    except zipfile.BadZipFile as error:
        raise ValueError(
            f"The file '{workbook_path.name}' is not a valid or readable Excel archive."
        ) from error

    return power_query_files


def select_excel_file() -> str:
    """Show a file dialog and return the selected workbook path."""
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        try:
            return filedialog.askopenfilename(
                title="Select an Excel workbook",
                filetypes=(
                    ("Excel workbooks", "*.xlsx *.xlsm *.xlsb"),
                    ("All files", "*.*"),
                ),
            )
        finally:
            root.destroy()
    except (ImportError, RuntimeError, OSError) as error:
        print(
            f"Error opening graphical file dialog: {error}\n"
            "Please specify the Excel file path directly as a command-line argument:\n"
            "  python power_query_extractor.py <path_to_excel_file>",
            file=sys.stderr,
        )
        return ""


def sanitize_filename(name: str, index: int = 1) -> str:
    """Sanitize a query name to a valid cross-platform filename."""
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    safe_name = safe_name.rstrip(" .").strip()
    if not safe_name:
        safe_name = f"PowerQuery_{index}"
    if safe_name.upper().split(".")[0] in WINDOWS_RESERVED_NAMES:
        safe_name = f"_{safe_name}"
    return safe_name


def save_power_query_files(
    queries: list[tuple[str, str]], output_dir: str | Path
) -> list[Path]:
    """Save extracted queries using sanitized names as .pq filenames.

    Automatically handles duplicate or colliding query filenames without data loss.
    """
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    saved_files: list[Path] = []
    used_names: set[str] = set()

    for index, (query_name, query_content) in enumerate(queries, start=1):
        base_safe_name = sanitize_filename(query_name, index)

        candidate_name = base_safe_name
        counter = 1
        while candidate_name.casefold() in used_names:
            counter += 1
            candidate_name = f"{base_safe_name} ({counter})"

        used_names.add(candidate_name.casefold())

        output_path = destination / f"{candidate_name}.pq"
        output_path.write_text(query_content, encoding="utf-8")
        saved_files.append(output_path)

    return saved_files


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract and decode Power Query M code from Excel workbooks (.xlsx, .xlsm, .xlsb)."
    )
    parser.add_argument(
        "file",
        nargs="?",
        default=None,
        help="Path to the Excel workbook. If omitted, a file picker dialog will open.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=None,
        help="Directory to save extracted .pq files. Defaults to a folder named after the workbook.",
    )

    args = parser.parse_args()

    excel_path = args.file
    if not excel_path:
        excel_path = select_excel_file()
        if not excel_path:
            print("Ekstraksi dibatalkan: tidak ada file yang dipilih.")
            return

    file_path = Path(excel_path)
    if not file_path.exists():
        print(f"Error: File tidak ditemukan: {file_path}", file=sys.stderr)
        sys.exit(1)

    try:
        queries = find_power_query_files(file_path)
    except (ValueError, FileNotFoundError, OSError) as exc:
        print(f"Error saat mengekstrak query: {exc}", file=sys.stderr)
        sys.exit(1)

    if args.output_dir:
        output_dir = Path(args.output_dir)
    else:
        output_dir = file_path.parent / file_path.stem

    saved_files = save_power_query_files(queries, output_dir)

    print(f"File sumber: {file_path.name}")
    print(f"Folder tujuan: {output_dir}")
    print(f"Jumlah query: {len(saved_files)}")
    print("File hasil ekstrak:")
    if saved_files:
        for saved in saved_files:
            print(f"- {saved.name}")
    else:
        print("- Tidak ada query yang ditemukan.")


if __name__ == "__main__":
    main()
