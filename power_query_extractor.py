import base64
import binascii
import re
import zipfile
from io import BytesIO
from pathlib import Path
from tkinter import Tk, filedialog

from lxml import etree


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb"}
DATA_MASHUP_NAMESPACE = {"d": "http://schemas.microsoft.com/DataMashup"}


def _split_query_definitions(section_content: str) -> list[tuple[str, str]]:
    """Split a Power Query section into named query definitions."""
    segments = []
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

    queries = []
    for segment in segments:
        match = re.match(
            r'^\s*(?:(?://[^\r\n]*(?:\r?\n|$)|/\*.*?\*/)\s*)*'
            r'shared\s+(#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=',
            segment,
            re.DOTALL,
        )
        if not match:
            continue

        query_name = match.group(1)
        if query_name.startswith('#"'):
            query_name = query_name[2:-1].replace('""', '"')
        query_content = segment[match.end() :].strip()
        queries.append((query_name, query_content))

    return queries


def find_power_query_files(excel_path: str | Path) -> list[tuple[str, str]]:
    """Return each Power Query name and its M code from an Excel workbook."""
    workbook_path = Path(excel_path)
    if workbook_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        supported_formats = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(f"Unsupported file format. Use {supported_formats}.")

    power_query_files = []
    with zipfile.ZipFile(workbook_path, "r") as workbook:
        for file_info in workbook.infolist():
            file_name = file_info.filename
            if not (
                file_name.startswith("customXml/")
                and file_name.lower().endswith(".xml")
            ):
                continue

            with workbook.open(file_info) as xml_file:
                try:
                    root = etree.fromstring(xml_file.read())
                except etree.XMLSyntaxError:
                    continue

            data_mashup_elements = root.xpath(
                "//d:DataMashup", namespaces=DATA_MASHUP_NAMESPACE
            )
            if not data_mashup_elements or not data_mashup_elements[0].text:
                continue

            try:
                decoded_content = base64.b64decode(data_mashup_elements[0].text)
            except (binascii.Error, ValueError) as error:
                raise ValueError(
                    f"Invalid DataMashup data in {file_name}."
                ) from error

            zip_start = decoded_content.find(b"PK\x03\x04")
            zip_end = decoded_content.find(b"PK\x05\x06")
            if zip_start == -1 or zip_end == -1:
                continue

            embedded_zip = BytesIO(decoded_content[zip_start : zip_end + 22])
            with zipfile.ZipFile(embedded_zip, "r") as mashup:
                if "Formulas/Section1.m" in mashup.namelist():
                    section_content = mashup.read(
                        "Formulas/Section1.m"
                    ).decode("utf-8")
                    power_query_files.extend(
                        _split_query_definitions(section_content)
                    )

    return power_query_files


def select_excel_file() -> str:
    """Show a file dialog and return the selected workbook path."""
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


def save_power_query_files(
    queries: list[tuple[str, str]], output_dir: str | Path
) -> list[Path]:
    """Save extracted queries using their names as .pq filenames."""
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    saved_files = []

    for index, (query_name, query_content) in enumerate(queries, start=1):
        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", query_name)
        safe_name = safe_name.rstrip(" .").strip()
        if not safe_name:
            safe_name = f"PowerQuery_{index}"
        reserved_names = {
            "CON",
            "PRN",
            "AUX",
            "NUL",
            *(f"COM{number}" for number in range(1, 10)),
            *(f"LPT{number}" for number in range(1, 10)),
        }
        if safe_name.upper().split(".")[0] in reserved_names:
            safe_name = f"_{safe_name}"

        output_path = destination / f"{safe_name}.pq"
        output_path.write_text(query_content, encoding="utf-8")
        saved_files.append(output_path)

    return saved_files


def main() -> None:
    excel_path = select_excel_file()
    if not excel_path:
        print("Ekstraksi dibatalkan: tidak ada file yang dipilih.")
        return

    queries = find_power_query_files(excel_path)
    output_dir = Path(excel_path).parent / Path(excel_path).stem
    saved_files = save_power_query_files(queries, output_dir)

    print(f"File sumber: {Path(excel_path).name}")
    print(f"Jumlah query: {len(saved_files)}")
    print("File hasil ekstrak:")
    if saved_files:
        for file_path in saved_files:
            print(f"- {file_path.name}")
    else:
        print("- Tidak ada query yang ditemukan.")


if __name__ == "__main__":
    main()
