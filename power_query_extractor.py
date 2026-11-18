import argparse
import base64
import sys
import zipfile
from io import BytesIO

from lxml import etree

# Signature of the local file header that starts every ZIP archive
LOCAL_HEADER_SIGNATURE = b'PK\x03\x04'
# Signature of the end of central directory record (EOCD) that terminates a ZIP archive
EOCD_SIGNATURE = b'PK\x05\x06'
# Fixed size of the EOCD record, excluding its optional trailing comment
EOCD_SIZE = 22
# Namespace used in the DataMashup elements of the customXml parts
NAMESPACE = {'d': 'http://schemas.microsoft.com/DataMashup'}

# .xlsb uses the binary BIFF12 format and is not a ZIP container
SUPPORTED_EXTENSIONS = ('.xlsx', '.xlsm')


def read_mashup_archives(zip_ref):
    # Yield the embedded ZIP archives found in the DataMashup elements of an Excel file
    for file_info in zip_ref.infolist():
        # Check if the file is located within the customXml directory and has an xml extension
        if not (file_info.filename.startswith('customXml/') and file_info.filename.endswith('.xml')):
            continue

        # Open and read the content of the XML file
        with zip_ref.open(file_info) as xml_file:
            xml_content_bytes = xml_file.read()

        try:
            # Parse the XML content
            root = etree.fromstring(xml_content_bytes)
        except etree.XMLSyntaxError as e:
            # Handle any XML parsing errors
            print(f"XML parsing error in {file_info.filename}: {e}", file=sys.stderr)
            continue

        # Search for DataMashup elements within the XML document
        data_mashup_elements = root.xpath('//d:DataMashup', namespaces=NAMESPACE)
        if not data_mashup_elements:
            print(f"DataMashup content not found in {file_info.filename}.", file=sys.stderr)
            continue

        # Decode the base64 content of the DataMashup element
        decoded_content = base64.b64decode(data_mashup_elements[0].text or '')

        # Locate the embedded ZIP archive using its start and end signatures.
        # rfind is required for the EOCD record because the same byte sequence can
        # occur inside the compressed payload.
        zip_start = decoded_content.find(LOCAL_HEADER_SIGNATURE)
        zip_end = decoded_content.rfind(EOCD_SIGNATURE)
        if zip_start == -1 or zip_end == -1:
            print(f"ZIP archive start or end signature not found in {file_info.filename}.", file=sys.stderr)
            continue

        # Include the EOCD record and its optional trailing comment
        comment_length = 0
        if zip_end + EOCD_SIZE <= len(decoded_content):
            comment_length = int.from_bytes(decoded_content[zip_end + 20:zip_end + 22], 'little')

        yield BytesIO(decoded_content[zip_start:zip_end + EOCD_SIZE + comment_length])


def find_power_query_files(excel_path):
    # List to store the Power Query M code found within the Excel file
    power_query_files = []

    # Check if the provided file is an Excel file based on its extension
    if excel_path.lower().endswith('.xlsb'):
        raise ValueError(".xlsb files use the binary BIFF12 format and cannot be read as a ZIP container. "
                         "Please use .xlsx or .xlsm files.")
    if not excel_path.lower().endswith(SUPPORTED_EXTENSIONS):
        raise ValueError("Unsupported file format. Please use .xlsx or .xlsm files.")

    # Open the Excel file as a zip archive
    with zipfile.ZipFile(excel_path, 'r') as zip_ref:
        for zip_data in read_mashup_archives(zip_ref):
            with zipfile.ZipFile(zip_data) as archive:
                # The Power Query formulas live in the Formulas/*.m parts (usually Section1.m)
                for name in archive.namelist():
                    if name.startswith('Formulas/') and name.endswith('.m'):
                        power_query_files.append(archive.read(name).decode('utf-8'))

    # Return the Power Query M code found in the workbook
    return power_query_files


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog='power-query-extractor',
        description='Extract and decode Power Query M code from an Excel workbook (.xlsx or .xlsm).',
    )
    parser.add_argument('excel_path', help='path to the .xlsx or .xlsm workbook to inspect')
    parser.add_argument('-o', '--output', help='write the extracted M code to this file instead of stdout')
    args = parser.parse_args(argv)

    try:
        sections = find_power_query_files(args.excel_path)
    except (ValueError, OSError, zipfile.BadZipFile) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    if not sections:
        print("No Power Query M code found in the workbook.", file=sys.stderr)
        return 2

    m_code = '\n\n'.join(sections)
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as output_file:
            output_file.write(m_code)
        print(f"Wrote {len(sections)} section(s) to {args.output}")
    else:
        print(m_code)
    return 0


if __name__ == '__main__':
    sys.exit(main())
