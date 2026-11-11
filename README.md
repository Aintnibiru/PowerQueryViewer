# PowerQueryViewer

The `power_query_extractor.py` script is a Python utility designed for data analysts, developers, and anyone involved in data transformation and analysis using Excel's Power Query feature. It automates the extraction and decoding of Power Query M code from Excel files (.xlsx, .xlsm, .xlsb), bypassing the need to manually open and navigate through Excel documents.

## Features

- **Supports Multiple Excel Formats**: Works seamlessly with `.xlsx`, `.xlsm`, and `.xlsb` file formats.
- **Efficient Extraction and Decoding**: Automatically navigates through Excel's internal ZIP and XML structures to extract and decode Power Query M code from `Formulas/Section1.m`.
- **Full Syntax Support**: Correctly handles Power Query parameters (`[IsParameterQuery=true]`), member attributes, quoted identifiers (`#"..."`), comments, and Unicode query names.
- **CLI & GUI Modes**: Run with a graphical file picker dialog or pass file paths directly via command-line arguments for automated pipelines.
- **Zero-Dependency Ready**: Operates with Python's standard library out of the box, with optional `lxml` support for faster XML processing.
- **Safe Output Sanitization**: Cross-platform file naming with automatic collision resolution to avoid overwriting queries with similar or colliding names.

## Getting Started

### Prerequisites

Python 3.10 or higher.

### Installation

1. Clone the repository to your local machine:
   ```bash
   git clone https://github.com/jamesdesantiago/PowerQueryViewer.git
   ```
2. Navigate to the cloned repository directory.
3. *(Optional)* Install `lxml` for accelerated XML parsing:
   ```powershell
   python -m pip install lxml
   ```

## Usage

### Interactive GUI Mode

Run the script without arguments to open a file selection dialog:

```powershell
python power_query_extractor.py
```

### Command-Line (CLI) Mode

Pass the Excel file directly:

```powershell
python power_query_extractor.py "path/to/workbook.xlsx"
```

Specify a custom output directory:

```powershell
python power_query_extractor.py "path/to/workbook.xlsx" -o "path/to/output_dir"
```

### Output

Each extracted query is saved as an individual `.pq` file containing the query expression (such as `let ... in ...`). By default, files are saved in a subfolder named after the workbook in the same directory as the source file.

## Testing

Run the automated test suite with standard Python:

```powershell
python -m unittest -v test_power_query_extractor.py
```

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the Apache 2.0 License. See `LICENSE` for more information.

## Contact

James De Santiago - james.desantiago@outlook.com

Project Link: [https://github.com/jamesdesantiago/PowerQueryViewer](https://github.com/jamesdesantiago/PowerQueryViewer)
