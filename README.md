# PowerQueryViewer

The `power_query_extractor.py` script is a powerful Python utility designed for data analysts, developers, and anyone involved in data transformation and analysis using Excel's Power Query feature. It automates the extraction and decoding of Power Query M code from Excel files (.xlsx, .xlsm, .xlsb), bypassing the need to manually open and navigate through Excel documents.

## Features

- **Supports Excel Workbooks**: Works with `.xlsx` and `.xlsm` files. Legacy binary `.xlsb` workbooks are rejected with a clear error, because they use the BIFF12 binary format and are not ZIP containers.
- **Efficient Extraction and Decoding**: Automatically navigates through Excel's internal ZIP and XML structures to extract and decode Power Query M code, targeting every `Formulas/*.m` part (usually `Formulas/Section1.m`).
- **Command-Line Interface**: Designed for ease of use, enabling users to quickly extract Power Query M code through a simple command-line instruction, printed to stdout or written to a file.

## Getting Started

### Prerequisites

Ensure you have Python 3.9 or newer installed on your system. If you don't have Python installed, you can download it from the [official Python website](https://www.python.org/downloads/).

### Installation

1. Clone the repository to your local machine:
git clone https://github.com/jamesdesantiago/PowerQueryViewer.git


2. Navigate to the cloned repository directory.
3. Create the environment and install dependencies:

```powershell
uv venv
uv sync
```

`lxml` is the single external dependency required by the script; `zipfile`, `base64`, `argparse` and `io` are part of the standard library. If you do not use uv, install it with `pip install lxml`. Note that environments created by uv do not include pip, so use `uv pip install lxml` inside them.

## Usage

Run the script from the command line, passing the workbook you want to inspect:

```powershell
python power_query_extractor.py path\to\workbook.xlsx
```

Or, once the project is installed, use the console script:

```powershell
power-query-extractor path\to\workbook.xlsx
```

The extracted M code is printed to stdout. Pass `--output queries.m` to write it to a file instead. The exit code is `0` on success, `1` on error, and `2` when the workbook contains no Power Query queries.

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

