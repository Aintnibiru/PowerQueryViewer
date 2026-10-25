# PowerQueryViewer

The `power_query_extractor.py` script is a powerful Python utility designed for data analysts, developers, and anyone involved in data transformation and analysis using Excel's Power Query feature. It automates the extraction and decoding of Power Query M code from Excel files (.xlsx, .xlsm, .xlsb), bypassing the need to manually open and navigate through Excel documents.

## Features

- **Supports Multiple Excel Formats**: Works seamlessly with .xlsx, .xlsm, and .xlsb file formats.
- **Efficient Extraction and Decoding**: Automatically navigates through Excel's internal ZIP and XML structures to extract and decode Power Query M code, specifically targeting content within `Formulas/Section1.m`.
- **Open File Dialog**: Select the Excel workbook to inspect from a graphical file picker.

## Getting Started

### Prerequisites

Ensure you have Python 3.x installed on your system. If you don't have Python installed, you can download it from the [official Python website](https://www.python.org/downloads/).

### Installation

1. Clone the repository to your local machine:
git clone https://github.com/jamesdesantiago/PowerQueryViewer.git


2. Navigate to the cloned repository directory.
3. Install the required XML parser:

   ```powershell
   python -m pip install lxml
   ```

## Usage

Run the script from a terminal:

```powershell
python power_query_extractor.py
```

Choose an `.xlsx`, `.xlsm`, or `.xlsb` workbook in the file dialog. Each extracted query is saved as a `.pq` file named after the query in a new folder named after the workbook, next to the selected workbook. The `.pq` file contains only the query expression (for example, `let ... in ...`), without the `section Section1;` header or `shared QueryName =` declaration. For example, a query named `Monthly Sales` in `Sales.xlsx` is saved as `Sales/Monthly Sales.pq`. If that file already exists, it is overwritten with the newly extracted query. The console displays a short summary with the source workbook name, number of queries, and names of the extracted `.pq` files. XML files that cannot be parsed are skipped without displaying individual warnings. Closing the dialog without selecting a file cancels the operation.

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
