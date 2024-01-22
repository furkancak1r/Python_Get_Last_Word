
# Excel City Name Extractor

## Overview
This script is designed to extract city names from a specified column in an Excel workbook and write them into another column. It's particularly useful for processing large datasets where city names are embedded within strings.

## Prerequisites
- Python 3
- openpyxl library

To install openpyxl, run:
```
pip install openpyxl
```

## Usage
The `extract_and_write` function is called with two arguments: the file path of the Excel workbook and the name of the sheet to be processed.

```python
extract_and_write("path/to/your/excel/file.xlsx", "SheetName")
```

## Functionality
- Iterates through rows 2 to 1297 in the specified sheet.
- Searches for city names in column E.
- Writes the found city name into column F.
- Saves the changes in the Excel workbook.

## Limitations
- The script currently supports a predefined list of Turkish cities.
- Only processes rows 2 to 1297.

## Contributing
Feel free to fork this repository and submit pull requests for enhancements.

## License
This project is open-source and available under the [MIT License](LICENSE).
