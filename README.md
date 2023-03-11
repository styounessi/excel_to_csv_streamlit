# Streamlit Excel to CSV Converter
![screenshot](https://i.imgur.com/1DDEQvu.png)

This converter takes Excel files and converts them directly to CSV files without changing the columns or rows in any way. There is a limit to the size and quantity of files that can be processed so use cases will vary but simple operations and files will benefit from the elimination of manual conversion.

Please ensure that input files are single sheet Excel files, not multi-sheet, since CSVs do not have multiple sheets.

Please ensure that only `.xlsx` or `.xls` files are used, using any other file type will result in an error.

## Technologies Used
- [Pandas](https://pypi.org/project/pandas/)
- [Streamlit](https://pypi.org/project/streamlit/)
- [Pillow](https://pypi.org/project/Pillow/)
