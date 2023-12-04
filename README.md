# Book Store Data to Excel Converter

This Node.js script reads book store data from a JSON file and converts it into an Excel spreadsheet using the `exceljs` library.

## Prerequisites

Before running the script, ensure you have Node.js installed on your machine.

- [Node.js](https://nodejs.org/)

## Getting Started

1. Clone the repository or download the script files:

    ```bash
    git clone https://github.com/yourusername/bookstore-to-excel.git
    ```

2. Install the required dependencies:

    ```bash
    npm install
    ```

## Usage

1. Place your book store data in a JSON file named `DATA.json` in the project root.

2. Open a terminal and run the script:

    ```bash
    node index.js
    ```

3. The script will generate an Excel file named `BookStoreData.xlsx` in the project root.

## Customization

- Modify the `DATA.json` file to include your book store data.
- Customize the Excel file's structure in the `FlattenData` function within the `script.js` file.

## Notes

- Ensure that the required modules (`fs` and `exceljs`) are correctly installed.
- Review and customize the script to suit your specific data structure.
