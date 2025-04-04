# Excel VBA Scripts for Coordinators and Managers

## Overview

This repository contains a set of Excel VBA scripts designed to automate and simplify the management of data within a workbook. These scripts focus on summing "PAGO NETO" values across multiple sheets, converting numbers to their Spanish word representations, and handling the creation and management of coordinator and manager data.

### Key Scripts

1. **SumPagoNetoCoordinacion**

   - **Version**: 1.6.2
   - **Author**: Juan Pablo Garcia Murillo
   - **Date**: 04/01/2025
   - **Description**: This script calculates the total sum of "PAGO NETO" values from column "D" across a list of specified sheets, as indicated in column "P" of the triggering sheet. If the active sheet is not listed in column "P", it processes the active sheet separately. The total sum is stored in cell "J4" of the triggering sheet.

2. **SumPagoNetoGerencia**

   - **Version**: 1.6.2
   - **Author**: Juan Pablo Garcia Murillo
   - **Date**: 04/01/2025
   - **Description**: This script calculates the total sum of "PAGO NETO" values from the "D" column across all visible worksheets in the workbook. The sum is calculated by calling the function `SumPagoNetoFromSheets`, which processes all visible sheets and stores the result in cell "J4" of the triggering sheet.

3. **Utils Module**
   - **Version**: 1.6.2
   - **Author**: Juan Pablo Garcia Murillo
   - **Date**: 04/01/2025
   - **Description**: This module contains several helper functions used throughout the workbook, including:
     - `SumPagoNetoFromSheets`: Calculates the sum of "PAGO NETO" values from multiple sheets.
     - `NumeroATexto`: Converts a numeric value into its Spanish word representation (e.g., "100" becomes "CIEN PESOS 00/100").
     - `ConvertirMenor1000`: Converts numbers less than 1000 into their Spanish word representation.
     - `IsInNewTabs`: Checks if a given sheet name exists in a collection of newly created tabs.
     - `IsRowEmpty`: Checks if a row is empty, ignoring the first column.

## How to Use

1. **Install the VBA Scripts**:

   - Open your Excel workbook and press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.
   - In the VBA editor, go to `Insert > Module` to add a new module.
   - Copy the code from each of the scripts above into the appropriate modules in the VBA editor.

2. **Run the Scripts**:

   - You can run the scripts by pressing `Alt + F8` in Excel, selecting the macro name, and clicking "Run".
   - The `SumPagoNetoCoordinacion` and `SumPagoNetoGerencia` scripts calculate and display the total "PAGO NETO" values in cell `J4` of the active sheet.
   - The `Utils` module provides various helper functions, which are used by other scripts.

3. **Trigger the Scripts**:
   - The scripts are designed to be triggered manually or through other Excel operations (e.g., button presses or sheet events).
   - Ensure that the sheet names and references are correctly set in the scripts to match your specific workbook setup.

## Function Descriptions

### `SumPagoNetoCoordinacion`

- **Purpose**: Sum "PAGO NETO" values from a list of sheets (specified in column "P" of the triggering sheet) and store the total in `J4`.
- **Parameters**: Uses sheet names from column "P" or processes the active sheet.
- **Output**: Total sum of "PAGO NETO" in cell `J4`.

### `SumPagoNetoGerencia`

- **Purpose**: Sum "PAGO NETO" values from all visible sheets in the workbook and store the result in `J4`.
- **Parameters**: None (it processes all visible sheets).
- **Output**: Total sum of "PAGO NETO" in cell `J4`.

### `Utils Module`

- **`SumPagoNetoFromSheets`**: Sums "PAGO NETO" values from multiple sheets.
- **`NumeroATexto`**: Converts a numeric value into its Spanish word representation.
- **`ConvertirMenor1000`**: Converts numbers less than 1000 into their Spanish word representation.
- **`IsInNewTabs`**: Checks if a sheet exists in a collection of newly created tabs.
- **`IsRowEmpty`**: Checks if a row is empty (ignoring the first column).

## Notes

- **Error Handling**: Each script includes basic error handling to ensure smooth execution, particularly in cases where sheet names or data are missing.
- **Customization**: You can modify the sheet references or cell locations to fit the structure of your specific workbook.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
