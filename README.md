# Excel Automation Suite - Version 1.6.4

## Description

The **Excel Automation Suite** is a collection of VBA scripts designed to automate and streamline data management processes in Excel workbooks. This suite includes subroutines that create coordinator and promoter-specific tabs, manage data validation, perform complex data lookups, and handle visibility and sheet creation with error handling and efficient workflows. The suite aims to reduce redundancy, improve maintainability, and ensure data consistency across various tasks.

## Features

### 1. **CreateCoordinatorTabs** (Version 1.6.4)

- **Automates Tab Creation**: Creates a new tab for each unique coordinator by copying a template and renaming it to the coordinator’s name.
- **Data Population**: Populates each new tab with filtered data specific to the coordinator.
- **Common Value Copying**: Copies shared values (e.g., `razonSocial`, `periodoDelPagoDel`) into all new tabs.
- **Sorting & Filtering**: Sorts coordinator names and filters the data for each coordinator.
- **Visibility Management**: Manages sheet visibility by restoring the original visibility state of sheets after processing.
- **Error Handling**: Handles cases where no coordinators are found or no matching data exists.

### 2. **CreatePromotorTabs** (Version 1.6.4)

- **Automates Promotor Tab Creation**: Similar to `CreateCoordinatorTabs`, this subroutine creates tabs for promotors associated with a specific coordinator.
- **Data Lookup and Transfer**: Populates the new tabs with filtered data for each promoter, ensuring the data is correctly mapped from the source sheet.
- **Template Consistency**: Ensures that new tabs follow the same format as the template sheet, maintaining consistency across the workbook.
- **Visibility Management**: Ensures proper handling of sheet visibility and tab creation.

### 3. **CreateBaseSalaryTabsIfMissing**

- **Base Salary Tab Creation**: Creates salary tabs for each promoter that is listed in the "Sueldos_Base" table under a specific coordinator.
- **Table Filtering**: Filters the data for each promotor and creates a new tab accordingly.

### 4. **Dynamic Data Validation Setup**

- **COORDINADOR Validation**: Applies dynamic data validation to the "COORDINADOR" column using values from the "Gerentes" and "Coordinadores" tables.
- **PROMOTOR Validation**: Applies dynamic validation to the "PROMOTOR" column based on the selected coordinator, ensuring that only relevant promoters are available for selection.

### 5. **Utility Functions in `mod_UtilsModule`**

- **Reusable Functions**: Includes various utility functions for general data management:
  - **Summing Values**: Sum values from specific columns across sheets.
  - **Number to Spanish Words**: Converts numbers to their Spanish word equivalents.
  - **Sheet Creation and Naming**: Handles creating and naming sheets from templates.
  - **Data Lookup**: Performs lookups to retrieve specific data from tables.
  - **Filtering and Sorting**: Provides functions to sort and filter tables easily.
  - **Table Population**: Automates the process of populating tables with filtered data.
  - **Sheet Existence Check**: Checks if a sheet exists before attempting to create it.
  - **Row Empty Check**: Determines if a row is empty, excluding certain columns.

## Requirements

- **Excel**: The suite is designed to run in Microsoft Excel.
- **VBA**: The suite is written in Visual Basic for Applications (VBA) and should be placed in the VBA editor of your Excel workbook.

## Setup Instructions

1. **Open Excel Workbook**: Open the Excel workbook where you want to run the scripts.
2. **Access VBA Editor**: Press `Alt + F11` to open the Visual Basic for Applications editor.
3. **Insert the Code**:
   - Insert each subroutine (e.g., `CreateCoordinatorTabs`, `CreatePromotorTabs`) into separate or shared modules in the workbook.
   - Include the utility functions in a module like `mod_UtilsModule`.
4. **Run the Scripts**: You can run individual subroutines like `CreateCoordinatorTabs` or link them to buttons or triggers in Excel.

## Parameters

- **None for Most Scripts**: These subroutines are generally designed without requiring input parameters. They operate based on the active workbook and existing data.
- **Dynamic Validation Parameters**: Validation setup functions may use the values already existing in specific columns and tables in the workbook.

## Returns

- **None**: The subroutines modify the workbook by creating tabs, updating data, and managing visibility, but they do not return any values.

## Notes

- **Sheet Name Sanitization**: All sheet names are sanitized to ensure they are valid according to Excel’s limitations (e.g., not exceeding 31 characters and excluding invalid characters such as `\`, `/`, `?`, `*`, `[`, `]`).
- **Error Handling**: The scripts include error handling for cases where no coordinators or promotors are found or where matches are not found for specific data.
- **Visibility Management**: The visibility state of sheets is managed carefully, with the original visibility restored once the processing is completed.
- **Template Consistency**: New tabs are created by copying a template sheet, ensuring that all generated tabs maintain a consistent format.

## Example Workflow

1. **Create Coordinator Tabs**:

   - Run the `CreateCoordinatorTabs` subroutine to automatically generate new tabs for each unique coordinator.
   - Each new tab will be populated with data specific to the coordinator and will include common values like `razonSocial`.

2. **Create Promotor Tabs**:

   - Use the `CreatePromotorTabs` script to create tabs for promotors under a specific coordinator, with data filtering based on the coordinator's information.

3. **Dynamic Data Validation**:

   - Apply data validation rules to ensure that the "COORDINADOR" and "PROMOTOR" columns dynamically adjust based on the available options in the respective tables.

4. **Base Salary Tab Creation**:
   - Run the `CreateBaseSalaryTabsIfMissing` subroutine to ensure that salary tabs are created for each promotor with base salary data.

## Version History

### Version 1.6.4

- **Added**: Enhanced tab creation and population for coordinators and promotors.
- **Improved**: Dynamic data validation handling for the "COORDINADOR" and "PROMOTOR" columns.
- **Fixed**: Visibility management and error handling improvements.
- **Optimized**: Code performance improvements for tab creation and data population.

### Version 1.6.2

- **Created**: `CreatePromotorTabs` to automate the creation of promoter tabs.
- **Improved**: Template consistency and data population for new tabs.

### Version 1.5.7

- **Created**: `CreateCoordinatorTabs` to automate the creation of coordinator tabs.
- **Improved**: Sheet name handling and error checks for existing sheets.

## License

This suite is provided as-is and can be freely used or modified for personal or commercial purposes. No warranty is provided.

---

For any questions or issues, please contact the author at [email@example.com].
