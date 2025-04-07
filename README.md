# Excel Automation Suite - Version 1.6.5

## Description

The **Excel Automation Suite** is a collection of VBA scripts designed to automate and streamline data management processes in Excel workbooks. This suite includes subroutines that create coordinator and promoter-specific tabs, manage data validation, perform complex data lookups, and handle visibility and sheet creation with error handling and efficient workflows. The suite aims to reduce redundancy, improve maintainability, and ensure data consistency across various tasks.

## Features

### 1. **CreateCoordinatorTabs** (Version 1.6.4)

- Automates the creation of new coordinator tabs from a template.
- Filters and populates each tab with coordinator-specific data.
- Manages visibility, sorting, and common value copying.

### 2. **CreatePromotorTabs** (Version 1.6.4)

- Creates promoter tabs under each coordinator's sheet.
- Copies data and formatting from a template.
- Populates each sheet with filtered promoter data.

### 3. **CreateBaseSalaryTabsIfMissing**

- Creates salary tabs for each promoter if their base salary exists in the "Sueldos_Base" table.
- Ensures no duplicate tabs are created for promotors.

### 4. **Dynamic Data Validation Setup**

- Sets up dependent dropdowns for:
  - **COORDINADOR** column, filtered by Gerente selection.
  - **PROMOTOR** column, filtered by Coordinador selection.

### 5. **RenameGerenteTabToAlias** (Version 1.6.5)

- **New Feature!**
- Automatically renames the active Gerente sheet based on the value in cell `B2`, using the **Alias** from the `Gerentes` table in the `Colaboradores` sheet.
- Prevents duplicate sheet names and ensures alignment with user-friendly naming.
- Includes built-in error handling and validation.

### 6. **Utility Functions in `mod_UtilsModule`**

- Reusable helper functions:
  - Sheet creation and cleanup
  - Sheet existence checks
  - Row validation
  - Filtering and sorting
  - Number-to-words conversion (Spanish)
  - Net pay summation
  - Lookup utilities

## Requirements

- **Excel**: Microsoft Excel (macro-enabled).
- **VBA**: All scripts must be installed via the VBA editor (`Alt + F11`).

## Setup Instructions

1. **Open Excel Workbook**.
2. **Access the VBA Editor** (`Alt + F11`).
3. **Add/Update Modules**:
   - Add or update `mod_UtilsModule` with the latest utility functions.
   - Ensure the `Colaboradores` sheet has a properly named table: `Gerentes`.
4. **Run the `CreateCoordinatorAndPromotorTabs` Subroutine**:
   - This subroutine will now include a call to `RenameGerenteTabToAlias`, which will rename the active sheet automatically.

## Parameters

- Most scripts operate on the current workbook context and require no parameters.
- The `RenameGerenteTabToAlias` function uses the value in cell `B2` of the active sheet.

## Returns

- Most subroutines return nothing.
- `RenameGerenteTabToAlias` returns a Boolean indicating if the rename was successful.

## Notes

- **Naming Rules**: Sheet names are sanitized to prevent invalid characters and duplicates.
- **B2 Format**: The Gerente's full name must be correctly entered in `B2`, matching exactly with the `NOMBRE` column of the `Gerentes` table.
- **SheetExists** utility is used to validate sheet name uniqueness before renaming.
- **Merged Cell Support**: Even if `B2:D2` is merged, only the `B2` value is read.

## Example Workflow

1. Open a Gerente sheet (e.g., “Sheet1”) and confirm the full name in `B2`.
2. Run `CreateCoordinatorAndPromotorTabs`.
3. The script:
   - Creates coordinator and promotor tabs.
   - Renames the current Gerente sheet to the appropriate alias (e.g., `"PEDRO MORA"`).
4. Ensures consistent, user-friendly sheet naming across the workbook.

## Version History

### Version 1.6.5

- **Added**: `RenameGerenteTabToAlias` function to auto-rename Gerente sheets based on the `Gerentes` table.
- **Improved**: Sheet naming consistency and clarity using aliases.
- **Enhanced**: Sheet existence check now integrated for safe renaming.

### Version 1.6.4

- Enhanced tab creation and data population for coordinators and promotors.
- Improved dynamic data validation for COORDINADOR and PROMOTOR columns.
- Fixed visibility handling and optimized performance.

### Version 1.6.2

- Introduced `CreatePromotorTabs` for automated promotor tab generation.

### Version 1.5.7

- Initial release of `CreateCoordinatorTabs` with sheet naming and data logic.

## License

This suite is provided as-is and may be freely used or modified for personal or commercial purposes. No warranty is provided.

---

For questions, feedback, or contributions, please contact Juan Pablo Garcia Murillo.
