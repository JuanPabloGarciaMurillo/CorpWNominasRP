# Excel Automation Suite - Version 0.7.0

## Description

The **Excel Automation Suite** is a powerful set of modular VBA scripts that streamline Excel data workflows. Designed for workbooks involving **managers (Gerentes), coordinators, and promotors**, this suite automates the creation of personalized sheets, sets up dynamic validations, manages tab visibility, and ensures data accuracy through robust utilities.

Version **0.7.0** marks a major refactor: **all utility functions** were broken out of the monolithic `mod_UtilsModule` into **eight focused utility modules**, enhancing readability, reuse, and maintainability.

## Features

### 1. **CreateCoordinatorTabs** (Version 0.6.5)

- Creates new tabs for each unique coordinator linked to a manager.
- Filters the `Coordinadores` table by the manager's **alias** (`B1`) from the `Gerentes` table.
- Uses a template sheet to create each new tab.
- Copies coordinator metadata and applies common formatting and sorting.

### 2. **CreatePromotorTabs** (Version 0.6.5)

- For each coordinator tab, generates tabs for associated promotors.
- Filters promotors from the `Promotores` table based on coordinator name.
- Validates against `Sueldos_Base` table before creating a tab.
- Ensures clean, unique naming and proper data placement.

### 3. **CreateBaseSalaryTabsIfMissing**

- Checks each promotor for a matching salary entry.
- Creates salary tabs **only if** they don’t already exist.
- Ensures data integrity with no duplication.

### 4. **RenameGerenteTabToAlias** (Version 0.6.5)

- Renames the active Gerente sheet based on the alias from the `Gerentes` table.
- Uses the value in `B2` to perform lookup.
- Prevents name collisions with robust error handling and sheet existence checks.

### 5. **Dynamic Data Validation**

- Column **COORDINADOR**: Validated based on selected **Gerente** (manager).
- Column **PROMOTOR**: Dynamically validated based on selected **Coordinador**.
- Validation lists update automatically based on hierarchy and selection.

### 6. **Modular Utility Functions (Version 0.7.0)**

Utility functions are now split into **8 focused modules**:

- `UtilsCollections`: Sheet existence, empty row checks, collection utilities.
- `UtilsCoordinator`: Filters and collects coordinators by manager alias.
- `UtilsData`: Data copying, merging, and lookup helpers.
- `UtilsManager`: Alias lookup and name resolution for Gerentes.
- `UtilsNumberToText`: Spanish number-to-text conversion.
- `UtilsSheet`: Sheet creation, naming, duplication, and visibility management.
- `UtilsTable`: Table-based filtering, sorting, and clearing.
- `UtilsValidation`: Setup of dependent dropdowns and data validation.

Each module handles a specific concern, making the codebase easier to test and extend.

## Requirements

- **Excel**: Microsoft Excel with macro support.
- **VBA Access**: Requires editor access via `Alt + F11`.
- Template sheet(s) must be present and correctly named.

## Setup Instructions

1. Open the workbook in Excel.
2. Press `Alt + F11` to open the VBA Editor.
3. Add or update the following modules:
   - All 8 utility modules (see Feature #6).
   - Main script modules (e.g., `CreateCoordinatorTabs`, `CreatePromotorTabs`).
4. Ensure:
   - `Colaboradores` sheet contains `Gerentes`, `Coordinadores`, and `Promotores` tables.
   - Template sheet(s) exist and follow naming conventions.

## Parameters

- **Most scripts** operate contextually on the active sheet.
- **RenameGerenteTabToAlias**: Uses `B2` for the Gerente full name.
- **CreateCoordinatorTabs**: Uses `B1` for the manager’s alias to match coordinators.

## Returns

- Most subroutines return no value (Sub).
- `RenameGerenteTabToAlias` returns a **Boolean** (success/failure).
- Utility functions return appropriate values (arrays, ranges, Booleans).

## Notes

- **Sheet Naming**: All sheet names are sanitized to remove invalid characters.
- **Template Usage**: New sheets are created by copying from a defined template.
- **Data Integrity**: All lookups are case-insensitive, and validations are pre-checked.
- **Merged Cells**: Lookups only consider the first cell in merged regions (e.g., `B2:D2` → `B2`).

## Example Workflow

1. On the Gerente sheet (e.g., “Sheet1”), confirm full name in `B2`, alias in `B1`.
2. Run `CreateCoordinatorTabs` to generate all related coordinator tabs.
3. Run `CreatePromotorTabs` for each coordinator sheet to generate promotors.
4. Run `CreateBaseSalaryTabsIfMissing` to generate any missing salary sheets.
5. Use `RenameGerenteTabToAlias` to rename the original sheet based on alias.

## Version History

### Version 0.7.0

- **Modularization Complete**: Split `mod_UtilsModule` into 8 separate modules.
- **Improved Maintainability**: Each module is now focused and independent.
- **Optimized Reuse**: Common logic now abstracted and reusable across all scripts.

### Version 0.6.5

- Added `RenameGerenteTabToAlias`.
- Enhanced coordinator/promotor tab creation.
- Improved error handling, sheet naming, and table population.

### Version 0.6.4

- Added dynamic data validation for COORDINADOR and PROMOTOR.
- Improved data filtering and formatting.
- Coordinators and promotors now filtered by Gerente alias.

### Version 0.6.2

- Initial release of `CreatePromotorTabs`.

### Version 1.5.7

- First release of `CreateCoordinatorTabs`.

## License

This suite is provided **as-is** and may be freely used or modified for personal or commercial projects. No warranty is provided.

---

For support or contributions, contact **Juan Pablo Garcia Murillo**.