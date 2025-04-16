# Excel Automation Suite - Version 0.8.0

## Description

The **Excel Automation Suite** is a powerful set of modular VBA scripts that streamline Excel data workflows. Designed for workbooks involving **managers (Gerentes), coordinators, and promotors**, this suite automates the creation of personalized sheets, sets up dynamic validations, manages tab visibility, and ensures data accuracy through robust utilities.

Version **0.8.0** introduces **reporting and visualization capabilities**:
- Adds two new sheets: `Resultados` (Pivot Reports) and `Dashboard` (Visual Charts).
- Automatically generates **pivot tables** by manager and builds **bar chart dashboards** for coordinators, promotors, locations, and courses.
- Protects these sheets from cleanup and data iteration logic.

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

### 6. **Reporting and Dashboarding (NEW in Version 0.8.0)**

#### 📊 Resultados (Pivot Reports)
- Automatically generates pivot tables using the active manager's data.
- Source data is aggregated and filtered from individual coordinator and promotor sheets.

#### 📈 Dashboard (Bar Chart Visualizations)
- Charts are created dynamically based on the pivot tables in `Resultados`.
- Includes the following visual reports:
  - **Ventas por Coordinación**
  - **Ventas por Promotor**
  - **Ventas por Plantel**
  - **Ventas por Curso**
  - **Ventas por Plantel por Curso**
  - **Ventas de Cursos por Plantel**
  - **Ventas de Coordinación por Plantel**

### 7. **Modular Utility Functions & Cross-Platform Class Support (Version 0.7.0)**

- Introduced a **Class Module** to encapsulate environment-specific logic (e.g., file paths, keyboard shortcuts).
- Ensures compatibility with **both Windows and Mac**, adapting behavior based on the user's platform at runtime.

Utility functions are split into **8 focused modules**:

- `UtilsCollections`: Sheet existence, empty row checks, collection utilities.
- `UtilsCoordinator`: Filters and collects coordinators by manager alias.
- `UtilsData`: Data copying, merging, and lookup helpers. (Now excludes `Resultados` and `Dashboard`)
- `UtilsManager`: Alias lookup and name resolution for Gerentes.
- `UtilsNumberToText`: Spanish number-to-text conversion.
- `UtilsSheet`: Sheet creation, naming, duplication, and visibility management.
- `UtilsTable`: Table-based filtering, sorting, and clearing.
- `UtilsValidation`: Setup of dependent dropdowns and data validation.

## Requirements

- **Excel**: Microsoft Excel with macro support.
- **VBA Access**: Requires editor access via `Alt + F11`.
- Template sheet(s) must be present and correctly named.

## Setup Instructions

1. Open the workbook in Excel.
2. Press `Alt + F11` to open the VBA Editor.
3. Add or update the following modules:
   - All 8 utility modules (see Feature #7).
   - Main script modules (e.g., `CreateCoordinatorTabs`, `CreatePromotorTabs`).
   - Reporting modules/scripts for pivot and chart creation.
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
- **Protected Sheets**: `Resultados` and `Dashboard` are excluded from deletion and data scanning logic.

## Example Workflow

1. On the Gerente sheet (e.g., “Sheet1”), confirm full name in `B2`, alias in `B1`.
2. Run `CreateCoordinatorTabs` to generate all related coordinator tabs.
3. Run `CreatePromotorTabs` for each coordinator sheet to generate promotors.
4. Run `CreateBaseSalaryTabsIfMissing` to generate any missing salary sheets.
5. Run reporting setup to generate `Resultados` pivot tables and the `Dashboard` charts.
6. Use `RenameGerenteTabToAlias` to rename the original sheet based on alias.

## Version History

### Version 0.8.0

- **New Reporting System**:
  - Added `Resultados` sheet with pivot tables.
  - Added `Dashboard` sheet with auto-generated bar charts.
  - Manager-specific visual reports now available in Excel directly.

- **Script Updates**:
  - `UtilsData`: Excludes `Resultados` and `Dashboard` from manager pay calculations.
  - `CreateCoordinatorAndPromotorTabs`: Prevents deletion of `Resultados` and `Dashboard` during cleanup.

### Version 0.7.0

- **Modularization Complete**: Split `mod_UtilsModule` into 8 separate modules.
- **Cross-Platform Support**: Introduced a **Class Module** to handle platform-specific behavior for Mac and Windows.
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

### Version 0.5.7

- First release of `CreateCoordinatorTabs`.

## License

This suite is provided **as-is** and may be freely used or modified for personal or commercial projects. No warranty is provided.

---

For support or contributions, contact **Juan Pablo Garcia Murillo**.