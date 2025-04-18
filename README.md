# Excel Automation Suite - Version 0.9.0

## Description

The **Excel Automation Suite** is a powerful collection of VBA tools that streamline worksheet creation, data validation, and reporting processes for managers, coordinators, and promotors. Built for efficiency and maintainability, this suite automates repetitive tasks and provides advanced features like dynamic dropdowns, smart filtering, and visual reporting dashboards.

Version **0.9.0** introduces enhancements to the `Dashboard` layout and dynamic chart generation, improving clarity and usability for visual reports. It also includes minor updates across utility modules to support the changes.

---

## Features

### 📄 **Coordinator & Promotor Management**

- **`CreateCoordinatorTabs`**: Generates tabs for each coordinator under a selected manager.
- **`CreatePromotorTabs`**: Creates tabs for promotors under each coordinator, verifying salary data.
- **`CreateBaseSalaryTabsIfMissing`**: Adds missing promotor salary tabs if needed.
- **`RenameGerenteTabToAlias`**: Renames Gerente tabs using their alias (pulled from cell `B2`).

### 🧩 **Dynamic Data Validation**

- **COORDINADOR Dropdown**: Depends on selected **Gerente**.
- **PROMOTOR Dropdown**: Depends on selected **Coordinador**.

### 📊 **Automated Reporting**

- **`Resultados`** Sheet:

  - Contains pivot tables for:
    - Ventas por Coordinación
    - Ventas por Promotor
    - Ventas por Plantel
    - Ventas por Curso

- **`Dashboard`** Sheet:
  - Contains **dynamically generated bar charts** based on `Resultados` pivots.
  - Charts are refreshed automatically for accurate reporting.
  - Redesigned in v0.9.0 with improved layout and chart spacing.

### 🧰 **Modular Utilities (v0.7.0+)**

Scripts are modularized across 8 utility modules for clarity and reuse:

- `UtilsCollections`: Collection helpers and sheet checks.
- `UtilsCoordinator`: Filters coordinators by manager alias.
- `UtilsData`: Table lookups, filters, and processing (excludes `Dashboard`/`Resultados`).
- `UtilsManager`: Gerente alias lookups.
- `UtilsNumberToText`: Converts numbers to Spanish words.
- `UtilsSheet`: Creates, renames, and sanitizes sheets.
- `UtilsTable`: Table filters, sorting, and extraction.
- `UtilsValidation`: Handles dynamic dropdown validations.

---

## Requirements

- Microsoft Excel with macro support.
- Sheets `Colaboradores`, `Resultados`, and `Dashboard` must be present.
- Template sheet(s) should exist and be named accordingly.

---

## Setup

1. Open the workbook.
2. Press `Alt + F11` to access the VBA editor.
3. Import all script modules and utility modules.
4. Ensure `Colaboradores` contains these tables:
   - **Gerentes** (A3:B)
   - **Coordinadores** (D3:F)
   - **Promotores** (A22:D)

---

## Parameters

| Subroutine                | Key Cell     | Notes                                     |
| ------------------------- | ------------ | ----------------------------------------- |
| `CreateCoordinatorTabs`   | `B1`         | Reads Gerente alias from active sheet     |
| `RenameGerenteTabToAlias` | `B2`         | Uses Gerente name to find and apply alias |
| `CreatePromotorTabs`      | `Sheet.Name` | Uses current sheet name as coordinator    |

---

## Returns

- Most routines are `Sub` procedures (no return).
- `RenameGerenteTabToAlias` returns a `Boolean` (success).
- Utility functions may return arrays, dictionaries, or filtered ranges.

---

## Example Workflow

1. Manager selects their name in `B2`, alias in `B1`.
2. Run `CreateCoordinatorTabs` to generate coordinator tabs.
3. Run `CreatePromotorTabs` from each coordinator sheet.
4. Run `CreateBaseSalaryTabsIfMissing` to ensure salary tabs exist.
5. Run report generation to update:
   - `Resultados` (pivot tables)
   - `Dashboard` (bar charts)
6. Use `RenameGerenteTabToAlias` to rename the main tab.

---

## Version History

### Version 0.9.0

- ✨ **Dashboard Layout Redesign**:
  - Improved spacing, labeling, and alignment of charts.
  - Ensures consistency across visual elements.
- 📈 **Dynamic Chart Generation**:
  - Automatically creates bar charts based on `Resultados`.
  - Removes duplicates before charting.
- 🧹 **Minor Utility Improvements**:
  - Updates to `UtilsData`, `UtilsSheet`, and `UtilsTable` for cleaner handling of `Dashboard`/`Resultados`.

### Version 0.7.0

- 🧩 Refactored utility functions into 8 dedicated modules.
- 🧪 Introduced cross-platform (Windows/Mac) compatibility.
- 🔄 Improved modular structure and error handling.
### Version 0.8.0

- 🧠 Introduced `Resultados` and `Dashboard` sheets.
- 📊 Created pivot tables for sales metrics.
- 📈 Initial implementation of bar charts.
- 🛡️ Excluded reporting sheets from deletion and data loops.

### Version 0.7.0

- 🧩 Refactored utility functions into 8 dedicated modules.
- 🧪 Introduced cross-platform (Windows/Mac) compatibility.
- 🔄 Improved modular structure and error handling.

### Version 0.6.5

- 🗂️ Coordinator and promotor tab creation logic improved.
- 🧼 Added alias-based sheet renaming.
- 🚦 Robust validations and cleaner formatting.

---

## License

This suite is provided **as-is** under an open-use model. No warranty is provided. Use or modify freely for personal or commercial use.

---

For feedback, suggestions, or bug reports, contact **Juan Pablo Garcia Murillo**.
