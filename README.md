# Excel Automation Suite - Version 0.9.2

## Description

The **Excel Automation Suite** is a powerful collection of VBA tools that streamline worksheet creation, data validation, and reporting processes for managers, coordinators, and promotors. Built for efficiency and maintainability, this suite automates repetitive tasks and provides advanced features like dynamic dropdowns, smart filtering, and visual reporting dashboards.

Version **0.9.2** introduces improved error handling across all major modules, enhanced sum calculations for `PAGO NETO`, and significant utility upgrades to support robust, reusable, and clean scripting. Key modules such as `UtilsCoordinator`, `UtilsData`, and `UtilsErrorHandling` have been refactored for better modularity, performance, and developer experience.

---

## Features

### 📄 **Coordinator & Promotor Management**

- **`CreateCoordinatorTabs`**: Generates tabs for each coordinator under a selected manager.
- **`CreatePromotorTabs`**: Creates tabs for promotors under each coordinator, verifying salary data.
- **`CreateBaseSalaryTabsIfMissing`**: Adds missing promotor salary tabs based on lookup in `Tabuladores`.
- **`RenameGerenteTabToAlias`**: Renames the active Gerente tab using the alias pulled from `B2`.

### 🧩 **Dynamic Data Validation**

- **COORDINADOR Dropdown**: Populated dynamically based on the selected **Gerente**.
- **PROMOTOR Dropdown**: Depends on the selected **Coordinador**, with lookup from the Promotores table.

### 📊 **Automated Reporting**

- **`Resultados`** Sheet:

  - Includes pivot tables for:
    - Ventas por Coordinación
    - Ventas por Promotor
    - Ventas por Plantel
    - Ventas por Curso
  - Auto-refreshed with updated data.

- **`Dashboard`** Sheet:
  - Displays **auto-generated bar charts** based on `Resultados`.
  - Updated layout for cleaner visuals and better chart spacing (since v0.9.0).
  - Syncs automatically with report pivots.

### 🧰 **Modular Utilities (v0.9.1)**

Scripts are organized across 9 utility modules to maximize clarity, reuse, and testability:

- `UtilsCollections`: Collection helpers, sheet checks, existence verification.
- `UtilsCoordinator`: Filters coordinators by manager alias; manages coordinator-specific tab generation.
- `UtilsData`: Handles table lookups, filtered range extraction, and `PAGO NETO` sum calculations.
- `UtilsManager`: Performs Gerente name-to-alias lookups and tab renaming.
- `UtilsNumberToText`: Converts numeric values to Spanish text (e.g., for salary in words).
- `UtilsSheet`: Creates, copies, sanitizes, and renames sheets.
- `UtilsTable`: Extracts, filters, and sorts table ranges.
- `UtilsValidation`: Applies dependent dropdown validations for Coordinador/Promotor.
- `UtilsErrorHandling`: Centralized error catching/logging with `HandleError`.

---

## Requirements

- Microsoft Excel with macro support enabled.
- Required sheets:
  - `Colaboradores`
  - `Dashboard`
  - `Resultados`
- Must include at least one template sheet (e.g., "Plantilla") for copying.

---

## Setup

1. Open the workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Import all script modules and utility modules from the suite.
4. In the `Colaboradores` sheet, ensure the following named tables exist:
   - **Gerentes** → Columns A3:B
   - **Coordinadores** → Columns D3:F
   - **Promotores** → Starts at A22:D (dynamic size)

---

## Parameters

| Subroutine                | Key Cell     | Notes                                     |
| ------------------------- | ------------ | ----------------------------------------- |
| `CreateCoordinatorTabs`   | `B1`         | Reads Gerente alias from the active sheet |
| `RenameGerenteTabToAlias` | `B2`         | Uses Gerente name to apply tab alias      |
| `CreatePromotorTabs`      | `Sheet.Name` | Uses sheet name as the coordinator alias  |

---

## Returns

- Most routines are `Sub` procedures (do not return a value).
- `RenameGerenteTabToAlias` returns a `Boolean` indicating success or failure.
- Utility functions may return:
  - Arrays
  - Dictionaries
  - Filtered `Range` objects
  - `Boolean` values for validation

---

## Example Workflow

1. Manager selects their **Gerente name** in cell `B2` and **alias** in `B1`.
2. Run `CreateCoordinatorTabs` to generate coordinator-specific tabs.
3. From each coordinator sheet, run `CreatePromotorTabs` to generate promotor tabs.
4. Run `CreateBaseSalaryTabsIfMissing` to ensure salary records exist for all promotors.
5. Refresh reports:
   - `Resultados` → Pivot tables
   - `Dashboard` → Auto-generated charts
6. Optionally run `RenameGerenteTabToAlias` to rename the main tab based on the alias.

---

## Version History

### Version 0.9.1

- 🧠 **Improved Error Handling**
  - Introduced centralized `HandleError` routine in `UtilsErrorHandling`.
  - Removed scattered `MsgBox` usage and replaced with structured logging.
- 🔧 **Refined Sum Calculations**
  - Simplified logic in `SumPagoNetoCoordinacion` and `SumPagoNetoGerencia`.
  - Enabled direct writing of totals with `StoreTotalInTargetCell`.
- 🧩 **Utility Upgrades**
  - Refactored `UtilsData`, `UtilsCollections`, and `UtilsCoordinator` with reusable patterns.
  - Improved error resilience and feedback mechanisms throughout utility layers.

### Version 0.9.0

- ✨ **Dashboard Layout Redesign**
  - Spaced and aligned visual elements for readability.
  - Ensures visual consistency across charts.
- 📈 **Automated Chart Generation**
  - Dynamically generates bar charts from pivot data in `Resultados`.
  - Removes duplicates and adjusts based on live data.
- 🧹 **Minor Utility Fixes**
  - Refined exclusion of `Dashboard` and `Resultados` from certain loops and cleanups.

### Version 0.8.0

- 📊 Introduced pivot reports in `Resultados` sheet.
- 📈 Added dynamic bar charts in `Dashboard`.
- 🛡️ Protected reporting sheets from accidental deletion.

### Version 0.7.0

- 🧩 Split original monolithic `UtilsModule` into 8 utility modules.
- 🧪 Enabled cross-platform VBA compatibility.
- 🧼 Improved script modularity and structure.

### Version 0.6.5

- 🗂️ Enhanced tab creation logic for coordinators and promotors.
- 🧼 Enabled alias-based renaming for cleaner sheet naming.
- 🧪 Improved table validation and format consistency.

---

## License

This suite is offered **as-is** under an open-use model. No warranties or guarantees. Freely use or modify for personal or commercial purposes.

---

For feedback, suggestions, or bug reports, please contact **Juan Pablo Garcia Murillo**.
