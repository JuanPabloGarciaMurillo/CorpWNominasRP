# Change Log for Module Changes (v0.9.2)

## Modules

### **1. `CreatePromotorTabs.vba`**

- **Refactored Tab Creation**:

  - Replaced repeated logic for creating new tabs with the reusable `CreateNewTab` function from `UtilsSheet`.
  - Improved readability and maintainability by delegating tab creation to a centralized function.

- **Refactored Table Population**:

  - Replaced the inline logic for populating tables with the reusable `PopulateTable` function from `UtilsTable`.
  - Centralized row validation, header mapping, and column index validation.

- **Error Handling**:

  - Unified error handling by routing exceptions through `HandleError`.
  - Added `Debug.Print` statements for better traceability during debugging.

- **Code Cleanup**:
  - Removed duplicate logic for resetting filters and restoring application settings.
  - Improved readability by breaking down large blocks of code into smaller, reusable functions.

---

### **2. `CreateCoordinatorTabs.vba`**

- **Refactored Tab Creation**:

  - Replaced repeated logic for creating new tabs with the reusable `CreateNewTab` function from `UtilsSheet`.

- **Refactored Table Population**:

  - Replaced the inline logic for populating tables with the reusable `PopulateTable` function from `UtilsTable`.

- **Added Filtering and Validation**:

  - Introduced the `FilterAndValidateVisibleCells` function to centralize filtering and validation logic for visible cells.
  - Improved maintainability by delegating filtering and validation to a reusable function.

- **Error Handling**:

  - Unified error handling by routing exceptions through `HandleError`.
  - Added `Debug.Print` statements for better traceability during debugging.

- **Code Cleanup**:
  - Removed duplicate logic for resetting filters and restoring application settings.
  - Improved readability by breaking down large blocks of code into smaller, reusable functions.

---

### **3. `Constants.vba`**

- **Improved Error Messages**:
  - Standardized error message formatting for better readability.
  - Updated constants like `ERROR_SHEET_NOT_FOUND`, `ERROR_EMPTY_MANAGER_CELL`, and `ERROR_NO_COORDINATORS` to improve clarity.

---

## UtilsModules

### **4. `UtilsSheet.vba`**

- **Added `CreateNewTab` Function**:
  - Centralized logic for creating new tabs from a template.
  - Added flexibility to handle `tabName` as a `Variant` and convert it to a string using `CStr`.
  - Improved error handling for invalid `placementAfter` parameters.
  - Ensures all new tabs are explicitly set to visible and tracked in the `newTabs` collection.

---

### **5. `UtilsTable.vba`**

- **Added `PopulateTable` Function**:

  - Centralized logic for populating tables with filtered data.
  - Handles row validation, header mapping, and column index validation.
  - Eliminated duplicate logic in `CreatePromotorTabs` and `CreateCoordinatorTabs`.

- **Added `FilterAndValidateVisibleCells` Function**:
  - Centralized logic for filtering and validating visible cells in a table.
  - Ensures that only valid rows matching the filter criteria are processed.
  - Handles cases where no visible cells are found or validation fails.

---

## **File-Specific Changes Summary**

| **File**                    | **Changes**                                                                                     |
| --------------------------- | ----------------------------------------------------------------------------------------------- |
| `CreatePromotorTabs.vba`    | Refactored tab creation and table population logic, improved error handling, and added cleanup. |
| `CreateCoordinatorTabs.vba` | Refactored tab creation, table population, and filtering logic, and improved error handling.    |
| `UtilsSheet.vba`            | Added `CreateNewTab` for reusable tab creation logic.                                           |
| `UtilsTable.vba`            | Added `PopulateTable` and `FilterAndValidateVisibleCells` for reusable table operations.        |
| `Constants.vba`             | Standardized error messages for better readability and consistency.                             |

---
