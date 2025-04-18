# Change Log for Module Changes (v0.9.0)

## Modules

### CreateCoordinatorAndPromotorTabs (v0.9.0)

- 🧠 **Dynamic protected tabs handling**: Replaced the hardcoded list of protected tabs with the `SKIP_SHEETS` variable. This allows for more flexible management of protected sheets.
- 💡 **Improved list management**: Added the activating sheet to the protected tabs list dynamically, ensuring it isn’t deleted in the process.
- 🧹 **General error handling**: Enhanced error messages for improved clarity. The message now includes a global `ERROR_GENERIC` message for better error reporting.
- 🛠 **Modular tab deletion**: The list of protected tabs is dynamically built and adjusted before calling `DeleteUnprotectedTabs`, making the subroutine more adaptable.
- 🚀 **Code optimization**: Removed redundancy by leveraging the `Split` function for better scalability in managing protected tabs.

### CreateCoordinatorTabs (v0.9.0)

- 🧠 Enhanced the process of creating coordinator-specific tabs with a streamlined and error-proof flow.
- 🔧 Added dynamic retrieval of column indices for "ALIAS" and "GERENCIA" in the "Coordinadores" table.
- 🔧 Incorporated improved handling of sheet creation, ensuring that new tabs are only created for unique coordinators.
- 🧹 Implemented checks to prevent errors if no coordinators are found or if no matches are found for a coordinator's alias.
- 🔧 Optimized the filter process to include only visible rows for each coordinator.
- 🔧 Automatically applied a sanitization process to ensure valid sheet names for each coordinator.
- 🔧 Introduced a more robust data transfer to newly created tabs, including common values like "razonSocial", "periodoDelPagoDel", etc.
- 🔧 Added error handling for issues with creating or updating tables in newly created sheets.
- 🧹 Fixed an issue where sheets were not being correctly hidden back to their original state after the process.
- 🧹 Cleaned up unnecessary references and improved the clarity of the code.

### CreatePromotorTabs (v0.9.0)

- 🧠 Enhanced subroutine to automate the creation of individual tabs for promotors based on filtered data from the "Promotores" table.
- 🔧 Added checks for matching promotor entries in the "Sueldos_Base" table before creating new tabs.
- 🔧 Improved data handling by sanitizing promotor names and ensuring valid sheet names.
- 🔧 Added sorting for the "Promotor" column to ensure the data is processed in ascending order.
- 🔧 Incorporated automatic table updates in newly created tabs by copying relevant data.
- 🧹 Prevented errors when no promotors are found by displaying an appropriate message and hiding all previously unhidden sheets.
- 🧹 Incorporated improved error handling to manage duplicate promotors and missing data.
- 🔧 Updated logic to automatically populate new tabs with common values (e.g., "razonSocial", "periodoDelPagoDel").
- 🧹 Refined the loop to clear previous data and ensure correct filtering of rows for each promotor before populating new tabs.
- 🧹 Enhanced performance by disabling screen updating and automatic calculation during processing.

### SumPagoNetoCoordinacion (v0.9.0)

- 🏷️ Replaced hardcoded references (`"P"`, `"J4"`) with **named constants**: `SHEET_NAMES_COLUMN` and `TARGET_CELL`.
- 💬 Swapped inline error strings with **standardized constants**: `ERROR_SHEET_NOT_FOUND` and `ERROR_GENERIC` for consistent messaging.
- 🧼 Improved readability and structure with **cleaner variable usage** and consistent naming.
- 🧯 Retained `ErrorHandler` block for **graceful error management** and user-friendly feedback.
- 📊 Maintained core logic: **sums all "PAGO NETO" values** from listed sheets, and adds active sheet if not listed.
- 📍 Final result still placed in cell defined by `TARGET_CELL` (previously hardcoded as `J4`).

### SumPagoNetoGerencia (v0.9.0)

- 🧠 Introduced `GetManagerPagoNeto` to **separately retrieve and include the manager's own "PAGO NETO" value**, ensuring it's not skipped from the overall sum.
- 🏷️ Replaced hardcoded `"J4"` reference with the **named constant** `TARGET_CELL` for clarity and maintainability.
- 💬 Replaced inline error message with **standardized constant** `ERROR_INVALID_SHEET` for consistent user feedback.
- 🧯 Retained global `ErrorHandler` with improved feedback using `ERROR_GENERIC` and procedure name.
- 📊 Final total now includes both **sum of all sheets** and **manager's own "PAGO NETO"**, improving accuracy.
- ✅ Maintains original function: result is still placed in the target sheet's defined cell.

### General Enhancements for Module (v0.9.0)

- 🧠 **Dynamic and flexible handling**: Replaced hardcoded values with dynamic variables like `SKIP_SHEETS` for improved management of protected sheets and flexibility in processing.
- 🔧 **Improved tab creation**: Streamlined processes for creating coordinator and promotor-specific tabs, including better data retrieval, dynamic sheet naming, and efficient table updates.
- 🧹 **Robust error handling**: Enhanced error messages and added checks to handle edge cases like missing coordinators, promotors, or invalid data.
- 🛠 **Optimized performance**: Leveraged functions like `Split` for managing protected tabs, and applied optimizations like disabling screen updates to boost performance during processing.
- 🚀 **Modular and scalable**: Improved modularity in tab deletion and list management, allowing for more adaptable and error-proof subroutines.
- 💬 **Consistent messaging**: Standardized error messages and constants (e.g., `ERROR_GENERIC`, `ERROR_SHEET_NOT_FOUND`) to ensure clarity across the module.

## UtilsModules

### UtilsCollections (v0.9.0)

- 🆕 Enhanced `IsInArray` with a new **`caseInsensitive` optional parameter**, allowing flexible string comparison.
- 🔁 Updated loop logic in `IsInArray` to **normalize case** when `caseInsensitive` is enabled.
- 🧪 Preserved original behavior as **case-sensitive by default**, maintaining compatibility with previous usage.
- 📄 No changes made to `KeyExists`, ensuring stable behavior for collection key lookups.
- ✅ Module retains clean separation of concerns: focused utility functions for **collection and array operations**.

### UtilsCoordinator (v0.9.0)

- 🧩 **Refactored hardcoded strings** into named constants for sheet names, table names, and column names (`COLABORADORES_SHEET`, `COORDINADORES_TABLE`, `GERENCIA_COLUMN`, `ALIAS_COLUMN`, `TAB_SUFFIX`), enhancing maintainability.
- 🛠 **Added error handling** in `CreateCoordinatorTabs_newTabs` to check if `newTabs` is initialized, preventing runtime errors.
- ✍️ **Standardized documentation** across functions and module header for consistency and clarity.
- 🔧 **Improved functionality** by ensuring the collection of newly created tabs is accessible only if properly initialized.

### UtilsData (v0.9.0)

- 🧩 **Added `GetManagerPagoNeto` function** to retrieve the "PAGO NETO" value from the manager's sheet, enhancing the module's capabilities.
- 🛠 **Refactored `SumPagoNetoFromSheets`** to use named constants (`SKIP_SHEETS`, `COLUMN_A`, `PAGO_NETO_TEXT`) for improved maintainability and readability.
- 🔧 **Introduced error handling** in the `GetManagerPagoNeto` function to provide feedback on errors and ensure better handling of invalid data.
- 🔄 **Refined the row-checking logic** in `IsRowEmpty` by adjusting the start column and improving the handling of table-based worksheets.
- ✍️ **Updated documentation** for better clarity on the functionality of each function and added details for new parameters and error handling strategies.

### UtilsManager (v0.9.0)

- 🧠 **Refactored hardcoded values** by replacing sheet and table names with dynamic constants (`COLABORADORES_SHEET`, `GERENTES_TABLE`, `NOMBRE_COLUMN`), enhancing flexibility.
- 💡 Improved **error handling** in the `GetManagerAliasFromNombreGerente` function, ensuring clearer messages for users when errors occur.
- 🔧 Enhanced **sheet renaming logic** to check for existing sheets with the same alias name before proceeding with renaming, preventing conflicts.
- 🧩 **Optimized the process** of retrieving the manager's alias from the "Gerentes" table, ensuring consistency in handling manager data across modules.

### UtilsNumberToText (v0.9.0)

- 🧠 Centralized number-to-text conversion logic for Spanish language, enabling reuse across various scripts.
- 🔧 Introduced dynamic constants (`MAX_LIMIT`, `ZERO_TEXT`, `PESOS_TEXT`) for easier adjustments and clearer code.
- 🧹 Improved error handling for values exceeding the defined limit, providing dynamic feedback based on the `MAX_LIMIT`.
- ⚙️ Optimized handling of numeric inputs, ensuring compatibility with other modules and enhancing reliability.

### UtilsSheet (v0.9.0)

- 🧠 Improved management of sheet functions with enhanced capabilities for sheet deletion and name sanitization.
- 🔧 Updated `DeleteUnprotectedTabs` to ensure better handling of protected tabs and the active tab.
- 🧹 Refined `SanitizeSheetName` function to handle invalid characters and name length limitations more efficiently.
- 🧹 Improved `IsInNewTabs` to check newly created sheets more effectively, ensuring compatibility with other modules.

### UtilsTable (v0.9.0)

- 🧠 Enhanced the `SortTableAlphabetically` function to support optional parameters for sorting order (`sortOrder`) and sorting data option (`dataOption`).
- 🧠 Improved the `PopulateTableWithCollection` function by adding an optional `columnIndex` parameter to specify which column to populate in the table.
- 🔧 Default values for `sortOrder` (xlAscending) and `dataOption` (xlSortNormal) are now used in `SortTableAlphabetically`, improving flexibility.
- 🔧 The `PopulateTableWithCollection` function now allows for populating a specific column (default is the first column).
- 🧹 Minor refinements and documentation updates.

### UtilsValidation (v0.9.0)

- 🧠 **Improved Formula for Alias List Retrieval**  
  The `GetAliasList` function was updated to use constants for table and column names (e.g., `PROMOTORES_TABLE`, `COORDINACION_COLUMN`, `ALIAS_COLUMN`) for easier maintenance and dynamic updates.

- 🔧 **Code Refactoring for Maintainability**
  - The alias list retrieval formula in `GetAliasList` now references the table and column names using constants instead of hardcoded strings.
- 🧹 **No Bug Fixes or Refinements in this Version**  
  This update focuses on improvements in code structure for better readability and flexibility.

#### General Utils Modules Enhancements (v0.9.0)

- 🧠 **General improvements across multiple modules** to enhance maintainability, flexibility, and error handling.
- 🔧 **Replaced hardcoded strings** with dynamic constants in various modules (e.g., `UtilsCoordinator`, `UtilsManager`, `UtilsValidation`), ensuring easier updates and better organization.
- 🧩 **Improved error handling** in several functions, including `CreateCoordinatorTabs_newTabs`, `GetManagerPagoNeto`, `GetManagerAliasFromNombreGerente`, and others, to ensure clear feedback and prevent runtime issues.
- ⚙️ **Optimized functions** like `IsInArray`, `SortTableAlphabetically`, `PopulateTableWithCollection`, and `SanitizeSheetName` to improve performance, flexibility, and compatibility with other modules.
- ✍️ **Updated documentation** across all modules, including standardized headers and improved descriptions for added parameters, error handling, and function capabilities.
- 🧹 **Refined table management** functions like `SortTableAlphabetically`, `PopulateTableWithCollection`, and `IsInNewTabs`, ensuring they handle specific scenarios and edge cases more effectively.
- 🔄 **Refined logic** in functions like `IsRowEmpty`, `SanitizeSheetName`, and `IsInNewTabs`, ensuring smoother integration with other parts of the codebase and improving user experience.

## Class Modules

### clsDictionary (v0.9.0)

- 🔑 **Add method**: Now raises an error using `ERROR_KEY_EXISTS` constant if the key already exists.
- ✅ **Exists method**: Checks if a key exists in the dictionary, now uses `CStr(key)` for consistency.
- 🔍 **GetValue method**: Now explicitly converts keys to a string (`CStr(key)`) and raises an error with `ERROR_KEY_NOT_FOUND` constant if the key is not found.
- 🗑 **Remove method**: Removes a key-value pair from the dictionary with error handling.
- 🔄 **Replace method**: Replaces the value for an existing key, raises an error if the key is not found.
- 📋 **GetKeys method**: Returns all keys in the dictionary as an array.
- 💼 **GetValues method**: Returns all values in the dictionary as an array.
- 🧹 **Clear method**: Clears all key-value pairs in the dictionary.
- 🔢 **Count method**: Returns the number of key-value pairs in the dictionary.
- 🔧 **FindKeyIndex helper function**: A private function that finds the index of a key in the dictionary.
- 📝 **Updated Documentation**: Includes versioning, author, and class description for better clarity.

This version introduces constants for error handling, refined type handling, and improved documentation, enhancing the overall maintainability and consistency of the class.

## Worksheet Modules

### Worksheet(Manager)

#### Worksheet_Change (v0.9.0)

- 🧠 **Dynamic handling for changes** in the "Nombre_Gerente" range and the "COORDINADOR" column.
- 🧹 **Enhanced error handling**: Added checks for empty "Nombre_Gerente" and invalid manager aliases to improve the user experience.
- 💡 **Coordinator alias population**: Dynamically retrieves and populates the "Coordinadores_Gerencia_Activa" table based on the manager's alias.
- 📊 **Table sorting**: After adding coordinator aliases, the table is sorted alphabetically by the "COORDINADOR" column.
- 🛠 **Improved row validation**: Introduced a `ProcessRow` function to handle row-by-row validation and ensure data consistency.
- 💬 **Clear user feedback**: Displays meaningful error messages when no coordinators are found for the specified manager.

#### Worksheet_Calculate (v0.9.0)

- 🧠 **Dynamic table handling**: The script now checks for the presence of a single table on the sheet before proceeding with calculations.
- 📋 **Processing coordinator data**: Loops through the "COORDINADOR" column and dynamically retrieves and processes the corresponding "PROMOTOR" column.
- 💡 **Improved error handling**: Gracefully handles cases where multiple tables are found or no tables exist, offering user-friendly error messages.
- 🧹 **Code cleanup**: Streamlined logic for processing rows in the "COORDINADOR" and "PROMOTOR" columns, ensuring efficiency.
- 🧽 **Increased modularity**: The updated flow for processing rows improves the script’s readability and maintainability.

#### General Worksheeet Enhancements (v0.9.0)

- 🧩 **Centralized error handling**: Introduced a global error handler for consistent management of errors across the two main events.
- 🧹 **Simplified logic** for handling updates to the "COORDINADOR" column, ensuring efficient processing without redundancy.
- 💬 **User-friendly feedback**: Enhanced error messages provide clearer guidance on issues such as missing or invalid data.

## Constants

### Constants(v0.9.0)

- 🧠 Introduced a new module containing public constants for improved code organization.
- 🔧 Grouped constants logically, such as sheet names, table names, column names, error messages, and text properties, to ensure consistency across the project.
- 🔧 Constants for sheets, tables, columns, properties, and error messages added for ease of reference and maintainability.
- 🧹 Avoided hardcoding values throughout the project by referencing these constants in relevant code sections.
