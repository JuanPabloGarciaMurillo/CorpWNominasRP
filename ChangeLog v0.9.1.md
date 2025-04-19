# Change Log for Module Changes (v0.9.0)

## Modules

### Constants (v0.9.1)

- 🧠 Version bump for module tracking
- 🔧 Updated version header from `0.9.0` to `0.9.1` in preparation for upcoming changes

### CreateCoordinatorAndPromotorTabs (v0.9.1)

- 🧠 Version bump for tracking latest enhancements
- 🔧 Replaced generic `MsgBox` error display with centralized `HandleError` function
- 🧹 Improved error handling consistency and maintainability

### CreateCoordinatorTabs (v0.9.1)

- 🧠 Version update for latest refinements
- 🔧 Replaced direct `MsgBox` call with `HandleError` for unified error reporting
- 🧹 Ensured consistent error logging via `Debug.Print` before `HandleError`

### CreatePromotorTabs (v0.9.1)

- 🧠 Improved robustness and error handling consistency
- 🔧 Replaced `MsgBox` with centralized `HandleError` call for unified feedback
- 🔧 Moved `SumPagoNetoCoordinacion` call to only run when `newTabs` has content
- 🔧 Standardized application settings reset (`CutCopyMode`, `ScreenUpdating`)
- 🧹 Cleaned up redundant exit logic and restored proper sheet visibility state

### SumPagoNetoCoordinacion (v0.9.1)

- 🧠 Refactored logic to simplify summing from a passed `newTabs` collection and the active sheet
- 🔧 Modified signature to accept `newTabs` and `targetSheet` as parameters for better modularity
- 🔧 Removed redundant logic related to range scanning and manual collection creation
- 🔧 Ensured the coordinator tab (target sheet) is always included in the total sum
- 🔧 Improved comments for clarity and consistency
- 🧹 Cleaned up unused variables and removed unnecessary conditional logic
- 🧹 Standardized error message using consistent message box formatting

### SumPagoNetoGerencia (v0.9.1)

- 🧠 Version bump for module tracking
- 🔧 Updated version header from `0.9.0` to `0.9.1` in preparation for upcoming changes

## UtilsModules

### UtilsCollections (v0.9.1)

- 🧠 Improved utility for working with collections
- ➕ Added `CollectionToArray` function to convert collections to arrays
- 🧹 Updated version and metadata header

### UtilsCoordinator (v0.9.1)

- 🧠 Enhanced functionality for coordinator-related operations
- ➕ Added `CreateCoordinatorTabs_newTabs` function to retrieve the new tabs collection
- ➕ Added `GetPromotersForCoordinator` function to retrieve a list of promoters for a given coordinator
- 🧹 Updated version and metadata header

### UtilsData (v0.9.1)

- 🧠 Improved functionality for handling sheet data and errors
- ➕ Added `StoreTotalInTargetCell` subroutine to store total sums in target cells
- 🧹 Refined error handling for `GetManagerPagoNeto` and `SumPagoNetoFromSheets`
- 🔧 Cleaned up code with additional handling for empty inputs

### UtilsErrorHandling (v0.9.1)

- 🧠 Added `HandleError` subroutine to centralize error handling
- 🔧 Displays error messages with optional source logging
- ➕ Provides an easy way to log errors with custom messages and sources

### UtilsManager (v0.9.1)

- 🧠 Version bump for module tracking
- 🔧 Updated version header from `0.9.0` to `0.9.1` in preparation for upcoming changes

### UtilsNumberToText (v0.9.1)

- 🧠 Version bump for module tracking
- 🔧 Updated version header from `0.9.0` to `0.9.1` in preparation for upcoming changes

### UtilsSheet (v0.9.1)

- 🧠 Introduced a new function for retrieving sheet names from a specified range.
- 🔧 Enhanced sheet validation with error messages for missing sheets.
- 🧹 Refined code to handle sheet name collection more efficiently.

### UtilsTable (v0.9.1)

- 🧠 Version bump for module tracking
- 🔧 Updated version header from `0.9.0` to `0.9.1` in preparation for upcoming changes

### UtilsValidation (v0.9.1)

- 🧠 Version bump for module tracking
- 🔧 Updated version header from `0.9.0` to `0.9.1` in preparation for upcoming changes

## Worksheet Modules

### Worksheet(Gerente) (v0.9.1)

- 🧠 Version bump for module tracking
- 🔧 Updated version header from `0.9.0` to `0.9.1` in preparation for upcoming changes

## Class Modules

### clsDictionary (v0.9.1)

- 🧠 Version bump for module tracking
- 🔧 Updated version header from `0.9.0` to `0.9.1` in preparation for upcoming changes
