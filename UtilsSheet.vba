'=========================================================
' Script: UtilsSheet
' Version: 0.9.2
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This module contains utility functions for working with sheets in Excel VBA. It includes functions for checking if a sheet exists, deleting unprotected tabs, and sanitizing sheet names. The module is designed to help manage the organization and naming of sheets in the workbook.
' and procedures in the workbook.
' Functions included in this module:
'   - SheetExists
'   - DeleteUnprotectedTabs
'   - SanitizeSheetName
'   - IsInNewTabs
'   - GetSheetNamesCollection
'   - RestoreSheetVisibility
'   - CreateNewTab
'=========================================================

'=========================================================
' Function: SheetExists
' Description:
'   Checks whether a worksheet with the specified name exists in the current workbook.
' Parameters:
'   - sheetName (Variant): The name of the worksheet to check for.
' Returns:
'   - True if the sheet exists, False otherwise.
' Notes:
'   - The check is case-insensitive.
'   - Suppresses runtime errors using On Error Resume Next.
'=========================================================
Public Function SheetExists(sheetName As Variant) As Boolean
    Dim sheetNameStr As String
    
    ' Convert the input to a string to avoid type mismatches
    On Error Resume Next
    sheetNameStr = CStr(sheetName)
    SheetExists = Not ThisWorkbook.Sheets(sheetNameStr) Is Nothing
    On Error GoTo 0
End Function

'=========================================================
' Function: DeleteUnprotectedTabs
' Description:
'   Deletes all unprotected tabs in the workbook, except for the active tab and any tabs specified in the protectedTabs array.
' Parameters:
'  - protectedTabs (Variant): An array of tab names to protect from deletion.
' Returns:
'  - None
' Notes:
'   - The function uses a loop to iterate through all sheets in the workbook.
'   - It checks if the tab name is in the protectedTabs array or if it is the active tab before deleting.
'=========================================================
Public Sub DeleteUnprotectedTabs(protectedTabs As Variant)
    Dim ws          As Worksheet
    Dim tabName     As String
    Dim activeTabName As String
    
    ' Get the name of the active tab
    activeTabName = ActiveSheet.Name
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        tabName = ws.Name
        
        ' Skip protected tabs and the active tab
        If tabName = activeTabName Then GoTo NextTab
        If IsInArray(tabName, protectedTabs) Then GoTo NextTab
        
        ' If not protected, delete the tab
        Application.DisplayAlerts = FALSE
        ws.Delete
        Application.DisplayAlerts = TRUE
        
        NextTab:
    Next ws
End Sub

'=========================================================
' Function: SanitizeSheetName
' Description:
'   This function sanitizes a given string to make it a valid Excel sheet name.
'   It replaces invalid characters with underscores and ensures the name
'   does not exceed 31 characters.
' Parameters:
'   - sheetName (Variant): The name to sanitize.
' Returns:
'   - String: The sanitized sheet name.
'=========================================================
Public Function SanitizeSheetName(sheetName As Variant) As String
    On Error Resume Next
    
    ' Ensure the input is treated as a string
    If IsError(sheetName) Or IsMissing(sheetName) Or IsNull(sheetName) Then
        sheetName = ""
    Else
        sheetName = CStr(sheetName)
    End If
    On Error GoTo 0
    
    ' Define invalid characters for sheet names
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", "?", "*", "[", "]", ":", "<", ">", "|")
    
    ' Replace invalid characters with underscores
    Dim i           As Long
    For i = LBound(invalidChars) To UBound(invalidChars)
        sheetName = Replace(sheetName, invalidChars(i), "_")
    Next i
    
    ' Ensure the name doesn't exceed 31 characters
    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If
    
    ' Return the sanitized sheet name
    SanitizeSheetName = sheetName
End Function

'=========================================================
' Function: IsInNewTabs
' Description:
'   This function checks if a given sheet name exists in a collection of newly created tabs.
'   It iterates through the collection and returns True if the sheet name is found; otherwise, it returns False.
' Parameters:
'   - sheetName (String): The name of the sheet to check.
'   - newTabs (Collection): A collection containing the names of newly created sheets.
' Returns:
'   - Boolean: True if the sheet is in the collection, otherwise False.
'=========================================================
Public Function IsInNewTabs(sheetName As String, newTabs As Collection) As Boolean
    Dim i           As Integer
    On Error Resume Next
    
    ' Check if the collection is initialized
    If newTabs Is Nothing Then
        IsInNewTabs = FALSE
        Exit Function
    End If
    
    ' Iterate through the collection to check for the sheet name
    For i = 1 To newTabs.Count
        If newTabs(i) = sheetName Then
            IsInNewTabs = TRUE
            Exit Function
        End If
    Next i
    
    ' Default to False if not found
    IsInNewTabs = FALSE
End Function

'=========================================================
' Function: GetSheetNamesCollection
' Description:
'    This function retrieves a collection of sheet names from a specified range in the active sheet.
' Parameters:
'   - nameRange (Range): The range containing the sheet names to check.
'   - targetSheet (Worksheet): The worksheet to check against.
'   - currentSheetIncluded (Boolean): A flag indicating if the current sheet is included in the collection.
'   - ERROR_SHEET_NOT_FOUND (String): The error message to display if a sheet is not found.
' Returns:
'   - Collection: A collection of valid sheet names.
' Notes:
'   - The function uses a loop to iterate through the specified range.
'   - It checks if each sheet name exists using the SheetExists function.
'   - If a sheet name does not exist, a message box is displayed and the function exits.
'=========================================================
Public Function GetSheetNamesCollection(nameRange As Range, targetSheet As Worksheet, ByRef currentSheetIncluded As Boolean) As Collection
    Dim sheetNamesCollection As New Collection
    Dim cell        As Range
    
    For Each cell In nameRange
        If cell.Value <> "" Then
            If Not SheetExists(cell.Value) Then
                MsgBox ERROR_SHEET_NOT_FOUND & cell.Value &        ' no existe.", vbExclamation, "Error"
                Exit Function
            End If
            sheetNamesCollection.Add cell.Value
            If UCase(cell.Value) = UCase(targetSheet.Name) Then
                currentSheetIncluded = TRUE
            End If
        End If
    Next cell
    
    Set GetSheetNamesCollection = sheetNamesCollection
End Function

'=========================================================
' Function: RestoreSheetVisibility
' Description:
'     This function restores the visibility of sheets based on their previous state.
' Parameters:
'   - sheetState (Collection): A collection containing the visibility state of each sheet.
'   - newTabs (Collection): A collection of newly created tabs that should remain visible.
' Returns:
'   - None
' Notes:
'   - The function iterates through all sheets in the workbook.
'   - It checks if each sheet is in the newTabs collection and restores its visibility accordingly.
'   - If a sheet is not in the newTabs collection, its visibility is set to the state stored in the sheetState collection.
'=========================================================
Public Sub RestoreSheetVisibility(sheetState As Collection, newTabs As Collection)
    Dim ws          As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If Not IsInNewTabs(ws.Name, newTabs) Then
            ws.Visible = sheetState(ws.Name)
        End If
    Next ws
End Sub

'=========================================================
' Function: CreateNewTab
' Description:
'     This function creates a new tab based on a template sheet.
'     It copies the template sheet to a new location, renames it, and sets its visibility.
' Parameters:
'   - templateSheet (Worksheet): The template sheet to copy from.
'   - tabName (Variant): The name for the new tab.
'   - placementAfter (Worksheet): The sheet after which the new tab will be placed.
'   - newTabs (Collection): A collection to track the newly created tabs.
' Returns:
'   - Worksheet: The newly created tab.
' Notes:
'   - The function copies the template sheet to a new location and renames it.
'   - It also sets the new tab to visible and adds it to the newTabs collection.
'=========================================================
Public Function CreateNewTab(templateSheet As Worksheet, tabName As Variant, placementAfter As Worksheet, newTabs As Collection) As Worksheet
    Dim newTab      As Worksheet
    Dim tabNameStr  As String
    
    ' Convert tabName to a string
    tabNameStr = CStr(tabName)
    
    ' Copy the template sheet to the specified location
    templateSheet.Copy After:=placementAfter
    
    ' Set the newly created sheet as newTab
    Set newTab = ThisWorkbook.Sheets(placementAfter.Index + 1)
    newTab.Name = tabNameStr
    
    ' Explicitly set the new tab to visible
    newTab.Visible = xlSheetVisible
    
    ' Track the newly created tab
    newTabs.Add newTab.Name
    
    ' Return the newly created tab
    Set CreateNewTab = newTab
End Function