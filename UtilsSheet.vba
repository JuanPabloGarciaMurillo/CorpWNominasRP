'==================================================
' Script: UtilsSheet
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This module contains utility functions for working with sheets in Excel VBA. It includes functions for checking if a sheet exists, deleting unprotected tabs, and sanitizing sheet names. The module is designed to help manage the organization and naming of sheets in the workbook.
' and procedures in the workbook.
' Functions included in this module:
'   - SheetExists
'   - DeleteUnprotectedTabs
'   - SanitizeSheetName
'   - IsInNewTabs
'==================================================

'====================================================
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
'====================================================
Public Function SheetExists(sheetName As Variant) As Boolean
    Dim sheetNameStr As String
    
    ' Convert the input to a string to avoid type mismatches
    On Error Resume Next
    sheetNameStr = CStr(sheetName)
    SheetExists = Not ThisWorkbook.Sheets(sheetNameStr) Is Nothing
    On Error GoTo 0
End Function

'====================================================
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
'====================================================
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

 '==================================================
' Function: SanitizeSheetName
' Description:
'   This function sanitizes a given string to make it a valid Excel sheet name.
'   It replaces invalid characters with underscores and ensures the name
'   does not exceed 31 characters.
' Parameters:
'   - sheetName (Variant): The name to sanitize.
' Returns:
'   - String: The sanitized sheet name.
'==================================================
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
    Dim i As Long
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

'==================================================
' Function: IsInNewTabs
' Description:
'   This function checks if a given sheet name exists in a collection of newly created tabs.
'   It iterates through the collection and returns True if the sheet name is found; otherwise, it returns False.
' Parameters:
'   - sheetName (String): The name of the sheet to check.
'   - newTabs (Collection): A collection containing the names of newly created sheets.
' Returns:
'   - Boolean: True if the sheet is in the collection, otherwise False.
'==================================================
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