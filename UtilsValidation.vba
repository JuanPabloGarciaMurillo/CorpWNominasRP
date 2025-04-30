' Script: UtilsValidation
' Version: 0.9.3
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'    This module contains utility functions for applying validation rules in Excel VBA. It includes functions for getting a list of aliases based on a coordination value, clearing validation rules, and applying validation dynamically to a table.
' Functions included in this module:
'   - GetAliasList
'   - ClearValidation
'   - ApplyValidation
'   - ApplyDynamicValidation
'   - ProcessRow

' Function: GetAliasList
' Description:
'   This function retrieves a list of aliases based on a coordination value from a table.
'   It concatenates the aliases into a comma-separated string.
' Parameters:
'   - coordValue (String): The coordination value to filter by.
' Returns:
'   -  String: A comma-separated list of aliases for the specified coordination value.
' Notes:
'   - The function uses a ListObject (table) named "Colaboradores" in the "Colaboradores" sheet.
'   - It retrieves the COORDINACION and ALIAS columns from the table.
'   - It assumes the aliases are stored in a table named "Promotores" with columns "COORDINACION" and "ALIAS".

Public Function GetAliasList(ByVal coordValue As String) As String
    Dim tbl As ListObject
    Dim aliasList As String
    Dim coordColumnIndex As Long
    Dim aliasColumnIndex As Long
    Dim row As ListRow
    
    ' Set the table reference
    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets(COLABORADORES_SHEET).ListObjects(PROMOTORES_TABLE)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        GetAliasList = ""
        Exit Function
    End If
    
    ' Get the column indices for COORDINACION and ALIAS
    On Error Resume Next
    coordColumnIndex = tbl.ListColumns(COORDINACION_COLUMN).Index
    aliasColumnIndex = tbl.ListColumns(ALIAS_COLUMN).Index
    On Error GoTo 0
    
    If coordColumnIndex = 0 Or aliasColumnIndex = 0 Then
        GetAliasList = ""
        Exit Function
    End If
    
    ' Loop through the rows in the table and build the alias list
    aliasList = ""
    For Each row In tbl.ListRows
        If row.Range.Cells(1, coordColumnIndex).Value = coordValue Then
            aliasList = aliasList & row.Range.Cells(1, aliasColumnIndex).Value & ","
        End If
    Next row
    
    ' Remove the trailing comma
    If Len(aliasList) > 0 Then
        aliasList = Left(aliasList, Len(aliasList) - 1)
    End If
    
    GetAliasList = aliasList
End Function

' Sub: ClearValidation
' Description:
'   This function clears any validation rules applied to a specified range.
' Parameters:
'   - validationRange (Range): The range to clear validation from.
' Notes:
'   - The function uses On Error Resume Next to suppress errors if the range has no validation.
'   - It deletes the validation rules from the specified range.

Public Sub ClearValidation(ByVal validationRange As Range)
    If validationRange Is Nothing Then
        Exit Sub
    End If
    
    On Error Resume Next
    validationRange.Validation.Delete
    On Error GoTo 0
End Sub

' Sub: ApplyValidation
' Description:
'   This function applies a validation rule to a specified range based on a list of aliases.
' Parameters:
'   - validationRange (Range): The range to apply validation to.
'   - aliasList (String): A comma-separated list of aliases to use for validation.
' Notes:
'   - The function splits the alias list into an array and trims spaces.
'   - It applies a list validation rule to the specified range using the Join function.

Public Sub ApplyValidation(ByVal validationRange As Range, ByVal aliasList As String)
    If validationRange Is Nothing Then
        Exit Sub
    End If
    
    If aliasList = "" Then
        Exit Sub
    End If
    
    Dim aliasArray() As String
    Dim i           As Long
    
    ' Split alias list into array and trim spaces
    aliasArray = Split(aliasList, ",")
    For i = LBound(aliasArray) To UBound(aliasArray)
        aliasArray(i) = Trim(aliasArray(i))
    Next i
    
    ' Apply validation
    On Error Resume Next
    validationRange.Validation.Delete
    validationRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, Formula1:=Join(aliasArray, ",")
    On Error GoTo 0
End Sub

' Sub: ApplyDynamicValidation
' Description:
'   This subroutine applies dynamic validation to a table based on a value column and a validated column.
' Parameters:
'   - ws (Worksheet): The worksheet containing the table.
'   - tableName (String): The name of the table to apply validation to.
'   - valueColumnName (String): The name of the column containing the values to filter by.
'   - validatedColumnName (String): The name of the column to apply validation to.
' Notes:
'   - The subroutine loops through each row in the table and applies validation dynamically.

Public Sub ApplyDynamicValidation(ByVal ws As Worksheet, ByVal tableName As String, ByVal valueColumnName As String, ByVal validatedColumnName As String)
    Dim tbl         As ListObject
    Dim row         As ListRow
    Dim coordValue  As String
    Dim aliasList   As String
    Dim validationRange As Range
    Dim valueColumnIndex As Long
    Dim validatedColumnIndex As Long
    
    ' Set references
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Exit Sub
    End If
    
    On Error Resume Next
    valueColumnIndex = tbl.ListColumns(valueColumnName).Index
    validatedColumnIndex = tbl.ListColumns(validatedColumnName).Index
    On Error GoTo 0
    
    If valueColumnIndex = 0 Or validatedColumnIndex = 0 Then
        Exit Sub
    End If
    
    ' Loop through each row in the table
    For Each row In tbl.ListRows
        coordValue = row.Range.Cells(1, valueColumnIndex).Value        ' Get COORDINADOR value
        Set validationRange = row.Range.Cells(1, validatedColumnIndex)
        
        If coordValue = "" Then
            ' If value is empty, clear validation
            ClearValidation validationRange
        Else
            ' Get alias list
            aliasList = GetAliasList(coordValue)
            
            ' Apply validation if alias list is not empty
            If aliasList <> "" Then
                ApplyValidation validationRange, aliasList
            Else
                ClearValidation validationRange
            End If
        End If
    Next row
End Sub

' Sub: ProcessRow
' Description:
'   This subroutine processes a row in the table and applies validation based on the COORDINADOR value.
' Parameters:
'   - coordCell (Range): The cell containing the COORDINADOR value.
'   - promotorCell (Range): The cell to apply validation to.
' Notes:
'   - The subroutine checks if the COORDINADOR value is empty and clears validation if so.
'   - It retrieves the alias list for the COORDINADOR value and applies validation if the list is not empty.
'   - If the current validation matches the alias list, it skips applying validation.
'   - It uses On Error Resume Next to suppress errors when checking the current validation.
'   - It assumes the alias list is a comma-separated string.

Public Sub ProcessRow(coordCell As Range, promotorCell As Range)
    Dim aliasList   As String
    Dim currentValidation As String
    
    ' Skip processing if COORDINADOR is empty
    If coordCell.Value = "" Then
        ClearValidation promotorCell
        Exit Sub
    End If
    
    ' Get the alias list for the COORDINADOR value
    aliasList = GetAliasList(coordCell.Value)
    
    ' Check if the current validation matches the alias list
    On Error Resume Next
    currentValidation = promotorCell.Validation.Formula1
    On Error GoTo 0
    
    If aliasList = currentValidation Then
        Exit Sub        ' Validation is already correct
    End If
    
    ' Apply validation if alias list is not empty, otherwise clear validation
    If aliasList <> "" Then
        ApplyValidation promotorCell, aliasList
    Else
        ClearValidation promotorCell
    End If
End Sub