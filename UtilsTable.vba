' Script: UtilsTable
' Version: 0.9.3
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'    This module contains utility functions for working with tables in Excel VBA. It includes functions for clearing table rows, sorting tables alphabetically, and populating tables with data from a collection.
' Functions included in this module:
'   - ClearTableRows
'   - SortTableAlphabetically
'   - PopulateTableWithCollection
'   - PopulateTable
'   - SetAliasInNewTab

' Function: ClearTableRows
' Description:
'   This function clears all rows in a given ListObject (table) in Excel.
' Parameters:
'   - table (ListObject): The table to clear rows from.
'   - It uses the DataBodyRange property to delete the rows.
' Notes:
'   - The function checks if the table has any rows before attempting to delete them.

Public Sub ClearTableRows(ByVal table As ListObject)
    If table.ListRows.Count > 0 Then
        table.DataBodyRange.Delete
    End If
End Sub

' Function: SortTableAlphabetically
' Description:
'   This function sorts a given ListObject (table) in Excel alphabetically
' Parameters:
'   - table (ListObject): The table to sort.
'   - columnIndex (Long): The index of the column to sort by (1-based index).
'   - sortOrder (Optional XlSortOrder): The order to sort the data (default is xlAscending).
'   - dataOption (Optional XlSortDataOption): The data option for sorting (default is xlSortNormal).
' Notes:
'   - The function uses the Sort method of the ListObject to perform the sorting.
'   - It clears any existing sort fields before applying the new sort.

Public Sub SortTableAlphabetically(ByVal table As ListObject, ByVal columnIndex As Long, _
    Optional ByVal sortOrder As XlSortOrder = xlAscending, _
    Optional ByVal dataOption As XlSortDataOption = xlSortNormal)
    With table.Sort
        .SortFields.Clear
        .SortFields.Add Key:=table.ListColumns(columnIndex).DataBodyRange, _
                        SortOn:=xlSortOnValues, Order:=sortOrder, DataOption:=dataOption
        .Header = xlYes
        .Apply
    End With
End Sub

' Subroutine: PopulateTableWithCollection
' Description:
'   This subroutine populates a given ListObject (table) in Excel with data from a Collection.
' Parameters:
'   - table (ListObject): The table to populate.
'   - data (Collection): The collection of data to add to the table.
'   - columnIndex (Optional Long): The index of the column to populate (default is 1).
' Notes:
'   - The subroutine iterates through the collection and adds each item as a new row in the table.
'   - It assumes the table has at least one column to insert data into.

Public Sub PopulateTableWithCollection(ByVal table As ListObject, ByVal data As Collection, _
    Optional ByVal columnIndex As Long = 1)
    Dim item        As Variant
    Dim newRow      As ListRow
    For Each item In data
        Set newRow = table.ListRows.Add
        newRow.Range(1, columnIndex).Value = item
    Next item
End Sub

' Subroutine: PopulateTable
' Description:
'    This subroutine populates a given ListObject (table) in Excel with data from a specified range.
' Parameters:
'   - newTable (ListObject): The table to populate.
'   - visibleCells (Range): The range of cells to populate the table with.
'   - tableObj (ListObject): The source table object to get data from.
'   - headerMapping (clsDictionary): A dictionary mapping headers to their respective column indices.
'   - tableStartRow (Long): The starting row for the table.
'   - wsSource (Worksheet): The worksheet containing the source data.
' Notes:
'   - The subroutine iterates through the visible cells in the specified range and populates the new table with data from the source table.
'   - It uses the headerMapping dictionary to determine the target column index for each header.
'   - It validates the row and column indices before populating the table.

Public Sub PopulateTable(newTable As ListObject, visibleCells As Range, tableObj As ListObject, headerMapping As clsDictionary, tableStartRow As Long, wsSource As Worksheet)
    Dim newRow      As ListRow
    Dim header      As String
    Dim targetColumnIndex As Long
    Dim cell        As Range
    Dim i           As Integer
    
    ' Loop through the visible cells
    For Each cell In visibleCells.Columns(1).Cells
        ' Validate the row
        If cell.Row >= tableStartRow And Not IsRowEmpty(wsSource, cell.Row) Then
            ' Add a new row to the table
            Set newRow = newTable.ListRows.Add
            If newRow Is Nothing Then
                MsgBox "Failed To add a New row To the target table.", vbCritical
                Exit Sub
            End If
            
            ' Populate the row with data
            For i = 1 To tableObj.ListColumns.Count
                header = tableObj.ListColumns(i).Name
                
                ' Skip headers that are not in the headerMapping dictionary
                If Not headerMapping.Exists(header) Then
                    GoTo NextHeader
                End If
                
                ' Validate the target column index
                targetColumnIndex = headerMapping.GetValue(header)
                If targetColumnIndex > 0 And targetColumnIndex <= newTable.ListColumns.Count Then
                    newRow.Range(1, targetColumnIndex).Value = wsSource.Cells(cell.Row, i).Value
                Else
                    MsgBox "Invalid column index For header: " & header & ", Index: " & targetColumnIndex, vbCritical
                End If
                
                NextHeader:
            Next i
        End If
    Next cell
End Sub

' Subroutine: SetAliasInNewTab
' Description:
'     Sets the alias (e.g., Coordinator or Promotor name) in a new tab
'     by performing a lookup in a specified table.
' Parameters:
'   - newTab (Worksheet): The worksheet where the alias will be set.
'   - tableName (String): The name of the table to perform the lookup in.
'   - aliasColumn (String): The column containing the lookup keys (e.g., Coordinator/Promotor names).
'   - nameColumn (String): The column containing the values to retrieve (e.g., aliases).
'   - lookupValue (Variant): The value to look up in the alias column.
'   - defaultAlias (String): The default value to use if no match is found.
' Notes:
'   - The function uses Application.Match to find the row index of the lookup value.
'   - If a match is found, it retrieves the corresponding value from the name column.
'   - If no match is found, it sets the default alias value.
'   - The alias value is set in the merged range B1:D1 of

Public Sub SetAliasInNewTab(newTab As Worksheet, tableName As String, aliasColumn As String, nameColumn As String, lookupValue As Variant, defaultAlias As String)
    Dim wsColaboradores As Worksheet
    Dim aliasRange As Range
    Dim nameRange As Range
    Dim matchRow As Variant
    Dim aliasValue As Variant
    
    ' Convert lookupValue to a string
    Dim lookupValueStr As String
    lookupValueStr = CStr(lookupValue)
    
    ' Assuming "Colaboradores" is the sheet containing the table
    Set wsColaboradores = ThisWorkbook.Sheets("Colaboradores")
    
    ' Set the range for ALIAS and NOMBRE columns (the table's actual range)
    Set aliasRange = wsColaboradores.ListObjects(tableName).ListColumns(aliasColumn).DataBodyRange
    Set nameRange = wsColaboradores.ListObjects(tableName).ListColumns(nameColumn).DataBodyRange
    
    ' Perform lookup using Application.Matche
    On Error Resume Next
    matchRow = Application.Match(lookupValueStr, aliasRange, 0)
    If Not IsError(matchRow) Then
        ' Retrieve the value from the matched row in the nameRange
        aliasValue = nameRange.Cells(matchRow, 1).Value
    Else
        ' Default value if no match is found
        aliasValue = defaultAlias
    End If
    On Error GoTo 0
    
    ' Set the alias value in B1 (merged B1:D1)
    newTab.Range("B1:D1").Value = aliasValue
End Sub

' Subroutine: SetTableName
' Description:
'      Sets the name of a ListObject (table) in Excel to a new name based on the provided entity name and prefix.
' Parameters:
'   - newTable (ListObject): The table to rename.
'   - entityName (Variant): The name of the entity to use for the new table name.
'   - prefix (String): The prefix to prepend to the new table name.
' Notes:
'   - The function sanitizes the entity name to ensure it is a valid Excel table name.
'   - It replaces spaces with underscores and appends the prefix.
'   - The function uses On Error Resume Next to handle any errors that may occur during the renaming process.
'   - The new table name is set to the sanitized entity name with the prefix.

Public Sub SetTableName(newTable As ListObject, entityName As Variant, prefix As String)
    Dim sanitizedName As String
    Dim newTableName As String
    
    ' Convert entityName to a string
    Dim entityNameStr As String
    entityNameStr = CStr(entityName)
    
    ' Sanitize the entity name
    sanitizedName = SanitizeSheetName(entityNameStr)
    
    ' Create the new table name
    newTableName = prefix & Replace(sanitizedName, " ", "_")
    
    ' Change the table name
    On Error Resume Next
    newTable.Name = newTableName
    On Error GoTo 0
End Sub