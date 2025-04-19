'=========================================================
' Script: UtilsTable
' Version: 0.9.2
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'    This module contains utility functions for working with tables in Excel VBA. It includes functions for clearing table rows, sorting tables alphabetically, and populating tables with data from a collection.
' Functions included in this module:
'   - ClearTableRows
'   - SortTableAlphabetically
'   - PopulateTableWithCollection
'   - PopulateTable
'=========================================================

'=========================================================
' Function: ClearTableRows
' Description:
'   This function clears all rows in a given ListObject (table) in Excel.
' Parameters:
'   - table (ListObject): The table to clear rows from.
'   - It uses the DataBodyRange property to delete the rows.
' Returns:
'   - None
' Notes:
'   - The function checks if the table has any rows before attempting to delete them.
'=========================================================
Public Sub ClearTableRows(ByVal table As ListObject)
    If table.ListRows.Count > 0 Then
        table.DataBodyRange.Delete
    End If
End Sub

'=========================================================
' Function: SortTableAlphabetically
' Description:
'   This function sorts a given ListObject (table) in Excel alphabetically
' Parameters:
'   - table (ListObject): The table to sort.
'   - columnIndex (Long): The index of the column to sort by (1-based index).
'   - sortOrder (Optional XlSortOrder): The order to sort the data (default is xlAscending).
'   - dataOption (Optional XlSortDataOption): The data option for sorting (default is xlSortNormal).
' Returns:
'   - None
' Notes:
'   - The function uses the Sort method of the ListObject to perform the sorting.
'   - It clears any existing sort fields before applying the new sort.
'=========================================================
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

'=========================================================
' Subroutine: PopulateTableWithCollection
' Description:
'   This subroutine populates a given ListObject (table) in Excel with data from a Collection.
' Parameters:
'   - table (ListObject): The table to populate.
'   - data (Collection): The collection of data to add to the table.
'   - columnIndex (Optional Long): The index of the column to populate (default is 1).
' Returns:
'   - None
' Notes:
'   - The subroutine iterates through the collection and adds each item as a new row in the table.
'   - It assumes the table has at least one column to insert data into.
'=========================================================
Public Sub PopulateTableWithCollection(ByVal table As ListObject, ByVal data As Collection, _
    Optional ByVal columnIndex As Long = 1)
    Dim item        As Variant
    Dim newRow      As ListRow
    For Each item In data
        Set newRow = table.ListRows.Add
        newRow.Range(1, columnIndex).Value = item
    Next item
End Sub

'=========================================================
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
' Returns:
'   - None
' Notes:
'   - The subroutine iterates through the visible cells in the specified range and populates the new table with data from the source table.
'   - It uses the headerMapping dictionary to determine the target column index for each header.
'   - It validates the row and column indices before populating the table.
'=========================================================
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