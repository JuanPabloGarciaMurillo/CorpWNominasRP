'=========================================================
' Script: UtilsTable
' Version: 0.9.0
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'    This module contains utility functions for working with tables in Excel VBA. It includes functions for clearing table rows, sorting tables alphabetically, and populating tables with data from a collection.
' Functions included in this module:
'   - ClearTableRows
'   - SortTableAlphabetically
'   - PopulateTableWithCollection
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
' Function: PopulateTableWithCollection
' Description:
'   This function populates a given ListObject (table) in Excel with data from a Collection.
' Parameters:
'   - table (ListObject): The table to populate.
'   - data (Collection): The collection of data to add to the table.
'   - columnIndex (Optional Long): The index of the column to populate (default is 1).
' Returns:
'   - None
' Notes:
'   - The function iterates through the collection and adds each item as a new row in the table.
'   - It assumes the table has at least one column to insert data into.
'=========================================================
Public Sub PopulateTableWithCollection(ByVal table As ListObject, ByVal data As Collection, _
                                       Optional ByVal columnIndex As Long = 1)
    Dim item As Variant
    Dim newRow As ListRow
    For Each item In data
        Set newRow = table.ListRows.Add
        newRow.Range(1, columnIndex).Value = item
    Next item
End Sub