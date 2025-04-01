'=======================================================================
' Script: CreateCoordinatorTabs
' Version: 1.5.6
' Author: Juan Pablo Garcia Murillo
' Date: 03/30/2025
' Description: 
'   This script automates the creation of new worksheet tabs for each unique coordinator found in a source table. It extracts data from the active sheet, applies necessary transformations, and populates the corresponding coordinator tabs while maintaining a predefined template format.
'   The script ensures that data is not duplicated by clearing existing records in the coordinator-specific tables before inserting fresh data. Additionally, it manages sheet visibility, prevents invalid sheet names, and maintains structured header mappings for accurate data placement.
'
'   Key functionalities:
'   - Identify and extract unique coordinator names from the source table
'   - Create new tabs (or reuse existing ones) based on a predefined template
'   - Clear existing data in coordinator-specific tables before inserting new records
'   - Copy filtered data from the source sheet to the corresponding coordinator tab
'   - Paste static values (e.g., company details, payment periods) into designated cells
'   - Preserve original sheet visibility settings while temporarily unhiding necessary sheets
'   - Ensure valid sheet names by replacing special characters and truncating long names
'   - Optimize performance by disabling screen updates and calculation during execution
'
'   This script improves data organization by ensuring each coordinator has a dedicated worksheet with up-to-date, non-duplicated records.

'  
'=======================================================================

Sub CreateCoordinatorTabs()
    Dim wsSource As Worksheet
    Dim templateSheet As Worksheet
    Dim coordName As Variant
    Dim lastRow As Long
    Dim coordDict As Object
    Dim newTab As Worksheet
    Dim cell As Range
    Dim coordColumn As Range
    Dim tableStartRow As Long
    Dim tableObj As ListObject
    Dim lastDataRow As Long
    Dim invalidChars As Variant
    Dim i As Integer
    Dim ws As Worksheet
    Dim sheetState As Object ' To store the visibility status of each sheet
    Dim newTabs As Collection ' To track newly created tabs
    Dim razonSocial As Variant, periodoDelPagoDel As Variant, fechaDeExpedicion As Variant, periodoDelPagoAl As Variant ' Values to be copied from the active sheet
    Dim headerMapping As Object
    Dim header As String

    ' Set the source sheet as the active sheet where the button is clicked
    Set wsSource = ActiveSheet ' Now dynamically set to the active sheet
    Set templateSheet = ThisWorkbook.Sheets("Ejemplo Coordinacion") ' The sheet to be copied

    ' Create a dictionary to store the visibility status of sheets
    Set sheetState = CreateObject("Scripting.Dictionary")
    Set newTabs = New Collection ' Initialize collection to track new tabs
    Set headerMapping = CreateObject("Scripting.Dictionary") ' Initialize the header mapping dictionary
    
    ' Define header mapping between source and new tab, excluding "COMISION"
    headerMapping.Add "PROMOTOR", 1
    headerMapping.Add "CREDENCIAL", 2
    headerMapping.Add "NOMBRE DEL ALUMNO", 3
    ' Skipping "COMISION" (no entry for it)
    headerMapping.Add "PLANTEL", 5
    headerMapping.Add "CURSO", 6
    headerMapping.Add "GRUPO", 7
    ' Skipping "PAGO" (no entry for it)
    headerMapping.Add "FECHA", 9
    headerMapping.Add "TS PLANTEL", 10
    headerMapping.Add "TS CREDENCIAL", 11

    ' Unhide all sheets and store their original state (hidden or visible)
    For Each ws In ThisWorkbook.Sheets
        ' Store visibility state
        sheetState.Add ws.Name, ws.Visible
        ws.Visible = xlSheetVisible
    Next ws

    ' Set the table range using ListObjects (Excel table object)
    Set tableObj = wsSource.ListObjects(1) ' Assuming there's only one table in the sheet

    ' Define the start row for the table
    tableStartRow = 9 ' Table data starts from row 9 (skip header row)

    ' Get the last row of the table data (excluding Totals Row)
    lastDataRow = tableObj.ListRows.Count + tableStartRow - 1 - 1 ' Subtracting 1 to exclude the Totals Row
    
    ' Define the range for the "Coordinador" column (from row 9 to the last data row)
    Set coordColumn = wsSource.Range("A" & tableStartRow & ":A" & lastDataRow)

    ' Create a dictionary to hold unique coordinator names
    Set coordDict = CreateObject("Scripting.Dictionary")

    ' Loop through the "Coordinador" column to collect unique values from the table only
    For Each cell In coordColumn
        coordName = Trim(cell.Value) ' Trim spaces from coordinator name

        ' Check if the cell has a valid coordinator name and it's not already in the dictionary
        If coordName <> "" And Not coordDict.exists(coordName) Then
            coordDict.Add coordName, Nothing ' Add unique coordinator to the dictionary
        End If
    Next cell

    ' Turn off screen updating and automatic calculation for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False

    ' Apply the filter to the "Coordinador" column in the table (for the full table in the active sheet)
    tableObj.Range.AutoFilter Field:=1, Criteria1:="*" ' Ensure that filter is initially applied for all rows

    ' List of invalid characters for sheet names
    invalidChars = Array("\", "/", "?", "*", "[", "]")

    ' Sort the "Coordinador" column in ascending order (A-Z)
    tableObj.Sort.SortFields.Clear
    tableObj.Sort.SortFields.Add Key:=wsSource.Range("A" & tableStartRow & ":A" & lastDataRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    tableObj.Sort.Apply

    ' Iterate over unique coordinators and create new tabs
    For Each coordName In coordDict.Keys
        ' Sanitize coordinator name to be a valid sheet name
        For i = LBound(invalidChars) To UBound(invalidChars)
            coordName = Replace(coordName, invalidChars(i), "_")
        Next i

        ' Ensure the name doesn't exceed 31 characters
        If Len(coordName) > 31 Then
            coordName = Left(coordName, 31)
        End If

        ' Check if sheet already exists
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(coordName)
        On Error GoTo 0

        ' If sheet doesn't exist, create it by copying the template
        If ws Is Nothing Then
            ' Copy the template sheet to the end of the workbook
            templateSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

            ' Set the newly created sheet as newTab
            Set newTab = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            newTab.Name = coordName

            ' Explicitly set the new tab to visible right after creation
            newTab.Visible = xlSheetVisible

            ' Track the newly created tab
            newTabs.Add newTab.Name

            ' Ensure correct table reference in copied sheet
            Dim newTable As ListObject
            Set newTable = newTab.ListObjects(1) ' Get the table in the copied sheet
    
            ' Perform the lookup to get the corresponding "NOMBRE" for the "COORDINADOR"
            Dim aliasRange As Range
            Dim nameRange As Range
            Dim coordAlias As Variant
            Dim wsColaboradores As Worksheet
    
            ' Assuming "Colaboradores" is the sheet containing the coordinators table
            Set wsColaboradores = ThisWorkbook.Sheets("Colaboradores") ' Replace with actual sheet name
    
            ' Set the range for ALIAS and NOMBRE columns (the table's actual range)
            Set aliasRange = wsColaboradores.ListObjects("Coordinadores").ListColumns("ALIAS").DataBodyRange
            Set nameRange = wsColaboradores.ListObjects("Coordinadores").ListColumns("NOMBRE").DataBodyRange
    
            ' Perform lookup using Application.Match instead of WorksheetFunction.Lookup
            On Error Resume Next
            Dim matchRow As Long
            matchRow = Application.Match(coordName, aliasRange, 0) ' Find the row where the coordinator matches
            
            If Not IsError(matchRow) Then
                ' If a match is found, get the corresponding NOMBRE
                coordAlias = nameRange.Cells(matchRow, 1).Value
            Else
                coordAlias = "Unknown Coordinator" ' If no match, set default value
            End If
            On Error GoTo 0
    
            ' Paste the found coordinator name (or default) into cell B1 (merged B1:D1) in the new tab
            newTab.Range("B1:D1").Value = coordAlias ' Place the value in the merged range B1:D1
    
            ' No filter is applied to the new tab, only the active sheet table
        End If
    Next coordName

    ' Get the values to copy from the active sheet (B2, B3, B6, D3)
    razonSocial = wsSource.Range("B2").Value
    periodoDelPagoDel = wsSource.Range("B3").Value
    fechaDeExpedicion = wsSource.Range("B6").Value
    periodoDelPagoAl = wsSource.Range("D3").Value

    ' Now loop through each new tab and paste the values into the corresponding cells (B2, B3, B6, D3)
    For Each coordName In newTabs
        Set newTab = ThisWorkbook.Sheets(coordName)
        
        ' Paste the values into the corresponding cells
        newTab.Range("B2").Value = razonSocial
        newTab.Range("B3").Value = periodoDelPagoDel
        newTab.Range("B6").Value = fechaDeExpedicion
        newTab.Range("D3").Value = periodoDelPagoAl

        ' Auto-fit columns after pasting data
        newTab.Cells.EntireColumn.AutoFit
    Next coordName
'
    ' Loop through the filtered rows (only visible rows for each coordinator) and copy the filtered data
    For Each coordName In coordDict.Keys
        ' Apply filter for the current coordinator
        tableObj.Range.AutoFilter Field:=1, Criteria1:=coordName

        ' Get the corresponding worksheet for the coordinator
        Set newTab = ThisWorkbook.Sheets(coordName)

        ' Get the table in the new tab
        Set newTable = newTab.ListObjects(1) ' Assuming each sheet has only one table

        ' **Clear existing data in the table before inserting new rows**
        If newTable.ListRows.Count > 0 Then
            newTable.DataBodyRange.Delete ' This removes all rows, keeping headers
        End If

        ' Check if this is the first iteration (to overwrite the first row)
        Dim isFirstIteration As Boolean
        isFirstIteration = True

        ' Loop through the visible rows after the filter is applied
        For Each cell In wsSource.ListObjects(1).DataBodyRange.SpecialCells(xlCellTypeVisible).Columns(1).Cells
            ' Skip header row and rows where all columns except the first are empty
            If cell.Row >= tableStartRow And Not IsRowEmpty(wsSource, cell.Row) Then
                ' Get the newly added row (it will be the last row of the table)
                newTable.ListRows.Add
                Set newRow = newTable.ListRows(newTable.ListRows.Count)

                ' Loop through the columns in the source table and paste values into the corresponding columns of the new row
                For i = 1 To wsSource.ListObjects(1).ListColumns.Count
                    header = wsSource.ListObjects(1).ListColumns(i).Name
                    If headerMapping.exists(header) Then
                        ' Paste the value in the correct column of the new row
                        newRow.Range(1, headerMapping(header)).Value = wsSource.Cells(cell.Row, i).Value
                    End If
                Next i
            End If
        Next cell

        ' Auto-fit columns after inserting data
        newTab.Cells.EntireColumn.AutoFit
    Next coordName
'    
    ' Clean up
    Application.CutCopyMode = False
    wsSource.AutoFilterMode = False ' Reset filter mode in the source sheet

    ' Restore the original filter state in the active sheet
    tableObj.Range.AutoFilter Field:=1 ' Apply the original filter in the active sheet

    ' Restore screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Hide the sheets back to their original state, excluding new tabs
    For Each ws In ThisWorkbook.Sheets
        If Not IsInNewTabs(ws.Name, newTabs) Then
            ws.Visible = sheetState(ws.Name) ' Restore visibility based on the stored state
        End If
    Next ws
End Sub

' Function to check if a sheet is one of the newly created tabs
Function IsInNewTabs(sheetName As String, newTabs As Collection) As Boolean
    On Error GoTo NotFound
    Dim i As Integer
    For i = 1 To newTabs.Count
        If newTabs(i) = sheetName Then
            IsInNewTabs = True
            Exit Function
        End If
    Next i
    Exit Function
NotFound:
    IsInNewTabs = False
End Function

' Function to check if the row is empty (ignores only the first column)
Function IsRowEmpty(ws As Worksheet, rowNum As Long) As Boolean
    Dim col As Long
    Dim lastColumn As Long
    lastColumn = ws.ListObjects(1).ListColumns.Count
    
    ' Check all columns except for the first column
    For col = 2 To lastColumn ' Skip the first column (column 1)
        If ws.Cells(rowNum, col).Value <> "" Then
            IsRowEmpty = False
            Exit Function
        End If
    Next col
    
    ' If none of the columns except the first column have data, the row is considered empty
    IsRowEmpty = True
End Function
