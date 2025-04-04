'=======================================================================
' Script: CreatePromotorTabs
' Version: 1.6.2
' Author: Juan Pablo Garcia Murillo
' Date: 03/30/2025
' Description:
'   This script automates the creation of new worksheet tabs for each unique Promotor found in a source table. It extracts data from the active sheet, applies necessary transformations, and populates the corresponding Promotor tabs while maintaining a predefined template format.
'
'   The script ensures that data is not duplicated by clearing existing records in the Promotor-specific tables before inserting fresh data. Additionally, it manages sheet visibility, prevents invalid sheet names, maintains structured header mappings for accurate data placement, and performs an automatic lookup to fetch the full Promotor name from a reference sheet.
'=======================================================================

Sub CreatePromotorTabs()
    Dim wsSource As Worksheet
    Dim templateSheet As Worksheet
    Dim promotorName As Variant
    Dim lastRow As Long
    Dim promotorDict As Object
    Dim newTab As Worksheet
    Dim cell As Range
    Dim promotorColumn As Range
    Dim tableStartRow As Long
    Dim tableObj As ListObject
    Dim lastDataRow As Long
    Dim invalidChars As Variant
    Dim i As Integer
    Dim ws As Worksheet
    Dim sheetState As Object ' To store the visibility status of each sheet
    Dim newTabs As Collection ' To track newly created tabs
    Dim headerMapping As Object
    Dim header As String

    ' Set source and template sheets
    ' Set the source sheet as the active sheet where the button is clicked
    Set wsSource = ActiveSheet ' Now dynamically set to the active sheet
    Set templateSheet = ThisWorkbook.Sheets("Ejemplo Promotor") ' The sheet to be copied

    ' Create a dictionary to store the visibility status of sheets
    Set sheetState = CreateObject("Scripting.Dictionary")
    Set newTabs = New Collection ' Initialize collection to track new tabs
    Set headerMapping = CreateObject("Scripting.Dictionary") ' Initialize the header mapping dictionary
    ' Create a dictionary to hold unique promotor names
    Set promotorDict = CreateObject("Scripting.Dictionary")

    ' Define header mapping between source and new tab, excluding "COMISION" and "PAGO")
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
    lastDataRow = tableObj.ListRows.Count + tableStartRow - 1 ' Subtracting 1 to exclude the Totals Row
    
    ' Define the range for the "Promotor" column (from row 9 to the last data row)
    Set promotorColumn = wsSource.Range("A" & tableStartRow & ":A" & lastDataRow)

    ' Collect unique promotor names (skip header row)
    ' Loop through the "Promotor" column to collect unique values from the table only
    For Each cell In promotorColumn
        promotorName = Trim(cell.Value) ' Trim spaces from promotor name

        ' Check if the cell has a valid promotor name and it's not already in the dictionary
        If promotorName <> "" And promotorName <> "PROMOTOR" And Not promotorDict.exists(promotorName) Then
            promotorDict.Add promotorName, Nothing ' Add unique promotor to the dictionary
        End If
    Next cell

    ' Prevent errors if no promotors are found
    If promotorDict.Count = 0 Then
        MsgBox "No valid promotors found.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Turn off screen updating and automatic calculation for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False

    ' Apply sorting to "Promotor" column
    ' Sort the "Promotor" column in ascending order (A-Z)
    tableObj.Sort.SortFields.Clear
    tableObj.Sort.SortFields.Add Key:=wsSource.Range("A" & tableStartRow & ":A" & lastDataRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    tableObj.Sort.Apply

    ' Define invalid characters for sheet names
    invalidChars = Array("\", "/", "?", "*", "[", "]")
    ' Iterate over unique promotors and create new tabs
    For Each promotorName In promotorDict.Keys
        ' Sanitize promotor name to be a valid sheet name
        For i = LBound(invalidChars) To UBound(invalidChars)
            promotorName = Replace(promotorName, invalidChars(i), "_")
        Next i

        ' Ensure the name doesn't exceed 31 characters
        If Len(promotorName) > 31 Then
            promotorName = Left(promotorName, 31)
        End If

        ' Check if sheet already exists
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(promotorName)
        On Error GoTo 0

        ' If sheet doesn't exist, create it by copying the template
        If ws Is Nothing Then
            ' Copy the template sheet after the wsSource (the sheet with the button)
            templateSheet.Copy After:=wsSource
            ' Set the newly created sheet as newTab
            Set newTab = ThisWorkbook.Sheets(wsSource.Index + 1)
            newTab.Name = promotorName

            ' Explicitly set the new tab to visible right after creation
            newTab.Visible = xlSheetVisible

            ' Track the newly created tab
            newTabs.Add newTab.Name

            ' Ensure correct table reference in copied sheet
            Dim newTable As ListObject
            Set newTable = newTab.ListObjects(1) ' Get the table in the copied sheet
    
            ' Perform the lookup to get the corresponding "NOMBRE" for the "PROMOTOR"
            Dim aliasRange As Range
            Dim nameRange As Range
            Dim promotorAlias As Variant
            Dim wsColaboradores As Worksheet
    
            ' Assuming "Colaboradores" is the sheet containing the promotors table
            Set wsColaboradores = ThisWorkbook.Sheets("Colaboradores") ' Replace with actual sheet name
    
            ' Set the range for ALIAS and NOMBRE columns (the table's actual range)
            Set aliasRange = wsColaboradores.ListObjects("Promotores").ListColumns("ALIAS").DataBodyRange
            Set nameRange = wsColaboradores.ListObjects("Promotores").ListColumns("NOMBRE").DataBodyRange
    
            ' Perform lookup using Application.Match instead of WorksheetFunction.Lookup
            On Error Resume Next
            Dim matchRow As Long
            matchRow = Application.Match(promotorName, aliasRange, 0) ' Find the row where the promotor matches
            
            If Not IsError(matchRow) Then
                ' If a match is found, get the corresponding NOMBRE
                promotorAlias = nameRange.Cells(matchRow, 1).Value
            Else
                promotorAlias = "Unknown Promotor" ' If no match, set default value
            End If
            On Error GoTo 0
    
            ' Paste the found promotor name (or default) into cell B1 (merged B1:D1) in the new tab
            newTab.Range("B1:D1").Value = promotorAlias ' Place the value in the merged range B1:D1
    
            ' No filter is applied to the new tab, only the active sheet table
        End If
    Next promotorName

    ' Copy common values to all tabs
    Dim razonSocial As Variant, periodoDelPagoDel As Variant
    Dim fechaDeExpedicion As Variant, periodoDelPagoAl As Variant
    ' Get the values to copy from the active sheet (B2, B3, B6, D3)
    razonSocial = wsSource.Range("B2").Value
    periodoDelPagoDel = wsSource.Range("B3").Value
    fechaDeExpedicion = wsSource.Range("B6").Value
    periodoDelPagoAl = wsSource.Range("D3").Value

    ' Now loop through each new tab and paste the values into the corresponding cells (B2, B3, B6, D3)
    For Each promotorName In newTabs
        Set newTab = ThisWorkbook.Sheets(promotorName)
        
        ' Paste the values into the corresponding cells
        newTab.Range("B2").Value = razonSocial
        newTab.Range("B3").Value = periodoDelPagoDel
        newTab.Range("B6").Value = fechaDeExpedicion
        newTab.Range("D3").Value = periodoDelPagoAl

        ' Auto-fit columns after pasting data
        newTab.Cells.EntireColumn.AutoFit
    Next promotorName
'
    ' Loop through the filtered rows (only visible rows for each promotor) and copy the filtered data
    Dim visibleCells As Range, newRow As ListRow

    For Each promotorName In promotorDict.Keys
        ' Apply filter for the current Promotor
        tableObj.Range.AutoFilter Field:=1, Criteria1:=promotorName

        ' Get the corresponding worksheet for the promotor
        On Error Resume Next
        Set visibleCells = tableObj.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not visibleCells Is Nothing Then
        Set newTab = ThisWorkbook.Sheets(promotorName)
        Set newTable = newTab.ListObjects(1)
          If newTable.ListRows.Count > 0 Then newTable.DataBodyRange.Delete

            ' Loop through filtered rows
            For Each cell In visibleCells.Columns(1).Cells
            If cell.Row >= tableStartRow And Not IsRowEmpty(wsSource, cell.Row) Then
                    Set newRow = newTable.ListRows.Add
                    For i = 1 To tableObj.ListColumns.Count
                        header = tableObj.ListColumns(i).Name
                        If headerMapping.Exists(header) Then
                        newRow.Range(1, headerMapping(header)).Value = wsSource.Cells(cell.Row, i).Value
                    End If
                Next i
            End If
        Next cell
        ' Auto-fit columns after inserting data
        newTab.Cells.EntireColumn.AutoFit
        End If
    Next promotorName
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