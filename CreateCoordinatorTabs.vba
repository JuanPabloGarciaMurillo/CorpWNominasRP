'=======================================================================
' Subroutine: CreateCoordinatorTabs
' Version: 1.6.4
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This subroutine automates the process of creating coordinator-specific 
'   tabs in the workbook. It first gathers the necessary data from the 
'   "Coordinadores" table in the "Colaboradores" sheet. For each valid 
'   coordinator, it creates a new tab by copying a template sheet and 
'   renaming it according to the coordinator's name. The subroutine then 
'   populates the new tabs with relevant data from the "Coordinadores" table, 
'   including the coordinator's alias, and applies filters to include only 
'   the relevant data for each coordinator. It also copies common values 
'   (e.g., "razonSocial", "periodoDelPagoDel") to the new tabs.
' Parameters:
'   - None
' Returns:
'   - None
' Notes:
'   - This subroutine creates a new tab for each unique coordinator by 
'     copying a template and renaming it to the coordinator's name, ensuring 
'     the name is valid and doesn't exceed Excel's name length limitations.
'   - The coordinator names are sanitized to ensure they are valid sheet names.
'   - It applies a filter to the data based on the coordinator name and 
'     copies the filtered data to the newly created tab.
'   - The process includes sorting the coordinator names and copying shared 
'     values to the new sheets (e.g., "razonSocial", "periodoDelPagoDel").
'   - The subroutine also handles errors when no coordinators are found or 
'     if no matches are found for a coordinator's alias.
'=======================================================================

' Declare newTabs at the module level
Public newTabs As Collection
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

    Dim headerMapping As Object
    Dim header As String

    ' Set source and template sheets
    ' Set the source sheet as the active sheet where the button is clicked
    Set wsSource = ActiveSheet ' Now dynamically set to the active sheet
    Set templateSheet = ThisWorkbook.Sheets("Ejemplo Coordinacion") ' The sheet to be copied

    ' Create a dictionary to store the visibility status of sheets
    Set sheetState = CreateObject("Scripting.Dictionary")
    Set newTabs = New Collection ' Initialize collection to track new tabs
    Set headerMapping = CreateObject("Scripting.Dictionary") ' Initialize the header mapping dictionary
    Set coordDict = CreateObject("Scripting.Dictionary")

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
    
    ' Define the range for the "Coordinador" column (from row 9 to the last data row)
    Set coordColumn = wsSource.Range("A" & tableStartRow & ":A" & lastDataRow)

    ' Collect unique coordinator names (skip header row)
    ' Loop through the "Coordinador" column to collect unique values from the table only
    For Each cell In coordColumn
        coordName = Trim(cell.Value) ' Trim spaces from coordinator name

        ' Check if the cell has a valid coordinator name and it's not already in the dictionary
        If coordName <> "" And coordName <> "COORDINADOR" And Not coordDict.Exists(coordName) Then
            coordDict.Add coordName, Nothing ' Add unique coordinator to the dictionary
        End If
    Next cell

    ' Prevent errors if no coordinators are found
    If coordDict.Count = 0 Then
        MsgBox "No valid coordinators found.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Turn off screen updating and automatic calculation for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False

    ' Apply sorting to "Coordinador" column
    ' Sort the "Coordinador" column in ascending order (A-Z)
    tableObj.Sort.SortFields.Clear
    tableObj.Sort.SortFields.Add Key:=wsSource.Range("A" & tableStartRow & ":A" & lastDataRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    tableObj.Sort.Apply

    ' Define invalid characters for sheet names
    invalidChars = Array("\", "/", "?", "*", "[", "]")

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

    ' Copy common values to all tabs
    Dim razonSocial As Variant, periodoDelPagoDel As Variant
    Dim fechaDeExpedicion As Variant, periodoDelPagoAl As Variant
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
    Dim visibleCells As Range, newRow As ListRow

    For Each coordName In coordDict.Keys
        ' Apply filter for the current coordinator
        tableObj.Range.AutoFilter Field:=1, Criteria1:=coordName

        ' Get the corresponding worksheet for the coordinator
        On Error Resume Next
        Set visibleCells = tableObj.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not visibleCells Is Nothing Then
        Set newTab = ThisWorkbook.Sheets(coordName)
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
