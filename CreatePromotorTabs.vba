'=======================================================================
' Subroutine: CreatePromotorTabs
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This subroutine automates the process of creating individual tabs for each
'   promotor in the workbook. It collects unique promotor names from the
'   "Promotores" table in the "Colaboradores" sheet and checks if each
'   promotor has a corresponding entry in the "Sueldos_Base" table in the
'   "Tabuladores" sheet. If the promotor exists in the "Sueldos_Base" table,
'   the subroutine creates a new sheet for that promotor by copying a template
'   sheet and renaming it according to the promotor's name. The subroutine
'   then populates the new tabs with relevant data from the "Promotores" table
'   and applies filters to include only the relevant data for each promotor.
'   Common values from the source sheet (e.g., "razonSocial", "periodoDelPagoDel")
'   are also copied into all the new tabs.
' Parameters:
'   - None
' Returns:
'   - None
' Notes:
'   - The subroutine creates a new tab for each promotor by copying a template
'     and renaming it to the promotor's name, ensuring the name is valid.
'   - It first checks for each promotor's base salary entry in the "Sueldos_Base"
'     table and only creates tabs for those with a valid match.
'   - The promotor names are sanitized to ensure they are valid sheet names
'     (e.g., replacing invalid characters with underscores).
'   - The subroutine automatically sorts the "Promotor" column in ascending order
'     before processing the rows.
'   - If no valid promotors are found, an error message is displayed.
'   - Filters are applied to each promotor's data, and new tabs are populated
'     with filtered rows.
'   - After creating the tabs, common values from the source sheet are pasted
'     into corresponding cells in each new tab.
'   - The original visibility state of the sheets is preserved, with only the
'     newly created tabs visible.
'=======================================================================


Public Sub CreatePromotorTabs()
    On Error GoTo ErrHandler
    Debug.Print "Entering CreatePromotorTabs..."
    Dim wsSource    As Worksheet
    Dim templateSheet As Worksheet
    Dim promotorName As Variant
    Dim lastRow     As Long
    Dim promotorDict As Object
    Dim newTab      As Worksheet
    Dim cell        As Range
    Dim promotorColumn As Range
    Dim tableStartRow As Long
    Dim tableObj    As ListObject
    Dim lastDataRow As Long
    Dim invalidChars As Variant
    Dim i           As Integer
    Dim ws          As Worksheet
    Dim sheetState  As Object
    Dim newTabs     As Collection
    Dim headerMapping As Object
    Dim header      As String
    Dim wsColaboradores As Worksheet
    Dim wsTabuladores As Worksheet
    Dim promotorTable As ListObject
    Dim baseSalaryTable As ListObject
    Dim promotorCoord As String
    Dim promotorRow As ListRow
    Dim matchRow    As Long
    Dim colNombreIndex As Long, colCoordIndex As Long
    Dim promotorTableFiltered As ListObject
    Dim tabuladorRow As ListRow
    Dim promotorMatch As Boolean
    Dim coordinatorName As String
    
    ' Set the source sheet as the active sheet where the button is clicked
    Set wsSource = ActiveSheet
    Set templateSheet = ThisWorkbook.Sheets("Ejemplo Promotor")
    
    Set wsColaboradores = ThisWorkbook.Sheets("Colaboradores")
    Set wsTabuladores = ThisWorkbook.Sheets("Tabuladores")
    
    ' Create a dictionary to store the visibility status of sheets
    Set sheetState = CreateObject("Scripting.Dictionary")
    Set newTabs = New Collection
    Set headerMapping = CreateObject("Scripting.Dictionary")
    
    ' Create a dictionary to hold unique promotor names
    Set promotorDict = CreateObject("Scripting.Dictionary")
    
    ' Set the tables
    Set promotorTable = wsColaboradores.ListObjects("Promotores")
    Set baseSalaryTable = wsTabuladores.ListObjects("Sueldos_Base")
    
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
    Set tableObj = wsSource.ListObjects(1)
    
    ' Define the start row for the table
    tableStartRow = 9
    
    ' Get the last row of the table data (excluding Totals Row)
    lastDataRow = tableObj.ListRows.Count + tableStartRow - 1
    
    ' Define the range for the "Promotor" column (from row 9 to the last data row)
    Set promotorColumn = wsSource.Range("A" & tableStartRow & ":A" & lastDataRow)
    
    Debug.Print "Entering Base Salary promotor sheet functionality"
    ' Loop through the filtered Promotores and find Tabulador data
    coordinatorName = wsSource.Name
    For Each promotorRow In promotorTable.ListRows
        
        baseSalaryPromotorName = promotorRow.Range.Cells(1, promotorTable.ListColumns("NOMBRE").Index).Value
        baseSalaryPromotorAlias = promotorRow.Range.Cells(1, promotorTable.ListColumns("ALIAS").Index).Value
        promotorCoord = promotorRow.Range.Cells(1, promotorTable.ListColumns("COORDINACION").Index).Value
        
        If Trim(UCase(promotorCoord)) = Trim(UCase(coordinatorName)) Then
            
            ' Search for the Promotor in the Tabuladores table (Sueldos_Base)
            promotorMatch = False
            For Each tabuladorRow In baseSalaryTable.ListRows
                ' Check if Promotor matches Tabulador COLABORADOR
                If tabuladorRow.Range.Cells(1, baseSalaryTable.ListColumns("COLABORADOR").Index).Value = baseSalaryPromotorName Then
                    promotorMatch = True
                    promotorDict.Add baseSalaryPromotorAlias, Nothing
                    Exit For
                End If
            Next tabuladorRow
            
        End If
    Next promotorRow
    
    Debug.Print "Exiting CreateBaseSalaryTabsIfMissing..."
    
    ' Loop through the "Promotor" column to collect unique values from the table only
    For Each cell In promotorColumn
        promotorName = Trim(cell.Value)
        
        ' Check if the cell has a valid promotor name and it's not already in the dictionary
        If promotorName <> "" And promotorName <> "PROMOTOR" And Not promotorDict.Exists(promotorName) Then
            promotorDict.Add promotorName, Nothing
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
            Set newTable = newTab.ListObjects(1)
            Dim newTableName As String
            newTableName = "Tabla_Promotor" & Replace(promotorName, " ", "_")
            On Error Resume Next
            newTable.Name = newTableName
            On Error GoTo 0
            
            ' Perform the lookup to get the corresponding "NOMBRE" for the "PROMOTOR"
            Dim aliasRange As Range
            Dim nameRange As Range
            Dim promotorAlias As Variant
            
            ' Assuming "Colaboradores" is the sheet containing the promotors table
            Set wsColaboradores = ThisWorkbook.Sheets("Colaboradores")
            
            ' Set the range for ALIAS and NOMBRE columns (the table's actual range)
            Set aliasRange = wsColaboradores.ListObjects("Promotores").ListColumns("ALIAS").DataBodyRange
            Set nameRange = wsColaboradores.ListObjects("Promotores").ListColumns("NOMBRE").DataBodyRange
            
            ' Perform lookup using Application.Match instead of WorksheetFunction.Lookup
            On Error Resume Next
            matchRow = Application.Match(promotorName, aliasRange, 0)
            
            If Not IsError(matchRow) Then
                ' If a match is found, get the corresponding NOMBRE
                promotorAlias = nameRange.Cells(matchRow, 1).Value
            Else
                promotorAlias = "Unknown Promotor"
            End If
            On Error GoTo 0
            
            ' Paste the found promotor name (or default) into cell B1 (merged B1:D1) in the new tab
            newTab.Range("B1:D1").Value = promotorAlias
            
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

    ' Loop through the filtered rows (only visible rows for each promotor) and copy the filtered data
    Dim visibleCells As Range, newRow As ListRow
    
    For Each promotorName In promotorDict.Keys
        ' Always clear visibleCells before applying the filter
        Set visibleCells = Nothing
        
        ' Apply filter for the current Promotor
        tableObj.Range.AutoFilter Field:=1, Criteria1:=promotorName
        
        ' Get the corresponding worksheet for the promotor
        On Error Resume Next
        Set visibleCells = tableObj.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        ' Clear the previous data in the new tab before populating (no matter if data is found or not)
        If Not visibleCells Is Nothing Then
            Set newTab = ThisWorkbook.Sheets(promotorName)
            Set newTable = newTab.ListObjects(1)
            
            ' Always clear the target table to prevent leftover data from previous iterations
            If newTable.ListRows.Count > 0 Then
                newTable.DataBodyRange.Delete
            End If
            
            ' Only loop through the filtered rows if there are visible cells
            If Not visibleCells Is Nothing Then
                ' Loop through filtered rows
                For Each cell In visibleCells.Columns(1).Cells
                    If cell.row >= tableStartRow And Not IsRowEmpty(wsSource, cell.row) Then
                        Set newRow = newTable.ListRows.Add
                        For i = 1 To tableObj.ListColumns.Count
                            header = tableObj.ListColumns(i).Name
                            If headerMapping.Exists(header) Then
                                newRow.Range(1, headerMapping(header)).Value = wsSource.Cells(cell.row, i).Value
                            End If
                        Next i
                    End If
                Next cell
                
                ' Auto-fit columns after inserting data
                newTab.Cells.EntireColumn.AutoFit
            End If
        End If
        
        ' Reset the filter after processing the current promotor
        If tableObj.AutoFilter.FilterMode Then
            tableObj.AutoFilter.ShowAllData
        End If
    Next promotorName
    
    ' Clean up
    Application.CutCopyMode = False
    wsSource.AutoFilterMode = False
    
    ' Restore the original filter state in the active sheet
    tableObj.Range.AutoFilter Field:=1
    
    ' Restore screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Hide the sheets back to their original state, excluding new tabs
    For Each ws In ThisWorkbook.Sheets
        If Not IsInNewTabs(ws.Name, newTabs) Then
            ws.Visible = sheetState(ws.Name)
        End If
    Next ws
    
ErrHandler:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "CreatePromotorTabs"
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
End Sub



