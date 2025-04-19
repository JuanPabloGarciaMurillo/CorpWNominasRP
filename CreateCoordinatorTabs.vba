'=========================================================
' Subroutine: CreateCoordinatorTabs
' Version: 0.9.2
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This subroutine automates the process of creating coordinator-specific tabs in the workbook. It first gathers the necessary data from the "Coordinadores" table in the "Colaboradores" sheet. For each valid coordinator, it creates a new tab by copying a template sheet and renaming it according to the coordinator's name. The subroutine then populates the new tabs with relevant data from the "Coordinadores" table, including the coordinator's alias, and applies filters to include only the relevant data for each coordinator. It also copies common value (e.g., "razonSocial", "periodoDelPagoDel") to the new tabs.
' Parameters:
'   - None
' Returns:
'   - None
' Notes:
'   - This subroutine creates a new tab for each unique coordinator by copying a template and renaming it to the coordinator's name, ensuring the name is valid and doesn't exceed Excel's name length limitations.
'   - The coordinator names are sanitized to ensure they are valid sheet names.
'   - It applies a filter to the data based on the coordinator name and copies the filtered data to the newly created tab.
'   - The process includes sorting the coordinator names and copying shared values to the new sheets (e.g., "razonSocial", "periodoDelPagoDel").
'   - The subroutine also handles errors when no coordinators are found or if no matches are found for a coordinator's alias.
'=========================================================

' Declare newTabs at the module level
Public newTabs      As Collection
Sub CreateCoordinatorTabs()
    On Error GoTo ErrHandler
    OptimizeApplicationSettings
    
    Dim wsSource    As Worksheet
    Dim templateSheet As Worksheet
    Dim newTab      As Worksheet
    Dim cell        As Range
    Dim tableStartRow As Long
    Dim tableObj    As ListObject
    Dim lastDataRow As Long
    Dim i           As Integer
    Dim ws          As Worksheet
    Dim idx         As Integer
    Dim header      As String
    Dim headerMapping As clsDictionary
    Dim headersArray     As Variant
    Dim sheetState  As Collection
    Dim wsColaboradores As Worksheet
    Dim columnIndex As Variant
    Dim razonSocial     As Variant
    Dim periodoDelPagoDel As Variant
    Dim fechaDeExpedicion As Variant
    Dim periodoDelPagoAl As Variant
    Dim coordName   As Variant
    Dim coordKeys   As Collection
    Dim gerenteNombre As String
    Dim gerenteAlias As String
    Dim gerentesTbl As ListObject
    Dim coordTbl    As ListObject
    Dim foundRow    As Range
    Dim iRow        As ListRow
    Dim coordAlias  As Variant
    Dim coordColumnIndex As Long
    Dim aliasColumnIndex As Long
    Dim gerenciaColumnIndex As Long
    Dim key             As Variant
    Dim uniqueKeys  As Object
    
    ' Initialize dictionaries using the custom dictionary class
    Set headerMapping = New clsDictionary
    Set uniqueKeys = New clsDictionary
    
    ' Set the source sheet as the active sheet where the button is clicked
    Set wsSource = ActiveSheet
    Set templateSheet = ThisWorkbook.Sheets(COORDINADORES_SHEET)
    Set wsColaboradores = ThisWorkbook.Sheets(COLABORADORES_SHEET)
    ' Set the tables
    Set gerentesTbl = wsColaboradores.ListObjects(GERENTES_TABLE)
    Set coordTbl = wsColaboradores.ListObjects(COORDINADORES_TABLE)
    
    ' Create a dictionary to store the visibility status of sheets
    Set sheetState = New Collection
    Set newTabs = New Collection
    Set coordKeys = New Collection
    
    ' Set the table range using ListObjects (Excel table object)
    Set tableObj = wsSource.ListObjects(1)
    
    ' Define headers and their corresponding column indices
    headersArray = Split(HEADERS, ",")
    columnIndex = Split(COLUMN_INDICES, ",")
    
    ' Loop through the headers and add them to the dictionary
    For idx = LBound(headersArray) To UBound(headersArray)
        If Not headerMapping.Exists(CStr(headersArray(idx))) Then
            headerMapping.Add CStr(headersArray(idx)), columnIndex(idx)
        End If
    Next idx
    
    ' Unhide all sheets and store their original state (hidden or visible)
    For Each ws In ThisWorkbook.Sheets
        ' Ensure ws.Name is unique and ws.Visible is valid
        If Not headerMapping.Exists(ws.Name) Then
            sheetState.Add ws.Visible, ws.Name
        Else
        End If
        ws.Visible = xlSheetVisible
    Next ws
    
    ' Define the start row for the table
    tableStartRow = 9
    
    ' Get the last row of the table data (excluding Totals Row)
    lastDataRow = tableObj.ListRows.Count + tableStartRow - 1
    
    ' Dynamically retrieve column indices for "ALIAS" and "GERENCIA"
    aliasColumnIndex = coordTbl.ListColumns(ALIAS_COLUMN).Index
    gerenciaColumnIndex = coordTbl.ListColumns(GERENCIA_COLUMN).Index
    
    ' Get manager's name from B1:D1
    gerenteNombre = Trim(wsSource.Range("B1").Value)
    
    ' Find the alias for the manager
    On Error Resume Next
    Set foundRow = gerentesTbl.ListColumns(NOMBRE_COLUMN).DataBodyRange.Find(What:=gerenteNombre, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not foundRow Is Nothing Then
        gerenteAlias = foundRow.Offset(0, 1).Value
    Else
        MsgBox "El Gerente        '" & gerenteNombre & "' no se encuentra en la tabla de Gerentes.", vbExclamation
        Exit Sub
    End If
    
    ' Add coordinators from the "Coordinadores" table where GERENCIA = gerenteAlias
    For Each iRow In coordTbl.ListRows
        If Trim(iRow.Range(1, gerenciaColumnIndex).Value) = gerenteAlias Then
            coordAlias = Trim(CStr(iRow.Range(1, aliasColumnIndex).Value))
            If CStr(coordAlias) <> "" And Not uniqueKeys.Exists(CStr(coordAlias)) Then
                uniqueKeys.Add CStr(coordAlias), CStr(coordAlias)
            End If
        End If
    Next iRow
    
    ' Add unique coordinators from uniqueKeys to coordKeys
    For Each key In uniqueKeys.GetKeys
        coordKeys.Add key
    Next key
    
    ' Prevent errors if no coordinators are found
    If coordKeys.Count = 0 Then
        MsgBox ERROR_NO_VALID_COORDINATOR, vbExclamation, "Error"
        Exit Sub
    End If
    
    If tableObj.ListRows.Count = 0 Then
        Exit Sub
    End If
    
    coordColumnIndex = tableObj.ListColumns(COORDINADOR_COLUMN).Index
    
    ' Sort the "Coordinador" column in ascending order (A-Z)
    tableObj.Sort.SortFields.Clear
    tableObj.Sort.SortFields.Add key:=tableObj.ListColumns(coordColumnIndex).Range, _
                                 SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    tableObj.Sort.Apply
    
    ' Iterate over unique coordinators and create new tabs
    For Each coordName In coordKeys
        ' Sanitize coordinator name using the helper function
        coordName = SanitizeSheetName(coordName)
        
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
            Set newTable = newTab.ListObjects(1)
            
            ' Perform the lookup to get the corresponding "NOMBRE" for the "COORDINADOR"
            Dim aliasRange As Range
            Dim nameRange As Range
            
            ' Set the range for ALIAS and NOMBRE columns (the table's actual range)
            Set aliasRange = wsColaboradores.ListObjects(COORDINADORES_TABLE).ListColumns(ALIAS_COLUMN).DataBodyRange
            Set nameRange = wsColaboradores.ListObjects(COORDINADORES_TABLE).ListColumns(NOMBRE_COLUMN).DataBodyRange
            
            ' Perform lookup using Application.Match instead of WorksheetFunction.Lookup
            On Error Resume Next
            Dim matchRow As Long
            matchRow = Application.Match(coordName, aliasRange, 0)
            If Not IsError(matchRow) Then
                coordAlias = nameRange.Cells(matchRow, 1).Value
            Else
                coordAlias = "Unknown Coordinator"
            End If
            On Error GoTo 0
            
            ' Paste the found coordinator name (or default) into cell B1 (merged B1:D1) in the new tab
            newTab.Range("B1:D1").Value = coordAlias
            
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
        
        ' Use the reusable function to copy shared values
        PasteCommonValues newTab, razonSocial, periodoDelPagoDel, fechaDeExpedicion, periodoDelPagoAl
        
        ' Auto-fit columns after pasting data
        newTab.Cells.EntireColumn.AutoFit
    Next coordName
    
    ' Loop through the filtered rows (only visible rows for each coordinator) and copy the filtered data
    Dim visibleCells    As Range, newRow As ListRow
    
    For Each coordName In coordKeys
        
        ' Apply filter for the current coordinator
        tableObj.Range.AutoFilter Field:=coordColumnIndex, Criteria1:=coordName
        
        ' Attempt to get visible cells
        On Error Resume Next
        Set visibleCells = tableObj.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        ' Ensure visibleCells only contains rows for the current coordinator
        Dim isValid     As Boolean
        isValid = TRUE
        
        If Not visibleCells Is Nothing Then
            For Each cell In visibleCells.Columns(coordColumnIndex).Cells
                If Trim(UCase(cell.Value)) <> coordName Then
                    isValid = FALSE
                    Exit For
                End If
            Next cell
        Else
            isValid = FALSE
        End If
        
        If Not isValid Then
            GoTo SkipCoordinator
        End If
        
        ' Check if the sheet exists
        On Error Resume Next
        Set newTab = ThisWorkbook.Sheets(coordName)
        On Error GoTo 0
        
        ' If the sheet doesn't exist, create it by copying the template
        If newTab Is Nothing Then
            ' Create a new tab
            Set newTab = CreateNewTab(templateSheet, coordName, ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count), newTabs)
        End If
        
        ' Ensure correct table reference in copied or existing sheet
        Set newTable = newTab.ListObjects(1)
        
        ' Set the table name based on the coordinator's name
        Dim newTableName As String
        newTableName = "Tabla_Coordinador_" & Replace(SanitizeSheetName(coordName), " ", "_")
        
        ' Change the table name
        On Error Resume Next
        newTable.Name = newTableName
        On Error GoTo 0
        
        ' Clear the table in the new tab
        If newTable.ListRows.Count > 0 Then
            newTable.DataBodyRange.Delete
        End If
        
        Call PopulateTable(newTable, visibleCells, tableObj, headerMapping, tableStartRow, wsSource)
        
        ' Auto-fit columns after inserting data
        newTab.Cells.EntireColumn.AutoFit
        
        SkipCoordinator:
    Next coordName
    
    RestoreApplicationSettings
    
    ' Reset the filter
    If wsSource.AutoFilterMode Then
        wsSource.AutoFilterMode = FALSE
    End If
    
    ' Restore the original filter state in the active sheet
    tableObj.Range.AutoFilter Field:=coordColumnIndex
    
    RestoreSheetVisibility sheetState, newTabs
    
    Exit Sub
    
    ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error in CreateCoordinatorTabs: " & Err.Description
        HandleError ERROR_GENERIC & " " & Err.Number & ": " & Err.Description, "CreateCoordinatorTabs"
    End If
    
    RestoreSheetVisibility sheetState, newTabs
    RestoreApplicationSettings
    Exit Sub
    
End Sub