' Subroutine: CreatePromotorTabs
' Version: 0.9.3
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This subroutine automates the process of creating individual tabs for each promotor in the workbook. It collects unique promotor names from the "Promotores" table in the "Colaboradores" sheet and checks if each promotor has a corresponding entry in the "Sueldos_Base" table in the "Tabuladores" sheet. If the promotor exists in the "Sueldos_Base" table, the subroutine creates a new sheet for that promotor by copying a template sheet and renaming it according to the promotor's name. The subroutine then populates the new tabs with relevant data from the "Promotores" table and applies filters to include only the relevant data for each promotor. Common values from the source sheet (e.g., "razonSocial", "periodoDelPagoDel") are also copied into all the new tabs.
' Notes:
'   - The subroutine creates a new tab for each promotor by copying a template and renaming it to the promotor's name, ensuring the name is valid.
'   - It first checks for each promotor's base salary entry in the "Sueldos_Base" table and only creates tabs for those with a valid match.
'   - The promotor names are sanitized to ensure they are valid sheet names (e.g., replacing invalid characters with underscores).
'   - The subroutine automatically sorts the "Promotor" column in ascending order before processing the rows.
'   - If no valid promotors are found, an error message is displayed.
'   - Filters are applied to each promotor's data, and new tabs are populated with filtered rows.
'   - After creating the tabs, common values from the source sheet are pasted into corresponding cells in each new tab.
'   - The original visibility state of the sheets is preserved, with only the newly created tabs visible.

Public Sub CreatePromotorTabs()
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
    Dim sheetState  As Collection
    Dim wsColaboradores As Worksheet
    Dim columnIndex As Variant
    Dim razonSocial     As Variant
    Dim periodoDelPagoDel As Variant
    Dim fechaDeExpedicion As Variant
    Dim periodoDelPagoAl As Variant
    Dim matchRow    As Long
    Dim newTable As ListObject
    Dim promotorName As Variant
    Dim promotorDict As clsDictionary
    Dim promotorColumn As Range
    Dim newTabs     As Collection
    Dim wsTabuladores As Worksheet
    Dim promotorTable As ListObject
    Dim baseSalaryTable As ListObject
    Dim promotorCoord As String
    Dim promotorRow As ListRow
    Dim tabuladorRow As ListRow
    Dim promotorMatch As Boolean
    Dim coordinatorName As String
    Dim promotorColumnIndex As Long

    ' Initialize dictionaries using the custom dictionary class
    Set headerMapping = New clsDictionary
    Set promotorDict = New clsDictionary
    
    ' Set the source sheet as the active sheet where the button is clicked
    Set wsSource = ActiveSheet
    Set templateSheet = ThisWorkbook.Sheets(PROMOTORES_SHEET)
    Set wsColaboradores = ThisWorkbook.Sheets(COLABORADORES_SHEET)
    Set wsTabuladores = ThisWorkbook.Sheets(TABULADORES_SHEET)
    ' Set the tables
    Set promotorTable = wsColaboradores.ListObjects(PROMOTORES_TABLE)
    Set baseSalaryTable = wsTabuladores.ListObjects(SUELDOS_BASE_TABLE)

    ' Create a dictionary to store the visibility status of sheets
    Set sheetState = New Collection
    Set newTabs = New Collection
    
    ' Set the table range using ListObjects (Excel table object)
    Set tableObj = wsSource.ListObjects(1)
    
    ' Initialize the header mapping using the reusable function
    InitializeHeaderMapping HEADERS, COLUMN_INDICES, headerMapping
    
    ' Unhide all sheets and store their original state (hidden or visible)
    For Each ws In ThisWorkbook.Sheets
        ' Store visibility state using ws.Name as the key
        sheetState.Add ws.Visible, ws.Name
        ws.Visible = xlSheetVisible
    Next ws
    
    ' Define the start row for the table
    tableStartRow = 9
    
    ' Get the last row of the table data (excluding Totals Row)
    lastDataRow = tableObj.ListRows.Count + tableStartRow - 1
    
    ' Define the range for the "Promotor" column (from row 9 to the last data row)
    Set promotorColumn = wsSource.Range("A" & tableStartRow & ":A" & lastDataRow)
    
    ' Loop through the filtered Promotores and find Tabulador data
    coordinatorName = wsSource.Name
    For Each promotorRow In promotorTable.ListRows
        
        baseSalaryPromotorName = promotorRow.Range.Cells(1, promotorTable.ListColumns(NOMBRE_COLUMN).Index).Value
        baseSalaryPromotorAlias = promotorRow.Range.Cells(1, promotorTable.ListColumns(ALIAS_COLUMN).Index).Value
        promotorCoord = promotorRow.Range.Cells(1, promotorTable.ListColumns(COORDINACION_COLUMN).Index).Value
        
        If Trim(UCase(promotorCoord)) = Trim(UCase(coordinatorName)) Then
            
            ' Search for the Promotor in the Tabuladores table (Sueldos_Base)
            promotorMatch = FALSE
            For Each tabuladorRow In baseSalaryTable.ListRows
                ' Check if Promotor matches Tabulador COLABORADOR
                If tabuladorRow.Range.Cells(1, baseSalaryTable.ListColumns("COLABORADOR").Index).Value = baseSalaryPromotorName Then
                    promotorMatch = TRUE
                    If Not promotorDict.Exists(baseSalaryPromotorAlias) Then
                        promotorDict.Add baseSalaryPromotorAlias, Nothing
                    End If
                    Exit For
                End If
            Next tabuladorRow
            
        End If
    Next promotorRow
    ' Loop through the "Promotor" column to collect unique values from the table only
    For Each cell In promotorColumn
        promotorName = Trim(cell.Value)
        
        ' Check if the cell has a valid promotor name and it's not already in the dictionary
        If promotorName <> "" And promotorName <> PROMOTOR_COLUMN And Not promotorDict.Exists(promotorName) Then
            promotorDict.Add promotorName, Nothing
        End If
    Next cell
    
    ' Prevent errors if no promotors are found
    If promotorDict.Count = 0 Then
        ' Restore application settings
        RestoreApplicationSettings
        
        ' Restore sheet visibility and exit
        RestoreSheetVisibility sheetState, newTabs
        Exit Sub
    End If
    
    promotorColumnIndex = tableObj.ListColumns(PROMOTOR_COLUMN).Index
    
    ' Sort the "Promotor" column in ascending order (A-Z)
    tableObj.Sort.SortFields.Clear
    tableObj.Sort.SortFields.Add Key:=tableObj.ListColumns(promotorColumnIndex).Range, _
                                 SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    tableObj.Sort.Apply
    ' Iterate over unique promotors and create or recreate tabs
    For Each promotorName In promotorDict.GetKeys
        ' Sanitize promotor name using the helper function
        promotorName = SanitizeSheetName(promotorName)
        
        ' Check if sheet already exists
        If SheetExists(promotorName) Then
            MsgBox "El promotor        '" & promotorName & "' tiene un registro asignado a otra coordinacion. " & _
                   "Porfavor revisa la data y vuelve a intentar.", vbCritical, "Promotor duplicado"
            GoTo ErrHandler
        End If
        
        ' Create a new tab
        Set newTab = CreateNewTab(templateSheet, promotorName, wsSource, newTabs)
        
        ' Ensure correct table reference in copied sheet
        Set newTable = newTab.ListObjects(1)

        ' Set the table name using the reusable function
        SetTableName newTable, promotorName, "Tabla_Promotor_"

        SetAliasInNewTab newTab, PROMOTORES_TABLE, ALIAS_COLUMN, NOMBRE_COLUMN, promotorName, "Unknown Promotor"
    Next promotorName
    
    ' Collect valid tabs based on promotors with a base salary
    Dim validTabs   As Collection
    Set validTabs = New Collection
    
    For Each promotorName In promotorDict.GetKeys
        validTabs.Add promotorName
    Next promotorName
    
    ' Get the values to copy from the active sheet (B2, B3, B6, D3)
    razonSocial = wsSource.Range("B2").Value
    periodoDelPagoDel = wsSource.Range("B3").Value
    fechaDeExpedicion = wsSource.Range("B6").Value
    periodoDelPagoAl = wsSource.Range("D3").Value
    
    ' Now loop through each new tab and paste the values into the corresponding cells (B2, B3, B6, D3)
    For Each promotorName In newTabs
        Set newTab = ThisWorkbook.Sheets(promotorName)
        
        ' Use the reusable function to copy shared values
        PasteCommonValues newTab, razonSocial, periodoDelPagoDel, fechaDeExpedicion, periodoDelPagoAl
        
        ' Auto-fit columns after pasting data
        newTab.Cells.EntireColumn.AutoFit
    Next promotorName
    
    ' Loop through the filtered rows (only visible rows for each promotor) and copy the filtered data
    Dim visibleCells As Range, newRow As ListRow
    
    For Each promotorName In promotorDict.GetKeys
        
        ' Always clear visibleCells before applying the filter
        Set visibleCells = Nothing
        
        ' Apply filter for the current Promotor
        tableObj.Range.AutoFilter Field:=promotorColumnIndex, Criteria1:=promotorName
        
        ' Get the corresponding worksheet for the promotor
        On Error Resume Next
        Set visibleCells = tableObj.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        ' Clear the previous data in the new tab before populating (no matter if data is found or not)
        If Not visibleCells Is Nothing Then
            Set newTab = ThisWorkbook.Sheets(promotorName)
            Set newTable = newTab.ListObjects(1)
            
            ' Clear the table in the new tab
            If newTable.ListRows.Count > 0 Then
                newTable.DataBodyRange.Delete
            End If
            
            ' Only loop through the filtered rows if there are visible cells
            If Not visibleCells Is Nothing Then
                ' Loop through filtered rows
                PopulateTable newTable, visibleCells, tableObj, headerMapping, tableStartRow, wsSource
                
                ' Auto-fit columns after inserting data
                newTab.Cells.EntireColumn.AutoFit
            End If
            
            ' Reset the filter after processing the current promotor
            If tableObj.AutoFilter.FilterMode Then
                tableObj.AutoFilter.ShowAllData
            End If
        End If
        
    Next promotorName
    
    RestoreApplicationSettings
    
    ' Reset the filter
    If wsSource.AutoFilterMode Then
        wsSource.AutoFilterMode = FALSE
    End If
    
    ' Restore the original filter state in the active sheet
    tableObj.Range.AutoFilter Field:=promotorColumnIndex
    
    RestoreSheetVisibility sheetState, newTabs
    
    ' Execute SumPagoNetoCoordinacion at the end
    If Not newTabs Is Nothing And newTabs.Count > 0 Then
        SumPagoNetoCoordinacion newTabs, wsSource
    End If
    
    ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error in CreatePromotorTabs: " & Err.Description
        HandleError ERROR_GENERIC & " " & Err.Number & ": " & Err.Description, "CreatePromotorTabs"
    End If
    
    RestoreSheetVisibility sheetState, newTabs
    RestoreApplicationSettings
    Exit Sub
    
End Sub