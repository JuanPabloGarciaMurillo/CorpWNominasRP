'=========================================================
' Script: Worksheet(Gerente)
' Version: 0.9.0
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This script handles events for the "Gerente" worksheet in Excel. It includes event handlers for changes in the worksheet and recalculations. The script manages data validation, populates tables with aliases, and refreshes PivotTables on the dashboard.
'
' Parameters:
'   - None (applies to the entire worksheet upon changes or recalculations).
' Returns:
'   - None
' Notes:
'   - The script assumes the presence of specific named ranges and structured tables in the worksheet.
'   - It uses constants for sheet names, table names, and column names to improve readability and maintainability.
'=========================================================


'=========================================================
' Event: Worksheet_Change
' Description:
'   Triggered when a cell on the worksheet changes. This event handler:
'     - Detects if the "Nombre_Gerente" cell was edited.
'     - Retrieves the alias for the specified manager.
'     - Filters and loads related coordinator aliases into the "Coordinadores_Gerencia_Activa" table on the "Colaboradores" sheet.
'     - Sorts the table by the "COORDINADOR" column.
'     - Refreshes all PivotTables on the "Resultados" (Dashboard) sheet.
'     - Applies dynamic data validation to "PROMOTOR" cells based on changes in the "COORDINADOR" column.
' Parameters:
'   - Target (Range): The range that was changed by the user.
' Returns:
'   - None
' Notes:
'   - Assumes the sheet contains a named range "Nombre_Gerente".
'   - Expects a structured table with "COORDINADOR" and "PROMOTOR" columns.
'   - Assumes a single table is present in the worksheet.
'=========================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsColaboradores As Worksheet
    Dim activeTable As ListObject
    Dim aliases As Collection
    Dim gerenteAlias As String
    Dim validationRange As Range
    Dim managerCell As Range
    Dim wsDashboard As Worksheet
    Dim pivotTable As PivotTable
    Dim coordColumnIndex As Long
    Dim promotorColumnIndex As Long

    ' Set the named range for the manager's name
    Set managerCell = Me.Range("Nombre_Gerente")
    ' Check if the change occurred in the named range "Nombre_Gerente"
    If Not Intersect(Target, managerCell) Is Nothing Then
        On Error GoTo ErrorHandler
        ' Check if the named range is empty
        If managerCell.value = "" Then
            MsgBox ERROR_EMPTY_MANAGER_CELL, vbExclamation, "Error"
            Exit Sub
        End If

        ' Ensure the active sheet is renamed correctly
        If RenameGerenteTabToAlias() = "" Then
            MsgBox ERROR_NO_VALID_MANAGER, vbExclamation, "Error"
            Exit Sub
        End If
        
        ' Get the manager alias for the name in the named range
        gerenteAlias = GetManagerAliasFromNombreGerente()
        If gerenteAlias = "" Then Exit Sub

        ' Set the Colaboradores sheet and the Coordinadores_Gerencia_Activa table
        Set wsColaboradores = ThisWorkbook.Sheets(COLABORADORES_SHEET)
        Set activeTable = wsColaboradores.ListObjects(ACTIVE_TABLE)
        
        ' Clear the existing rows in the table
        Call ClearTableRows(activeTable)
        
        ' Get the coordinator aliases for the manager alias
        Set aliases = GetCoordinatorAliases(gerenteAlias)
        If aliases Is Nothing Then Exit Sub
        
        ' Check if there are no aliases
        If aliases.Count = 0 Then
            MsgBox ERROR_NO_COORDINATORS & managerCell.value & "'.", vbInformation, "Sin Resultados"
            Exit Sub
        End If
        
        ' Populate the table with the aliases
        Call PopulateTableWithCollection(activeTable, aliases)
        
        ' Sort the table alphabetically by the "COORDINADOR" column
        coordColumnIndex = activeTable.ListColumns(COORDINADOR_COLUMN).Index
        Call SortTableAlphabetically(activeTable, coordColumnIndex)
        
    End If

    ' Refresh all pivot tables on the Dashboard sheet
    On Error Resume Next

    Set wsDashboard = ThisWorkbook.Sheets(RESULTADOS_SHEET)
    If Not wsDashboard Is Nothing Then
        For Each pivotTable In wsDashboard.PivotTables
            pivotTable.RefreshTable
        Next pivotTable
    End If
    On Error GoTo 0

    ' Check if the change occurred in the "COORDINADOR" column
    If Not Intersect(Target, Me.ListObjects(1).ListColumns(COORDINADOR_COLUMN).DataBodyRange) Is Nothing Then
        On Error GoTo ErrorHandler
        
        Dim cell As Range
        
        ' Get dynamic column indices
        coordColumnIndex = Me.ListObjects(1).ListColumns(COORDINADOR_COLUMN).Index
        promotorColumnIndex = Me.ListObjects(1).ListColumns(PROMOTOR_COLUMN).Index

        For Each cell In Intersect(Target, Me.ListObjects(1).ListColumns(COORDINADOR_COLUMN).DataBodyRange)
            ' Set the validation range to the corresponding cell in the "PROMOTOR" column
            Set validationRange = Me.ListObjects(1).ListColumns(PROMOTOR_COLUMN).DataBodyRange.Cells(cell.Row - Me.ListObjects(1).DataBodyRange.Row + 1)
            ' Call ProcessRow to handle validation
            ProcessRow cell, validationRange
        Next cell
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox ERROR_UPDATE_TABLE & Err.Description & vbNewLine & "Error en la fila: " & Target.row, vbCritical, "Error"
    If Err.Number <> 0 Then
        MsgBox ERROR_UPDATE_DASHBOARD & Err.Description, vbCritical, "Error"
    End If
End Sub

'=========================================================
' Event: Worksheet_Calculate
' Description:
'   Triggered when the worksheet is recalculated. This handler:
'     - Loops through the "COORDINADOR" column of the first table in the sheet.
'     - Finds the corresponding cell in the "PROMOTOR" column.
'     - Calls the `ProcessRow` subroutine to apply data validation or other processing logic row by row.
' Parameters:
'   - None (applies to the entire worksheet upon recalculation).
' Returns:
'   - None
' Notes:
'   - Requires exactly one structured table on the worksheet.
'   - Table must contain both "COORDINADOR" and "PROMOTOR" columns.
'   - Uses dynamic positioning to align data across the two columns.
'=========================================================
Private Sub Worksheet_Calculate()
    On Error GoTo ErrorHandler
    Dim wsCurrent As Worksheet
    Dim tblGerente As ListObject
    Dim coordRange As Range
    Dim coordCell As Range
    Dim promotorCell As Range
    Dim promotorColumn As ListColumn

    ' Set the current sheet
    Set wsCurrent = Me

    ' Dynamically retrieve the only table on the sheet
    If wsCurrent.ListObjects.Count = 1 Then
        Set tblGerente = wsCurrent.ListObjects(1)
    Else
        MsgBox "Error: No table or multiple tables found on the sheet.", vbCritical, "Error"
        Exit Sub
    End If

    ' Get the range of "COORDINADOR" within the table
    Set coordRange = tblGerente.ListColumns(COORDINADOR_COLUMN).DataBodyRange
    Set promotorColumn = tblGerente.ListColumns(PROMOTOR_COLUMN)

    ' Loop through each cell in "COORDINADOR"
    For Each coordCell In coordRange
        ' Set the corresponding "PROMOTOR" cell dynamically
        Set promotorCell = promotorColumn.DataBodyRange.Cells(coordCell.Row - tblGerente.DataBodyRange.Row + 1)
        ' Process the row
        ProcessRow coordCell, promotorCell
    Next coordCell

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in Worksheet_Calculate: " & Err.Description, vbCritical, "Error"
End Sub