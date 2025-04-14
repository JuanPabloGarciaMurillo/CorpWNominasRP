Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsColaboradores As Worksheet
    Dim activeTable As ListObject
    Dim aliases As Collection
    Dim gerenteAlias As String
    Dim validationRange As Range
    Dim managerCell As Range

    ' Set the named range for the manager's name
    Set managerCell = Me.Range("Nombre_Gerente")
    ' Check if the change occurred in the named range "Nombre_Gerente"
    If Not Intersect(Target, managerCell) Is Nothing Then
        On Error GoTo ErrorHandler
        ' Check if the named range is empty
        If managerCell.value = "" Then
            MsgBox "La celda 'Nombre_Gerente' está vacía. Por favor, ingrese un nombre de gerente válido.", vbExclamation, "Error"
            Exit Sub
        End If

        ' Ensure the active sheet is renamed correctly
        If RenameGerenteTabToAlias() = "" Then
            MsgBox "No se encontraron gerentes válidos. Saliendo de la macro.", vbExclamation, "Error"
            Exit Sub
        End If
        
        ' Get the manager alias for the name in the named range
        gerenteAlias = GetManagerAliasFromNombreGerente()
        If gerenteAlias = "" Then Exit Sub

        ' Set the Colaboradores sheet and the Coordinadores_Gerencia_Activa table
        Set wsColaboradores = ThisWorkbook.Sheets("Colaboradores")
        Set activeTable = wsColaboradores.ListObjects("Coordinadores_Gerencia_Activa")
        
        ' Clear the existing rows in the table
        Call ClearTableRows(activeTable)
        
        ' Get the coordinator aliases for the manager alias
        Set aliases = GetCoordinatorAliases(gerenteAlias)
        If aliases Is Nothing Then Exit Sub
        
        ' Check if there are no aliases
        If aliases.Count = 0 Then
            MsgBox "No se encontraron coordinadores para el gerente '" & managerCell.value & "'.", vbInformation, "Sin Resultados"
            Exit Sub
        End If
        
        ' Populate the table with the aliases
        Call PopulateTableWithCollection(activeTable, aliases)
        
        ' Sort the table alphabetically by the first column
        Call SortTableAlphabetically(activeTable, 1)
    End If

    ' Check if the change occurred in the COORDINADOR column (column A)
    If Not Intersect(Target, Me.Columns("A")) Is Nothing Then
        On Error GoTo ErrorHandler
        
        Dim cell As Range
        For Each cell In Intersect(Target, Me.Columns("A"))
            ' Set the validation range to the corresponding cell in column J
            Set validationRange = Me.Cells(cell.row, "J")
            ' Call ProcessRow to handle validation
            ProcessRow cell, validationRange
        Next cell
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error al actualizar la tabla Coordinadores_Gerencia_Activa: " & Err.Description & _
           vbNewLine & "Error en la fila: " & Target.row, vbCritical, "Error"
End Sub

Private Sub Worksheet_Calculate()
    On Error GoTo ErrorHandler
    Dim wsCurrent As Worksheet
    Dim tblGerente As ListObject
    Dim coordRange As Range
    Dim coordCell As Range
    Dim promotorCell As Range

    ' Set the current sheet
    Set wsCurrent = Me

    ' Dynamically retrieve the only table on the sheet
    If wsCurrent.ListObjects.Count = 1 Then
        Set tblGerente = wsCurrent.ListObjects(1)
    Else
        MsgBox "Error: No table or multiple tables found on the sheet.", vbCritical, "Error"
        Exit Sub
    End If

    ' Get the range of COORDINADOR (Column A) within the table
    Set coordRange = tblGerente.ListColumns("COORDINADOR").DataBodyRange

    ' Loop through each cell in COORDINADOR (Column A)
    For Each coordCell In coordRange
        ' Set the corresponding PROMOTOR cell (Column J)
        Set promotorCell = wsCurrent.Cells(coordCell.row, "J")
        ' Process the row
        ProcessRow coordCell, promotorCell
    Next coordCell

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in Worksheet_Calculate: " & Err.Description, vbCritical, "Error"
End Sub


