Public Sub CreateBaseSalaryTabsIfMissing( _
    coordinatorName As String, _
    templateSheet As Worksheet, _
    wsSource As Worksheet, _
    razonSocial As Variant, _
    periodoDelPagoDel As Variant, _
    periodoDelPagoAl As Variant, _
    fechaDeExpedicion As Variant, _
    newTabs As Collection)

    On Error GoTo ErrHandler
    Debug.Print "Entering CreateBaseSalaryTabsIfMissing..."
    
    Dim wsColaboradores As Worksheet
    Dim wsTabuladores As Worksheet
    Dim promotorTable As ListObject
    Dim baseSalaryTable As ListObject
    Dim promotorName As String
    Dim promotorCoord As String
    Dim promotorRow As ListRow
    Dim newTab As Worksheet
    Dim invalidChars As Variant
    Dim i As Integer
    Dim matchRow As Long
    Dim colNombreIndex As Long, colCoordIndex As Long
    Dim promotorTableFiltered As ListObject
    Dim tabuladorRow As ListRow
    Dim promotorMatch As Boolean
    
    ' Set the worksheets
    Set wsColaboradores = ThisWorkbook.Sheets("Colaboradores")
    Set wsTabuladores = ThisWorkbook.Sheets("Tabuladores")
    
    ' Set the tables
    Set promotorTable = wsColaboradores.ListObjects("Promotores")
    Set baseSalaryTable = wsTabuladores.ListObjects("Sueldos_Base")
    
    ' Loop through the filtered Promotores and find Tabulador data
    For Each promotorRow In promotorTable.ListRows
        promotorName = promotorRow.Range.Cells(1, promotorTable.ListColumns("NOMBRE").Index).Value
        promotorAlias = promotorRow.Range.Cells(1, promotorTable.ListColumns("ALIAS").Index).Value
        promotorCoord = promotorRow.Range.Cells(1, promotorTable.ListColumns("COORDINACION").Index).Value
        
        If Trim(UCase(promotorCoord)) = Trim(UCase(coordinatorName)) Then
            Debug.Print "Processing Promotor: " & promotorName & " | " & promotorAlias
            
            ' Search for the Promotor in the Tabuladores table (Sueldos_Base)
            promotorMatch = False
            For Each tabuladorRow In baseSalaryTable.ListRows
                ' Check if Promotor matches Tabulador COLABORADOR
                If tabuladorRow.Range.Cells(1, baseSalaryTable.ListColumns("COLABORADOR").Index).Value = promotorName Then
                    Debug.Print "Found matching Tabulador for Promotor: " & promotorAlias
                    
                    ' Extract required data from the Tabuladores table (e.g., base salary or other columns)
                    Dim baseSalary As Variant
                    baseSalary = tabuladorRow.Range.Cells(1, baseSalaryTable.ListColumns("SUELDO BASE").Index).Value
                    Debug.Print "Found Promotor: " & promotorName & " with Base Salary: " & baseSalary
                    Debug.Print ""
                    ' Example: Create a new sheet for the Promotor
                     Set newTab = ThisWorkbook.Sheets.Add
                     newTab.Name = promotorAlias
                    
                    promotorMatch = True
                    Exit For
                End If
            Next tabuladorRow
            
            If Not promotorMatch Then
                Debug.Print "No matching Tabulador found for Promotor: " & promotorName
                Debug.Print ""
            End If
        End If
    Next promotorRow

    Debug.Print "Exiting CreateBaseSalaryTabsIfMissing..."
    Exit Sub

ErrHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description & " in CreateBaseSalaryTabsIfMissing"
End Sub


