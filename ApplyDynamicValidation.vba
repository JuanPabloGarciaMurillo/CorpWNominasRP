'Public Sub ApplyDynamicValidation(ByVal ws As Worksheet, ByVal tableName As String, ByVal valueColumnName As String, ByVal validatedColumnName As String)
'    Dim tbl As ListObject
'    Dim row As ListRow
'    Dim coordValue As String
'    Dim aliasList As String
'    Dim validationRange As Range
'    Dim ValueColumn As Long
'
'    ' Set references
'    Set tbl = ws.ListObjects(tableName)
'    ValueColumn = tbl.ListColumns(valueColumnName).Index
'
'    ' Loop through each row in the table
'    For Each row In tbl.ListRows
'        coordValue = row.Range.Cells(1, ValueColumn).value ' Get COORDINADOR value
'        Set validationRange = row.Range.Cells(1, tbl.ListColumns(validatedColumnName).Index)
'
'        If coordValue = "" Then
'            ' If value is empty, clear validation
'            ClearValidation validationRange
'        Else
'            ' Get alias list
'            aliasList = GetAliasList(coordValue)
'
'            ' Apply validation if alias list is not empty
'            If aliasList <> "" Then
'                ApplyValidation validationRange, aliasList
'            Else
'                ClearValidation validationRange
'            End If
'        End If
'    Next row
'End Sub
'
'' Function to retrieve alias list using TEXTJOIN formula
'Private Function GetAliasList(ByVal coordValue As String) As String
'    Dim validationFormula As String
'    validationFormula = "=TEXTJOIN("","", TRUE, IF(Promotores[COORDINACION] = """ & coordValue & """, Promotores[ALIAS], """" ))"
'    GetAliasList = Trim(Evaluate(validationFormula)) ' Evaluate and return trimmed result
'End Function
'
'' Function to clear validation from a cell
'Private Sub ClearValidation(ByVal validationRange As Range)
'    On Error Resume Next
'    validationRange.Validation.Delete
'    On Error GoTo 0
'End Sub
'
'' Function to apply validation from a list
'Private Sub ApplyValidation(ByVal validationRange As Range, ByVal aliasList As String)
'    Dim aliasArray() As String
'    Dim i As Long
'
'    ' Split alias list into array and trim spaces
'    aliasArray = Split(aliasList, ",")
'    For i = LBound(aliasArray) To UBound(aliasArray)
'        aliasArray(i) = Trim(aliasArray(i))
'    Next i
'
'    ' Apply validation
'    validationRange.Validation.Delete
'    validationRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
'        Operator:=xlBetween, Formula1:=Join(aliasArray, ",")
'End Sub