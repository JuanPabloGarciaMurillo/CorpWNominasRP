'Private Sub Worksheet_Change(ByVal Target As Range)
'    Dim tbl As ListObject
'    Set tbl = Me.ListObjects("Tabla_Gerente")
'
'    ' Check if the change happened in the "COORDINADOR" column
'    If Not Intersect(Target, tbl.ListColumns("COORDINADOR").DataBodyRange) Is Nothing Then
'        ApplyDynamicValidation Me, "Tabla_Gerente", "COORDINADOR", "PROMOTOR"
'    End If
'End Sub