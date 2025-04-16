'==================================================
' Script: UtilsManager
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This module contains utility functions for working with managers in Excel VBA.
'   It includes functions for renaming sheets based on manager names and retrieving aliases.
'   It also includes a function to get the manager's alias from a specified cell.
' Functions included in this module:
'   - GetManagerAliasFromNombreGerente
'   - RenameGerenteTabToAlias
'==================================================

'==================================================
' Function: GetManagerAliasFromNombreGerente
' Description:
'    This function retrieves the manager's name from cell B1 and checks for their aliases.
'    It uses the GetCoordinatorAliases function to get the aliases based on the manager's name.
' Parameters:
'   -  None
' Returns:
'   -   Collection: A collection of aliases for the specified manager.
' Notes:
'   -  The function checks if cell B1 is empty and shows a message if so.
'   -  It retrieves the manager's name from cell B1 and uses it to call the GetCoordinatorAliases function.
'======================================================================
Public Function GetManagerAliasFromNombreGerente() As String
    Dim wsColaboradores As Worksheet
    Dim gerentesTbl As ListObject
    Dim gerenteNombre As String
    Dim gerenteAlias As String
    Dim foundRow As Range

    On Error GoTo ErrorHandler

    ' Get the manager's name from the named range "Nombre_Gerente"
    gerenteNombre = Trim(ThisWorkbook.ActiveSheet.Range("Nombre_Gerente").Value)

    ' Check if the named range is empty
    If gerenteNombre = "" Then
        MsgBox "La celda 'Nombre_Gerente' está vacía. Por favor, ingrese un nombre válido.", vbExclamation, "Error"
        Exit Function
    End If

    ' Set the Colaboradores sheet and Gerentes table
    Set wsColaboradores = ThisWorkbook.Sheets("Colaboradores")
    Set gerentesTbl = wsColaboradores.ListObjects("Gerentes")

    ' Find the alias for the manager in the Gerentes table
    On Error Resume Next
    Set foundRow = gerentesTbl.ListColumns("NOMBRE").DataBodyRange.Find(What:=gerenteNombre, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0

    If Not foundRow Is Nothing Then
        gerenteAlias = foundRow.Offset(0, 1).Value ' Assuming the alias is in the next column
        GetManagerAliasFromNombreGerente = gerenteAlias
    Else
        MsgBox "El Gerente '" & gerenteNombre & "' no se encuentra en la tabla de Gerentes.", vbExclamation, "Error"
        Exit Function
    End If

    Exit Function

ErrorHandler:
    MsgBox "Error en la función GetManagerAliasFromNombreGerente: " & Err.Description, vbCritical, "Error"
    GetManagerAliasFromNombreGerente = ""
End Function

'====================================================
' Function: RenameGerenteTabToAlias
' Description:
'   Renames the active Gerente sheet based on the value in cell B1
'   (Nombre_Gerente), using the ALIAS from the "Gerentes" table on
'   the "Colaboradores" sheet.
' Returns:
'   - True if renamed successfully, False otherwise.
' Notes:
'   - Checks if B1 is empty and shows a message if so.
'   - Searches for the Gerente name in the "Gerentes" table.
'   - Prevents renaming if the alias already exists as a sheet.
'====================================================

Public Function RenameGerenteTabToAlias() As String
    On Error GoTo ErrHandler
    
    Dim wsActive As Worksheet
    Set wsActive = ActiveSheet
    
    Dim aliasName As String
    Dim wsColab As Worksheet
    Dim tbl As ListObject

    ' Get the manager alias using the existing function
    aliasName = GetManagerAliasFromNombreGerente()
    
    ' Check if the alias is empty
    If aliasName = "" Then
        ' The error message is already handled in GetManagerAliasFromNombreGerente
        Exit Function
    End If
    
    ' Sanitize alias name using the helper function
    aliasName = SanitizeSheetName(aliasName)
    
    ' Validate the alias name
    If aliasName = "" Then
        MsgBox "Alias invalido, por favor revisa los valores de la tabla de Gerentes", vbExclamation, "RenameGerenteTabToAlias"
        Exit Function
    End If
    
    ' Check if the current sheet name matches the alias
    If wsActive.Name = aliasName Then
        RenameGerenteTabToAlias = aliasName ' Return the alias without renaming
        Exit Function
    End If
    
    ' Check if a sheet with the same name already exists
    If SheetExists(aliasName) Then
        MsgBox "No se puede renombrar: La hoja llamada '" & aliasName & "' ya existe.", vbExclamation, "RenameGerenteTabToAlias"
        RenameGerenteTabToAlias = ""
        Exit Function
    End If
    
    ' Rename the sheet
    wsActive.Name = aliasName
    RenameGerenteTabToAlias = aliasName ' Return the alias
    Exit Function

ErrHandler:
    MsgBox "Error al renombrar la hoja del gerente: " & Err.Description, vbCritical, "RenameGerenteTabToAlias"
    RenameGerenteTabToAlias = ""
End Function