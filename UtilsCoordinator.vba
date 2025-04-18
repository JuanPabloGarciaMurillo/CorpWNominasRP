'=========================================================
' Script: UtilsCoordinator
' Version: 0.9.0
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This module contains utility functions for working with coordinators in Excel VBA.
'   It includes functions for retrieving aliases based on manager names and deleting tabs.
'   It also includes a function to create new tabs based on the coordinator's aliases.
' Functions included in this module:
'   - GetCoordinatorAliases
'   - DeleteManagerCoordinatorTab
'   - CreateCoordinatorTabs_newTabs
'=========================================================

'=========================================================
' Function: GetCoordinatorAliases
' Description:
'   Checks for the aliases of coordinators based on the manager name.
' Parameters:
'   - managerName (String): The name of the manager to check for.
' Returns:
'  - Collection: A collection of aliases for the specified manager.
' Notes:
'   - The function uses a ListObject (table) named "Coordinadores" in the "Colaboradores" sheet.
'   - It retrieves the GERENCIA and ALIAS columns from the table.
'=========================================================
Public Function GetCoordinatorAliases(Optional ByVal managerName As String = "") As Collection
    Dim aliases As Collection
    Dim wsColaboradores As Worksheet
    Dim coordinadoresTable As ListObject
    Dim coordinatorRow As ListRow
    Dim gerenciaColumn As ListColumn
    Dim aliasColumn As ListColumn
    
    ' Initialize the collection to store aliases
    Set aliases = New Collection

    On Error GoTo ErrorHandler

    ' If no manager name is provided, use GetManagerAliasFromNombreGerente
    If managerName = "" Then
        Dim managerAlias As String
        managerAlias = GetManagerAliasFromNombreGerente()
        If managerAlias = "" Then Exit Function
        managerName = managerAlias ' Use the alias as the manager name
    End If

    Set wsColaboradores = ThisWorkbook.Sheets(COLABORADORES_SHEET)
    Set coordinadoresTable = wsColaboradores.ListObjects(COORDINADORES_TABLE)
    Set gerenciaColumn = coordinadoresTable.ListColumns(GERENCIA_COLUMN)
    Set aliasColumn = coordinadoresTable.ListColumns(ALIAS_COLUMN)
    
    ' Loop through the rows in the Coordinadores table
    For Each coordinatorRow In coordinadoresTable.ListRows
        ' Check if the GERENCIA matches the manager name
        If Trim(UCase(coordinatorRow.Range.Cells(1, gerenciaColumn.Index).Value)) = Trim(UCase(managerName)) Then
            ' Add the ALIAS to the collection
            aliases.Add coordinatorRow.Range.Cells(1, aliasColumn.Index).Value
        End If
    Next coordinatorRow
    
    ' Return the collection of aliases
    Set GetCoordinatorAliases = aliases
    Exit Function

ErrorHandler:
    MsgBox "Error en la función GetCoordinatorAliases, por favor contacta a tu administrador: " & Err.Description, vbCritical, "Error"
    Set GetCoordinatorAliases = Nothing
End Function

'=========================================================
' Function: DeleteManagerCoordinatorTab
' Description:
'   Deletes all tabs in the workbook whose names end with " (C)".
' Parameters:
'   - None
' Returns:
'   - None
' Notes:
'   - This function iterates through all sheets in the workbook and deletes those whose names end with " (C)".
'=========================================================
Public Sub DeleteManagerCoordinatorTab()
    Dim ws          As Worksheet
    Dim tabName     As String
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        tabName = ws.Name
        
        ' Check if the tab name ends with "(C)" (ignoring trailing spaces)
        If Right(Trim(tabName), Len(TAB_SUFFIX)) = TAB_SUFFIX Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub

'=========================================================
' Function: CreateCoordinatorTabs_newTabs
' Description:
'   This function returns the collection of newly created tabs (newTabs).
'   It provides access to the collection for checking or further processing
'   of the newly created sheets.
' Parameters:
'   - None
' Returns:
'   - Collection: The collection of sheet names representing the newly
'     created tabs.
' Notes:
'   - The function assumes that the collection `newTabs` has been properly
'     populated elsewhere in the code.
'=========================================================

' Function to return the newTabs collection from the global scope
Public Function CreateCoordinatorTabs_newTabs() As Collection
    ' This function returns the newTabs collection
    If newTabs Is Nothing Then
        MsgBox "Error: La colección 'newTabs' no está inicializada.", vbCritical, "Error"
        Exit Function
    End If
    Set CreateCoordinatorTabs_newTabs = newTabs
End Function