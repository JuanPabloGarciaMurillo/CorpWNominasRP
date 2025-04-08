'=======================================================================
' Subroutine: CreateCoordinatorAndPromotorTabs
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This subroutine automates the process of creating new coordinator and
'   promotor tabs within the workbook.
'   It first calls the `CreateCoordinatorTabs` subroutine to create the
'   coordinator-specific worksheets and then iterates over the newly created
'   tabs. For each new coordinator tab, it activates the sheet and runs the
'   `CreatePromotorTabs` subroutine to create promotor-specific worksheets.
'   Finally, it calls `SumPagoNetoGerencia` to calculate the total "PAGO NETO"
'   values based on the activating sheet.
' Parameters:
'   - None
' Returns:
'   - None
' Notes:
'   - The `newTabs` collection, containing the names of the newly created
'     coordinator tabs, is retrieved from the `CreateCoordinatorTabs_newTabs`
'     function in the `Utils` module.
'   - The code uses error handling to capture and display any runtime errors
'     encountered during execution.
'=======================================================================

Sub CreateCoordinatorAndPromotorTabs()
    On Error GoTo ErrHandler
    Debug.Print "Starting CreateCoordinatorAndPromotorTabs..."
    
    ' Capture the sheet that activated the macro
    Dim activatingSheet As Worksheet
    Set activatingSheet = ActiveSheet
    
    ' Run RenameGerenteTabToAlias to rename the Manager(Gerente) tab
    Call RenameGerenteTabToAlias

    ' Run CreateCoordinatorTabs to create new coordinator tabs
    Call CreateCoordinatorTabs
    
    ' Now retrieve the newTabs collection from the Utils module
    Dim newTabs As Collection
    Set newTabs = CreateCoordinatorTabs_newTabs()
    
    ' Iterate through the newly created tabs
    Dim newTabName As Variant ' Declare as Variant (or Object)
    For Each newTabName In newTabs
        Debug.Print "Running CreatePromotorTabs for " & newTabName
        
        ' Activate the new coordinator tab
        Set ws = ThisWorkbook.Sheets(newTabName)
        ws.Activate
        
        ' Run CreatePromotorTabs for the current coordinator tab
        Call CreatePromotorTabs
        
        Debug.Print "Finished CreatePromotorTabs for " & newTabName
    Next newTabName
    
    ' Call SumPagoNetoGerencia and pass the activating sheet
    Call SumPagoNetoGerencia(activatingSheet)
    
    Debug.Print "Finished CreateCoordinatorAndPromotorTabs."
    
    Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "CreateCoordinatorAndPromotorTabs"
    End If
End Sub

