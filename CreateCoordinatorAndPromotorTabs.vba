'=========================================================
' Subroutine: CreateCoordinatorAndPromotorTabs
' Version: 0.9.2
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
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
'=========================================================

Sub CreateCoordinatorAndPromotorTabs()
    On Error GoTo ErrHandler
    
    ' Capture the sheet that activated the macro
    Dim activatingSheet As Worksheet
    Set activatingSheet = ActiveSheet
    
    ' Define the list of protected tabs
    Dim protectedTabs As Variant
    protectedTabs = Split(SKIP_SHEETS, ",")
    ReDim Preserve protectedTabs(UBound(protectedTabs) + 1)
    protectedTabs(UBound(protectedTabs)) = activatingSheet.Name
    
    ' Delete all unprotected tabs
    Call DeleteUnprotectedTabs(protectedTabs)
    
    ' Run CreateCoordinatorTabs to create new coordinator tabs
    Call CreateCoordinatorTabs
    
    ' Now retrieve the newTabs collection from the Utils module
    Dim newTabs     As Collection
    Set newTabs = CreateCoordinatorTabs_newTabs()
    
    ' Iterate through the newly created tabs
    Dim newTabName  As Variant        ' Declare as Variant (or Object)
    For Each newTabName In newTabs
        
        ' Activate the new coordinator tab
        Set ws = ThisWorkbook.Sheets(newTabName)
        ws.Activate
        
        ' Run CreatePromotorTabs for the current coordinator tab
        Call CreatePromotorTabs()
        
    Next newTabName
    
    ' Delete tabs ending with "(C)"
    Call DeleteManagerCoordinatorTab
    
    ' Call SumPagoNetoGerencia and pass the activating sheet
    Call SumPagoNetoGerencia(activatingSheet)
    
    Exit Sub
    
    ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error in CreateCoordinatorAndPromotorTabs: " & Err.Description
        HandleError ERROR_GENERIC & " " & Err.Number & ": " & Err.Description, "CreateCoordinatorAndPromotorTabs"
    End If
End Sub