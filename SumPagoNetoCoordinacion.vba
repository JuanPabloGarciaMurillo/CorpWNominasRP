' Subroutine: SumPagoNetoCoordinacion
' Version: 0.9.3
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This subroutine sums the "PAGO NETO" values from multiple sheets and stores the total in cell J4 of the target sheet. It first checks for a list of sheet names in column P of the active sheet. If sheet names are provided, it sums the "PAGO NETO" values from those sheets.
'   If no sheet names are listed, it sums the "PAGO NETO" values from the active sheet only. The subroutine also handles the case where the active sheet is not included in the list, adding its values to
'   the sum separately.
' Notes:
'   - The "PAGO NETO" values are assumed to be located in a consistent location across the sheets being summed.
'   - If no sheet names are provided in column P, the subroutine will sum the "PAGO NETO" values from the active sheet only.
'   - If the active sheet is not included in the list of sheet names, its values will still be added to the sum.
'   - The total sum is stored in cell J4 of the target sheet.

Public Sub SumPagoNetoCoordinacion(newTabs As Collection, targetSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim totalPagoNeto As Currency
    Dim sheetNames() As String
    Dim i           As Long
    
    ' Validate targetSheet
    If targetSheet Is Nothing Then
        MsgBox "Error: Target sheet Is Not specified.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Convert the Collection (newTabs) to an array of sheet names
    ReDim sheetNames(1 To newTabs.Count + 1)
    For i = 1 To newTabs.Count
        sheetNames(i) = newTabs(i)
    Next i
    
    ' Add the Coordinator Tab (targetSheet) to the array
    sheetNames(newTabs.Count + 1) = targetSheet.Name
    
    ' Call SumPagoNetoFromSheets with the array of sheet names
    totalPagoNeto = SumPagoNetoFromSheets(sheetNames)
    
    ' Store the total sum in TARGET_CELL of the Coordinator Tab
    targetSheet.Range(TARGET_CELL).Value = totalPagoNeto
    
    Exit Sub
    
    ErrorHandler:
    MsgBox "Error in SumPagoNetoCoordinacion: " & Err.Description, vbCritical, "Error"
End Sub