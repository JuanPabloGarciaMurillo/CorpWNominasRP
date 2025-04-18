'=========================================================
' Subroutine: SumPagoNetoGerencia
' Version: 0.9.0
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This subroutine calculates the sum of the "PAGO NETO" values
'   across all sheets in the workbook and stores the total in cell
'   J4 of the specified target sheet. It calls the `SumPagoNetoFromSheets`
'   function without specifying any specific sheets, implying it sums
'   from all available sheets in the workbook.
' Parameters:
'   - targetSheet (Worksheet): The sheet where the calculated total
'     will be placed in cell J4.
' Returns:
'   - None
' Notes:
'   - The "PAGO NETO" values are assumed to be in a consistent location
'     across all sheets being summed.
'   - The total sum is stored in cell J4 of the specified target sheet.
'=========================================================

Public Sub SumPagoNetoGerencia(targetSheet As Worksheet)
    On Error GoTo ErrorHandler
    Dim managerPagoNeto As Currency
    Dim total As Currency

    ' Validate targetSheet
    If targetSheet Is Nothing Then
        MsgBox ERROR_INVALID_SHEET, vbExclamation, "Error"
        Exit Sub
    End If

    ' Perform the calculation for all sheets
    total = SumPagoNetoFromSheets(Empty)

    ' Add the "PAGO NETO" value from the manager sheet
    managerPagoNeto = GetManagerPagoNeto(targetSheet)

    ' Add the manager's "PAGO NETO" to the total
    total = total + managerPagoNeto

    ' Store the result in the target cell
    targetSheet.Range(TARGET_CELL).Value = total

    Exit Sub

ErrorHandler:
    MsgBox ERROR_GENERIC & Err.Description, vbCritical, "Error", "SumPagoNetoGerencia"
End Sub