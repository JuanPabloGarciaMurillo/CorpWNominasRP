'=======================================================================
' Subroutine: SumPagoNetoGerencia
' Version: 1.6.5
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
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
'=======================================================================


Public Sub SumPagoNetoGerencia(targetSheet As Worksheet)
    targetSheet.Range("J4").Value = SumPagoNetoFromSheets(Empty)
End Sub


