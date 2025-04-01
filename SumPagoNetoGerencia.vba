'=======================================================================
' Script: SumPagoNetoGerencia
' Version: 1.5.7
' Author: Juan Pablo Garcia Murillo
' Date: 04/01/2025
' Description: 
'   This script calculates the total sum of "PAGO NETO" values from the 
'   "D" column across all visible worksheets in the workbook. It calls 
'   the function `SumPagoNetoFromSheets` to calculate the sum from 
'   all visible sheets and store the result in cell "J4" of the 
'   sheet that triggered the macro.
'=======================================================================

Sub SumPagoNetoGerencia()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    targetSheet.Range("J4").Value = SumPagoNetoFromSheets(Empty)
End Sub
