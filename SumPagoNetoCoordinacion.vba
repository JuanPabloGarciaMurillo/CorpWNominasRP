'=======================================================================
' Script: SumPagoNetoCoordinacion
' Version: 1.6.3
' Author: Juan Pablo Garcia Murillo
' Date: 04/01/2025
' Description:
'   This script calculates the total sum of "PAGO NETO" values from column "D" across a list of specified sheets, as indicated in column "P" of the triggering sheet.
'   It checks each listed sheet for the presence of the "PAGO NETO" label in column "A", sums the corresponding values in column "D", and then stores the total sum in cell "J4" of the triggering sheet.
'   If the active sheet is not listed in column "P", it processes the active sheet separately.
'=======================================================================

Sub SumPagoNetoCoordinacion()
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim nameRange As Range
    Dim sheetNames() As String
    Dim i As Integer
    Dim currentSheetIncluded As Boolean
    Dim totalPagoNeto As Currency

    Set targetSheet = ActiveSheet
    currentSheetIncluded = False
    totalPagoNeto = 0 ' Initialize sum

    ' Find last row in column P
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "P").End(xlUp).Row

    ' Check if there are any sheet names listed
    If lastRow < 2 Then
        ' If no sheets are listed, only sum from the active sheet
        totalPagoNeto = SumPagoNetoFromSheets(Array(targetSheet.Name))
        targetSheet.Range("J4").Value = totalPagoNeto
        Exit Sub
    End If

    ' Store sheet names into an array
    Set nameRange = targetSheet.Range("P2:P" & lastRow)
    ReDim sheetNames(1 To nameRange.Cells.Count)

    i = 1
    For Each cell In nameRange
        If cell.Value <> "" Then
            sheetNames(i) = cell.Value
            ' Check if the active sheet is listed
            If cell.Value = targetSheet.Name Then
                currentSheetIncluded = True
            End If
            i = i + 1
        End If
    Next cell

    ' Sum from listed sheets
    totalPagoNeto = SumPagoNetoFromSheets(sheetNames)

    ' If the active sheet is NOT listed, sum its "PAGO NETO" values separately
    If Not currentSheetIncluded Then
        totalPagoNeto = totalPagoNeto + SumPagoNetoFromSheets(Array(targetSheet.Name))
    End If

    ' Store total sum in J4
    targetSheet.Range("J4").Value = totalPagoNeto
End Sub