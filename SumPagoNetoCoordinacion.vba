'=======================================================================
' Subroutine: SumPagoNetoCoordinacion
' Version: 1.6.5
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This subroutine sums the "PAGO NETO" values from multiple sheets
'   and stores the total in cell J4 of the target sheet. It first checks
'   for a list of sheet names in column P of the active sheet. If sheet
'   names are provided, it sums the "PAGO NETO" values from those sheets.
'   If no sheet names are listed, it sums the "PAGO NETO" values from
'   the active sheet only. The subroutine also handles the case where
'   the active sheet is not included in the list, adding its values to
'   the sum separately.
' Parameters:
'   - None
' Returns:
'   - None
' Notes:
'   - The "PAGO NETO" values are assumed to be located in a consistent
'     location across the sheets being summed.
'   - If no sheet names are provided in column P, the subroutine will
'     sum the "PAGO NETO" values from the active sheet only.
'   - If the active sheet is not included in the list of sheet names,
'     its values will still be added to the sum.
'   - The total sum is stored in cell J4 of the target sheet.
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
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "P").End(xlUp).row

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
