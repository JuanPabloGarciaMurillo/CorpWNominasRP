'=======================================================================
' Function: SumPagoNetoFromSheets
' Version: 1.5.7
' Author: Juan Pablo Garcia Murillo
' Date: 04/01/2025
' Description: 
'   This function calculates the total sum of "PAGO NETO" values from column "D" across a list of specified sheets.
'   It checks each sheet for the presence of the "PAGO NETO" label in column "A" and sums the corresponding values in column "D".
'   If no sheet names are provided, it processes all visible sheets in the workbook.
'   The result is returned as a Currency type value.
'=======================================================================

Function SumPagoNetoFromSheets(sheetNames As Variant) As Currency
    Dim ws As Worksheet
    Dim totalPagoNeto As Currency
    Dim lastRow As Long
    Dim rngA As Range, rngD As Range
    Dim i As Long
    Dim processAllSheets As Boolean
    
    totalPagoNeto = 0 ' Initialize sum
    processAllSheets = IsEmpty(sheetNames) ' If sheetNames is empty, process all visible sheets
    
    ' Loop through the sheets
    For Each ws In ThisWorkbook.Worksheets
        ' If processing all visible sheets OR the sheet is in the provided list, proceed
        If (processAllSheets And ws.Visible = xlSheetVisible) Or (Not processAllSheets And Not IsError(Application.Match(ws.Name, sheetNames, 0))) Then
            ' Find the last row in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            If lastRow >= 1 Then
                ' Set ranges for columns A and D
                Set rngA = ws.Range("A1:A" & lastRow)
                Set rngD = ws.Range("D1:D" & lastRow)
                
                ' Loop through column A, summing values in column D where "PAGO NETO" is found
                For i = 1 To rngA.Rows.Count
                    If rngA.Cells(i, 1).Value = "PAGO NETO" Then
                        totalPagoNeto = totalPagoNeto + rngD.Cells(i, 1).Value
                    End If
                Next i
            End If
        End If
    Next ws
    
    SumPagoNetoFromSheets = totalPagoNeto ' Return total sum
End Function
