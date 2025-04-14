'==================================================
' Script: UtilsData
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This module contains utility functions for working with data in Excel worksheets. It includes functions for summing specific values across multiple sheets, checking if a row is empty, and other data manipulation tasks. The module is designed to help streamline data processing and validation in Excel.
' Functions included in this module:
'   - SumPagoNetoFromSheets
'   - IsRowEmpty
'==================================================

'==================================================
' Function: SumPagoNetoFromSheets
' Description:
'   This function calculates the total "PAGO NETO" value from column D across multiple worksheets.
'   It searches for the term "PAGO NETO" in column A and sums the corresponding values in column D.
' Parameters:
'   - sheetNames (Variant): An array of sheet names to process.
'       - If empty, all visible sheets in the workbook are processed.
' Returns:
'   - Currency: The total sum of "PAGO NETO" values across the specified sheets.
' Notes:
'   - Only sums values where "PAGO NETO" appears in column A.
'   - If `sheetNames` is empty, it processes all visible sheets in the workbook.
'==================================================
Public Function SumPagoNetoFromSheets(sheetNames As Variant) As Currency
    Dim ws          As Worksheet
    Dim totalPagoNeto As Currency
    Dim lastRow     As Long
    Dim rngA        As Range
    Dim dataA       As Variant, dataD As Variant
    Dim i           As Long
    Dim processAllSheets As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Initialize variables
    totalPagoNeto = 0
    processAllSheets = IsEmpty(sheetNames)        ' If sheetNames is empty, process all visible sheets
    
    ' Turn off screen updating and calculation for performance
    Application.ScreenUpdating = FALSE
    Application.Calculation = xlCalculationManual
    
    ' Loop through the sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Skip irrelevant sheets
        Select Case ws.Name
            Case "Premios", "Planteles", "Tabuladores", "Colaboradores", "Ejemplo Coordinacion", "Ejemplo Promotor"
                GoTo NextSheet
        End Select
        
        ' If processing all visible sheets OR the sheet is in the provided list, proceed
        If (processAllSheets And ws.Visible = xlSheetVisible) Or (Not processAllSheets And Not IsError(Application.Match(ws.Name, sheetNames, 0))) Then
            ' Find the last row in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            If lastRow > 0 Then
                ' Read data into arrays for faster processing
                dataA = ws.Range("A1:A" & lastRow).Value
                dataD = ws.Range("D1:D" & lastRow).Value
                
                ' Loop through the data in column A
                For i = 1 To UBound(dataA, 1)
                    If UCase(dataA(i, 1)) = "PAGO NETO" Then
                        ' Ensure the corresponding value in column D is numeric
                        If IsNumeric(dataD(i, 1)) Then
                            totalPagoNeto = totalPagoNeto + dataD(i, 1)
                        End If
                    End If
                Next i
            End If
        End If
        
        NextSheet:
    Next ws
    
    ' Restore screen updating and calculation
    Application.ScreenUpdating = TRUE
    Application.Calculation = xlCalculationAutomatic
    
    ' Return the total sum
    SumPagoNetoFromSheets = totalPagoNeto
    Exit Function
    
    ErrorHandler:
    Debug.Print "Error in sheet: " & ws.Name & " - " & Err.Description
    SumPagoNetoFromSheets = 0
    Application.ScreenUpdating = TRUE
    Application.Calculation = xlCalculationAutomatic
End Function

'==================================================
' Function: IsRowEmpty
' Description:
'   This function checks if a specified row in a worksheet is empty, ignoring the first column.
'   It iterates through all columns except the first and returns True if no data is found.
' Parameters:
'   - ws (Worksheet): The worksheet containing the row to check.
'   - rowNum (Long): The row number to evaluate.
' Returns:
'   - Boolean: True if all checked columns are empty, otherwise False.
' Notes:
'   - Assumes the worksheet contains at least one table (ListObject) to determine the last column.
'   - Only considers columns from the second column onward.
'==================================================
' Function to check if the row is empty (ignores only the first column)
Public Function IsRowEmpty(ws As Worksheet, rowNum As Long) As Boolean
    Dim col         As Long
    Dim lastColumn  As Long
    
    ' Check if the worksheet contains any tables
    If ws.ListObjects.Count = 0 Then
        IsRowEmpty = TRUE        ' Assume row is empty if no table exists
        Exit Function
    End If
    
    lastColumn = ws.ListObjects(1).ListColumns.Count
    
    ' Check all columns except for the first column
    For col = 2 To lastColumn        ' Skip the first column (column 1)
        If ws.Cells(rowNum, col).Value <> "" Then
            IsRowEmpty = FALSE
            Exit Function
        End If
    Next col
    
    ' If none of the columns except the first column have data, the row is considered empty
    IsRowEmpty = TRUE
End Function


