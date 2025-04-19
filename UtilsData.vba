'=========================================================
' Script: UtilsData
' Version: 0.9.1
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This module contains utility functions for working with data in Excel worksheets. It includes functions for summing specific values across multiple sheets, checking if a row is empty, and other data manipulation tasks. The module is designed to help streamline data processing and validation in Excel.
' Functions included in this module:
'   - SumPagoNetoFromSheets
'   - IsRowEmpty
'   - GetManagerPagoNeto
'=========================================================

'=========================================================
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
'=========================================================

Public Function SumPagoNetoFromSheets(sheetNames As Variant) As Currency
    Dim ws          As Worksheet
    Dim totalPagoNeto As Currency
    Dim lastRow     As Long
    Dim rngA        As Range
    Dim dataA       As Variant, dataD As Variant
    Dim i           As Long
    Dim processAllSheets As Boolean
    Dim skipSheets() As String
    
    On Error GoTo ErrorHandler

    ' Initialize variables
    totalPagoNeto = 0
    processAllSheets = IsEmpty(sheetNames)
    skipSheets = Split(SKIP_SHEETS, ",")
    
    ' Turn off screen updating and calculation for performance
    Application.ScreenUpdating = FALSE
    Application.Calculation = xlCalculationManual
    
    ' Loop through the sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Skip irrelevant sheets
        If Not IsError(Application.Match(ws.Name, skipSheets, 0)) Then GoTo NextSheet
        
        ' If processing all visible sheets OR the sheet is in the provided list, proceed
        If (processAllSheets And ws.Visible = xlSheetVisible) Or (Not processAllSheets And Not IsError(Application.Match(ws.Name, sheetNames, 0))) Then
            ' Find the last row in column A
            lastRow = ws.Cells(ws.Rows.Count, COLUMN_A).End(xlUp).Row
            
            If lastRow > 0 Then
                ' Read data into arrays for faster processing
                dataA = ws.Range("A1:A" & lastRow).Value
                dataD = ws.Range("D1:D" & lastRow).Value
                
                ' Loop through the data in column A
                For i = 1 To UBound(dataA, 1)
                    If UCase(dataA(i, 1)) = PAGO_NETO_TEXT Then
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

'=========================================================
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
'=========================================================

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
    For col = START_COLUMN To lastColumn        ' Skip the first column (column 1)
        If ws.Cells(rowNum, col).Value <> "" Then
            IsRowEmpty = FALSE
            Exit Function
        End If
    Next col
    
    ' If none of the columns except the first column have data, the row is considered empty
    IsRowEmpty = TRUE
End Function

'=========================================================
' Function: GetManagerPagoNeto
' Description:
'   This function retrieves the "PAGO NETO" value from the manager sheet.
'   It searches for the term "PAGO NETO" in column A and returns the corresponding value from column E.
' Parameters:
'   - managerSheet (Worksheet): The worksheet containing the "PAGO NETO" data.
' Returns:
'   - Currency: The "PAGO NETO" value from column E, or 0 if not found.
' Notes:
'   - Assumes the "PAGO NETO" text is in uppercase.
'   - If the value in column E is not numeric, it returns 0.
'=========================================================
Public Function GetManagerPagoNeto(managerSheet As Worksheet) As Currency
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim pagoNetoValue As Currency

    On Error GoTo ErrorHandler

    ' Find the last row in column A
    lastRow = managerSheet.Range(COLUMN_A & managerSheet.Rows.Count).End(xlUp).Row

    ' Loop through column A to find "PAGO NETO"
    For rowIndex = 1 To lastRow
        If UCase(managerSheet.Range(COLUMN_A & rowIndex).Value) = PAGO_NETO_TEXT Then
            ' Get the value from column E in the same row
            If IsNumeric(managerSheet.Range(COLUMN_E & rowIndex).Value) Then
                pagoNetoValue = managerSheet.Range(COLUMN_E & rowIndex).Value
            End If
            Exit For
        End If
    Next rowIndex

    ' Return the value
    GetManagerPagoNeto = pagoNetoValue
    Exit Function

ErrorHandler:
    Debug.Print "Error in GetManagerPagoNeto: " & Err.Description
    HandleError ERROR_GENERIC & " " & Err.Number & ": " & Err.Description, "GetManagerPagoNeto"
    GetManagerPagoNeto = 0
End Function

'=========================================================
' Subroutine: StoreTotalInTargetCell
' Description:
'    This subroutine stores the total sum in a specified target cell.
' Parameters:
'   - targetSheet (Worksheet): The worksheet where the total will be stored.
'   - total (Currency): The total sum to be stored.
' Returns:
'   - None
' Notes:
'   - The target cell is defined by the constant TARGET_CELL.
'   - The function assumes the target cell is in the specified worksheet.
'=========================================================
' Subroutine to store the total sum in the target cell
Public Sub StoreTotalInTargetCell(targetSheet As Worksheet, total As Currency)
    targetSheet.Range(TARGET_CELL).Value = total
End Sub