'==================================================
' Script: UtilsCollections
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This module contains utility functions for working with collections in Excel VBA.
'   It includes functions for checking if a key exists in a collection, checking if a value exists in an array, and validating input values.'
' Functions included in this module:
'   - KeyExists
'   - IsInArray
'==================================================

'====================================================
' Function: KeyExists
' Description:
'   Checks if a key exists in a collection.
' Parameters:
'   - col (Collection): The collection to check for the key.
'   - key (String): The key to check for in the collection.
' Returns:
'   - True if the sheet exists, False otherwise.
' Notes:
'   - Uses On Error Resume Next to suppress errors.
'   - Returns True if the key exists, False otherwise.
'   - The function is case-sensitive.
'====================================================

Public Function KeyExists(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim temp        As Variant
    
    ' Check if the collection is initialized
    If col Is Nothing Then
        KeyExists = FALSE
        Exit Function
    End If
    
    ' Check if the key exists in the collection
    temp = col(key)
    KeyExists = (Err.Number = 0)
    On Error GoTo 0
End Function

'====================================================
' Function: IsInArray
' Description:
'   Checks if a value exists in an array.
' Parameters:
'   - value (Variant): The value to check for in the array.
'   - arr (Variant): The array to check against.
' Returns:
'   - True if the value exists in the array, False otherwise.
' Notes:
'   - Uses On Error Resume Next to suppress errors.
'   - The function is case-sensitive.
'   - The function converts the value to a string before checking.
'====================================================
Public Function IsInArray(value As Variant, arr As Variant) As Boolean
    Dim i           As Long
    Dim strValue    As String
    
    ' Convert the value to a string
    strValue = CStr(value)
    
    ' Check if the value exists in the array
    For i = LBound(arr) To UBound(arr)
        If arr(i) = strValue Then
            IsInArray = TRUE
            Exit Function
        End If
    Next i
    IsInArray = FALSE
End Function