' Script: UtilsCollections
' Version: 0.9.3
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This module contains utility functions for working with collections in Excel VBA.
'   It includes functions for checking if a key exists in a collection, checking if a value exists in an array, and validating input values.'
' Functions included in this module:
'   - KeyExists
'   - IsInArray
'   - CollectionToArray
'   - InitializeHeaderMapping

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

' Function: IsInArray
' Description:
'   Checks if a value exists in an array.
' Parameters:
'   - value (Variant): The value to check for in the array.
'   - arr (Variant): The array to check against.
'   - caseInsensitive (Boolean, Optional): Whether to perform a case-insensitive check. Default is False.
' Returns:
'   - True if the value exists in the array, False otherwise.
' Notes:
'   - Uses On Error Resume Next to suppress errors.
'   - The function is case-sensitive by default.
'   - The function converts the value to a string before checking.

Public Function IsInArray(value As Variant, arr As Variant, Optional caseInsensitive As Boolean = False) As Boolean
    Dim i           As Long
    Dim strValue    As String
    Dim arrayValue  As String
    
    ' Convert the value to a string
    strValue = CStr(value)
    If caseInsensitive Then strValue = LCase(strValue)
    
    ' Check if the value exists in the array
    For i = LBound(arr) To UBound(arr)
        arrayValue = CStr(arr(i))
        If caseInsensitive Then arrayValue = LCase(arrayValue)
        
        If arrayValue = strValue Then
            IsInArray = TRUE
            Exit Function
        End If
    Next i
    IsInArray = FALSE
End Function

' Function: CollectionToArray
' Description:
'    Converts a Collection to an array.
' Parameters:
'   - coll (Collection): The collection to convert.
' Returns:
'   - arr (Variant): The converted array.
' Notes:
'   - The function assumes the collection contains string values.
'   - The function uses ReDim to create a dynamic array based on the collection count.
'   - The function iterates through the collection and assigns each value to the array.

Public Function CollectionToArray(coll As Collection) As Variant
    Dim arr()       As String
    Dim i           As Integer
    
    ReDim arr(1 To coll.Count)
    For i = 1 To coll.Count
        arr(i) = coll(i)
    Next i
    
    CollectionToArray = arr
End Function

' Function: InitializeHeaderMapping
' Description:
'     Initializes a dictionary with header mappings.
' Parameters:
'   - HEADERS (String): A comma-separated string of headers.
'   - COLUMN_INDICES (String): A comma-separated string of column indices.
'   - headerMapping (clsDictionary): The dictionary to store the header mappings.
' Notes:
'   - The function splits the headers and column indices into arrays.
'   - It loops through the headers and adds them to the dictionary.
'   - The function checks if the header already exists in the dictionary before adding it.

Public Sub InitializeHeaderMapping(HEADERS As String, COLUMN_INDICES As String, headerMapping As clsDictionary)
    Dim headersArray As Variant
    Dim columnIndex As Variant
    Dim idx As Integer
    
    ' Split the headers and column indices into arrays
    headersArray = Split(HEADERS, ",")
    columnIndex = Split(COLUMN_INDICES, ",")
    
    ' Loop through the headers and add them to the dictionary
    For idx = LBound(headersArray) To UBound(headersArray)
        If Not headerMapping.Exists(CStr(headersArray(idx))) Then
            headerMapping.Add CStr(headersArray(idx)), columnIndex(idx)
        End If
    Next idx
End Sub