' Class Module: clsDictionary
' Version: 0.9.3
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   A custom dictionary-like class that stores key-value pairs using internal collections. Supports basic dictionary operations such as:
'     - Add, Remove, Replace, Exists
'     - Retrieve keys or values
'     - Count and Clear
'     - Error handling for duplicate or missing keys
' Properties:
'   - dictKeys (Collection): Internal collection for storing string keys.
'   - dictValues (Collection): Internal collection for storing values.
' Public Methods:
'   - Add(key, value): Adds a new key-value pair. Raises error if key exists.
'   - Exists(key): Returns True if the key exists.
'   - GetValue(key): Returns the value for a given key. Raises error if not found.
'   - Remove(key): Removes a key-value pair. Raises error if not found.
'   - Replace(key, value): Replaces value for existing key. Raises error if not found.
'   - GetKeys(): Returns an array of all keys.
'   - GetValues(): Returns an array of all values.
'   - Clear(): Clears all key-value pairs.
'   - Count(): Returns the number of items in the dictionary.
' Private Methods:
'   - FindKeyIndex(key): Returns the index of a key in dictKeys, or 0 if not found.
' Notes:
'   - Keys are treated as strings and stored in insertion order.
'   - Duplicate keys are not allowed.
'   - This class mimics basic Dictionary behavior in environments where the Scripting.Dictionary object is unavailable or undesired.

Private dictKeys    As Collection
Private dictValues  As Collection

Private Sub Class_Initialize()
    ' Initialize the collections for keys and values
    Set dictKeys = New Collection
    Set dictValues = New Collection
End Sub

' Add a key-value pair to the dictionary
Public Sub Add(key  As Variant, value As Variant)
    If Exists(CStr(key)) Then
        Err.Raise vbObjectError + 1, "clsDictionary", ERROR_KEY_EXISTS & CStr(key)
    End If
    dictKeys.Add CStr(key)
    dictValues.Add value
End Sub

' Check if a key exists in the dictionary
Public Function Exists(key As Variant) As Boolean
    Exists = (FindKeyIndex(CStr(key)) > 0)
End Function

' Get the value associated with a key
Public Function GetValue(key As Variant) As Variant
    Dim index       As Long
    ' Convert the key to a string for internal processing
    index = FindKeyIndex(CStr(key))
    If index > 0 Then
        GetValue = dictValues(index)
    Else
        Err.Raise vbObjectError + 4, "clsDictionary", ERROR_KEY_NOT_FOUND & CStr(key)
    End If
End Function

' Remove a key-value pair from the dictionary
Public Sub Remove(key As String)
    Dim index       As Long
    index = FindKeyIndex(key)
    If index > 0 Then
        dictKeys.Remove index
        dictValues.Remove index
    Else
        Err.Raise vbObjectError + 3, "clsDictionary", "Key Not found: " & key
    End If
End Sub

' Replace the value for an existing key
Public Sub Replace(key As String, value As Variant)
    Dim index       As Long
    index = FindKeyIndex(key)
    If index > 0 Then
        dictValues(index) = value
    Else
        Err.Raise vbObjectError + 4, "clsDictionary", "Key Not found: " & key
    End If
End Sub

' Get all keys      as an array
Public Function GetKeys() As Variant
    Dim keys()      As String
    ReDim keys(1 To dictKeys.Count)
    Dim i           As Long
    For i = 1 To dictKeys.Count
        keys(i) = dictKeys(i)
    Next i
    GetKeys = keys
End Function

' Get all values    as an array
Public Function GetValues() As Variant
    Dim values()    As Variant
    ReDim values(1 To dictValues.Count)
    Dim i           As Long
    For i = 1 To dictValues.Count
        values(i) = dictValues(i)
    Next i
    GetValues = values
End Function

' Clear all key-value pairs from the dictionary
Public Sub Clear()
    Set dictKeys = New Collection
    Set dictValues = New Collection
End Sub

' Get the number of key-value pairs in the dictionary
Public Function Count() As Long
    Count = dictKeys.Count
End Function

' Helper function to find the index of a key
Private Function FindKeyIndex(key As String) As Long
    Dim i           As Long
    For i = 1 To dictKeys.Count
        If dictKeys(i) = key Then
            FindKeyIndex = i
            Exit Function
        End If
    Next i
    FindKeyIndex = 0
End Function