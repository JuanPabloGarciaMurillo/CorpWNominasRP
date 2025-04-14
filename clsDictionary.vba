Private dictKeys As Collection
Private dictValues As Collection

Private Sub Class_Initialize()
    ' Initialize the collections for keys and values
    Set dictKeys = New Collection
    Set dictValues = New Collection
End Sub

' Add a key-value pair to the dictionary
Public Sub Add(key As Variant, value As Variant)
    If Exists(CStr(key)) Then
        Err.Raise vbObjectError + 1, "clsDictionary", "Key already exists: " & CStr(key)
    End If
    dictKeys.Add CStr(key)
    dictValues.Add value
End Sub

' Check if a key exists in the dictionary
Public Function Exists(key As Variant) As Boolean
    Exists = (FindKeyIndex(CStr(key)) > 0)
End Function

' Get the value associated with a key
Public Function GetValue(key As String) As Variant
    Dim index As Long
    index = FindKeyIndex(key)
    If index > 0 Then
        GetValue = dictValues(index)
    Else
        Err.Raise vbObjectError + 2, "clsDictionary", "Key not found: " & key
    End If
End Function

' Remove a key-value pair from the dictionary
Public Sub Remove(key As String)
    Dim index As Long
    index = FindKeyIndex(key)
    If index > 0 Then
        dictKeys.Remove index
        dictValues.Remove index
    Else
        Err.Raise vbObjectError + 3, "clsDictionary", "Key not found: " & key
    End If
End Sub

' Replace the value for an existing key
Public Sub Replace(key As String, value As Variant)
    Dim index As Long
    index = FindKeyIndex(key)
    If index > 0 Then
        dictValues(index) = value
    Else
        Err.Raise vbObjectError + 4, "clsDictionary", "Key not found: " & key
    End If
End Sub

' Get all keys as an array
Public Function GetKeys() As Variant
    Dim keys() As String
    ReDim keys(1 To dictKeys.Count)
    Dim i As Long
    For i = 1 To dictKeys.Count
        keys(i) = dictKeys(i)
    Next i
    GetKeys = keys
End Function

' Get all values as an array
Public Function GetValues() As Variant
    Dim values() As Variant
    ReDim values(1 To dictValues.Count)
    Dim i As Long
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
    Dim i As Long
    For i = 1 To dictKeys.Count
        If dictKeys(i) = key Then
            FindKeyIndex = i
            Exit Function
        End If
    Next i
    FindKeyIndex = 0 ' Key not found
End Function