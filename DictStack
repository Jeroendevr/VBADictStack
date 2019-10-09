Option Explicit
    
Dim dict

Private Sub Class_Initialize()
    Set dict = CreateObject("scripting.dictionary")
End Sub
Public Sub push(key As Variant, value As Variant)
    dict.Add key, value
End Sub

Public Function pop() As Variant
    Dim key As Variant
    Dim value As Variant
    Dim keys()
    
    'First check if the array is not empty
    If is_not_empty() Then
        key = first_key
        value = dict.Item(key)
        pop = Array(key, value)
        dict.Remove key
    End If
End Function

Private Function is_not_empty() As Boolean
    If dict.Count = 0 Then
        is_not_empty = False
    Else
        is_not_empty = True
    End If
End Function

Private Function first_key() As Variant
    Dim keys() As Variant
    keys = dict.keys
    first_key = keys(0)
End Function

