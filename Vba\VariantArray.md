```vb
Function IsArrayInitialized(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayInitialized = Not IsError(LBound(arr, 1))
    On Error GoTo 0
End Function
```

```vb
If IsArrayInitialized(temp) Then
    Debug.Print UBound(temp, 1)
Else
    Debug.Print "配列は未初期化です"
End If
```