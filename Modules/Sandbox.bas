Attribute VB_Name = "Sandbox"
Public Sub X()
    Dim Y As Variant
    For Each Y In Array(1)
        Debug.Print 15
    Next
End Sub

Public Sub Y(Cond As Boolean)
    Debug.Print 1
End Sub

Public Sub Z()
    Y (1 = "A")
End Sub
