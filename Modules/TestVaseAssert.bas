Attribute VB_Name = "TestVaseAssert"
Public Sub TestSoloExecutionPass()
    VaseAssert.AssertTrue True ' Visual test
    
    VaseAssert.Ping_
End Sub

Public Sub TestSoloExecutionFailed()
    VaseAssert.AssertFalse True, "Sample Message" ' Visual test
    
    VaseAssert.Ping_
End Sub

Public Sub TestArrayEquals()
    Dim Arr1 As Variant, Arr2 As Variant
    Arr1 = Array(1, 2, 3)
    Arr2 = Arr1
    
    VaseAssert.AssertEqualArrays Arr1, Arr2

End Sub
