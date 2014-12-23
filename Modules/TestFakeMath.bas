Attribute VB_Name = "TestFakeMath"
Public Sub TestAddition()
    VaseAssert.AssertEqual 1 + 0, 1
End Sub

Public Sub TestSubtraction()
    VaseAssert.AssertEqual 2 - 1, 1
End Sub

