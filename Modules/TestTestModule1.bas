Attribute VB_Name = "TestTestModule1"
Public Sub TestFailCompute()
    VaseAssert.AssertTrue False
End Sub

Public Sub TestSuccessCompute()
    VaseAssert.AssertFalse False
End Sub

Public Sub NotTested()
    VaseAssert.AssertTrue True
End Sub


Public Sub TestUncaughtException()
On Error Resume Next
    'Not handled yet or will be
    VaseAssert.AssertTrue 1 = "A"
End Sub
