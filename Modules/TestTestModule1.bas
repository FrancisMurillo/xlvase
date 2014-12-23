Attribute VB_Name = "TestTestModule1"
Public Sub TestFailCompute()
    AssertTrue False
End Sub

Public Sub TestSuccessCompute()
    AssertFalse False
End Sub

Public Sub NotTested()
    AssertTrue True
End Sub



