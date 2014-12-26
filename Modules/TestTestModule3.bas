Attribute VB_Name = "TestTestModule3"
Private gUpTime As Date
Private gDownTime As Date

Public Sub Setup()
    gUpTime = Now
End Sub

Public Sub TestSetupTeardown()
    VaseAssert.AssertGreaterThan gUpTime, gDownTime
    VaseAssert.Ping_
End Sub

Public Sub Teardown()
    gDownTime = Now
End Sub
