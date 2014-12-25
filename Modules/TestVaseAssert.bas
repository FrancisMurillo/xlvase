Attribute VB_Name = "TestVaseAssert"
Public Sub TestSoloExecutionPass()
    VaseAssert.AssertTrue True ' Visual test
    
    VaseAssert.Ping_
End Sub

Public Sub TestSoloExecutionFailed()
    VaseAssert.AssertFalse True, "Sample Message" ' Visual test
    
    VaseAssert.Ping_
End Sub
