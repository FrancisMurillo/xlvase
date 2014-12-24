Attribute VB_Name = "TestVaseLib"
Public gTempVar As Variant

Public Sub TestRunVase()
    ' Do not test this as this recursively executes the test runner
End Sub


Public Sub TestDiscoveredModules()
On Error Resume Next
    Dim TModules As Variant, TModule As Variant
    TModules = VaseLib.FindTestModules(ActiveWorkbook)
    
    Dim ExpectedModules As Variant
    ExpectedModules = Array( _
        "TestFakeMath", _
        "TestTestModule1", _
        "TestTestModule2", _
        "TestVaseLib")
    For Each TModule In TModules
        VaseAssert.AssertInArray TModule.Name, ExpectedModules
    Next
End Sub

Public Sub TestDiscoveredMethods()
On Error Resume Next
    Dim TMod As VBComponent, TMethods As Variant, TMethod As Variant
    Dim ExpectedMethods As Variant
    Set TMod = ActiveWorkbook.VBProject.VBComponents("TestFakeMath")
    TMethods = VaseLib.FindTestMethods(TMod)
    
    ExpectedMethods = Array("TestAddition", "TestSubtraction")
    VaseAssert.AssertEqualArrays ExpectedMethods, TMethods
End Sub

Public Sub TestRunTestCase()
On Error Resume Next
    gTempVar = "Meow"
    VaseLib.RunTestMethod ActiveWorkbook, "TestVaseLib", "RunThis"
    AssertEqual "Roar", gTempVar
End Sub

Private Sub RunThis()
    gTempVar = "Roar"
End Sub

Public Sub TestZip()
On Error Resume Next
    Dim LeftArr As Variant, RightArr As Variant, ComboArr As Variant, Tuple As Variant
    LeftArr = Array(1, 2, 3)
    RightArr = Array(3, 1)
    ComboArr = VaseLib.Zip(LeftArr, RightArr)
    
    VaseAssert.AssertArraySize 2, ComboArr
    
    VaseAssert.AssertEqual ComboArr(0)(0), 1
    VaseAssert.AssertEqual ComboArr(0)(1), 3
    
    VaseAssert.AssertEqual ComboArr(1)(0), 2
    VaseAssert.AssertEqual ComboArr(1)(1), 1

    VaseAssert.AssertEmptyArray Zip(Array(), RightArr)
    VaseAssert.AssertEmptyArray Zip(LeftArr, Array())
End Sub
