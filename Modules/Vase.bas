Attribute VB_Name = "Vase"
'=======================
'--- Globals         ---
'=======================

'# This global variable indicates if the main test runner is active
'# This also allows output for solo assertions the runner is not active
'! Make sure this is only True while RunTests is executing
Private gIsRunningTest As Boolean
Public Property Get IsRunningTest() As Boolean
    IsRunningTest = gIsRunningTest
End Property

'=======================
'--- User Function   ---
'=======================

'# Like Nose, this starts test discovery and runs each test found by it
Public Sub RunTests()
On Error GoTo ErrHandler
    gIsRunningTest = True
    VaseLib.ClearScreen
    
    Debug.Print "Vase Test Framework"
    Debug.Print "Don't break the vase."
    Debug.Print "======================="
    
    VaseLib.RunVaseSuite ActiveWorkbook, Verbose:=True ' The output result is printed out, so no need to capture the output
    
    Debug.Print "Vase was filled"
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print _
            "Whoops! There was error running Vase. Check if you put the water in the vase correctly."
    End If
    Err.Clear
    gIsRunningTest = False
End Sub
