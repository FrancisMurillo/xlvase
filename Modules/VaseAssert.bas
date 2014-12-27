Attribute VB_Name = "VaseAssert"
'# This global variable determines if an assert method failed or passed
Private gAssertion As Boolean
Private gFirstFailed As String
Private gFirstFailedMessage As String
Private gFirstFailedAssertMessage As String
Private gFirstFailedParent As String

Private gTestCaseStarted As Boolean ' Defaults to False

Public Property Get TestResult() As Boolean
    TestResult = gAssertion
End Property
Public Property Get FirstFailedTestMethod() As String
    FirstFailedTestMethod = gFirstFailed
End Property
Public Property Get FirstFailedTestMessage() As String
    FirstFailedTestMessage = gFirstFailedMessage
End Property
Public Property Get FirstFailedTestAssertMessage() As String
    FirstFailedTestAssertMessage = gFirstFailedAssertMessage
End Property
Public Property Get FirstFailedTestParentMethod() As String
    FirstFailedTestParentMethod = gFirstFailedParent
End Property

'=======================
'--- Assertion Tools ---
'=======================

'# Sets the Assertion globals for use
Public Sub InitAssert()
    gAssertion = True
    gFirstFailed = ""
    gFirstFailedMessage = ""
End Sub


'# Base Assert Method
Private Sub Assert_(Cond As Boolean, _
        Optional Message As String = "", _
        Optional AssertName As String = "Assert", _
        Optional AssertFailMessage As String = "Assert Failed", _
        Optional AssertParentName As String = "", _
        Optional Params As Variant = Empty) ' Name to avoid Debug.Assert conflict or confusion
    If Not Vase.IsRunningTest Then ' Allow output for solo execution
        If Not gTestCaseStarted Then
            VaseLib.ClearScreen
            Debug.Print "Test Case Started"
            Debug.Print "---------------------"
            
            InitAssert
            gTestCaseStarted = True
        End If
    
        If Cond Then
            If AssertParentName = "" Then
                Debug.Print "+ " & AssertName
            Else
                Debug.Print "+ " & AssertParentName
            End If
        Else
            If AssertParentName = "" Then
                Debug.Print "- " & AssertName & vbCrLf & _
                            "-> " & AssertFailMessage & _
                                IIf(Message <> "", vbCrLf & " ->> " & Message, "")
            Else
                Debug.Print "- " & AssertParentName & "(" & AssertName & ")" & vbCrLf & _
                            "-> " & AssertFailMessage & _
                                IIf(Message <> "", vbCrLf & " ->> " & Message, "")
            End If
        End If
    End If
    
    gAssertion = gAssertion And Cond ' Update assertion variable
    If Not Cond And gFirstFailed = "" Then ' Log the first fail condition for logging
        gFirstFailed = AssertName
        gFirstFailedMessage = Message
        gFirstFailedAssertMessage = AssertFailMessage
        gFirstFailedParent = AssertParentName
    End If
End Sub

Private Sub Fail_(Optional Message As String = "", _
        Optional AssertName As String = "Assert", _
        Optional AssertFailMessage As String = "Assert Failed", _
        Optional AssertParentName As String = "", _
        Optional Params As Variant = Empty)
    Assert_ False, _
        Message:=Message, _
        AssertName:=AssertName, _
        AssertFailMessage:=AssertFailMessage, _
        AssertParentName:=AssertParentName, _
        Params:=Params
End Sub

'# Resets the test case flag
Public Sub Rewind_()
    Debug.Print "Reset Test Case Flag"
    gTestCaseStarted = False
End Sub

'# Shows the current status of the test execution as well as providing a start and stop function
Public Sub Ping_()
    If Not Vase.IsRunningTest Then
        Debug.Print ""
        Debug.Print "Test Execution: " & IIf(gAssertion, "Passed", "Failed")
        If Not gAssertion Then
            Debug.Print "First Failed On: " & vbCrLf & "- " & gFirstFailed & " >> " & gFirstFailedAssertMessage & _
                IIf(gFirstFailedMessage <> "", " >> " & gFirstFailedMessage, "")
        End If
        
        gTestCaseStarted = False
    End If
End Sub

'# Assert if condition is true
Public Sub AssertTrue(Cond As Boolean, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Cond, Message:=Message, AssertName:="AssertTrue", _
        AssertFailMessage:="Got False", _
        AssertParentName:=AssertParentName
End Sub

'# Assert condition is false
Public Sub AssertFalse(Cond As Boolean, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Not Cond, Message:=Message, AssertName:="AssertFalse", _
        AssertFailMessage:="Got True", _
        AssertParentName:=AssertParentName
End Sub

'# Assert two variables are equal
Public Sub AssertEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Equal_(LeftVal, RightVal), Message:=Message, AssertName:="AssertEqual", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " <> " & ToSafeString(RightVal), _
        AssertParentName:=AssertParentName
End Sub

'# Assert left variable is like the right variable
Public Sub AssertLike(LeftVal As Variant, RightVal As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Like_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLike", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not Like " & ToSafeString(RightVal), _
        AssertParentName:=AssertParentName
End Sub

'# Assert greater than
Public Sub AssertGreaterThan(LeftVal As Variant, RightVal As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ GreaterThan_(LeftVal, RightVal), Message:=Message, AssertName:="AssertGreaterThan", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not > " & ToSafeString(RightVal), _
        AssertParentName:=AssertParentName
End Sub

'# Assert greater than or equal
Public Sub AssertGreaterThanOrEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ GreaterThanOrEqual_(LeftVal, RightVal), Message:=Message, AssertName:="AssertGreaterThanOrEqual", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not >= " & ToSafeString(RightVal), _
        AssertParentName:=AssertParentName
End Sub

'# Assert less than
Public Sub AssertLessThan(LeftVal As Variant, RightVal As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ LessThan_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLess", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not < " & ToSafeString(RightVal), _
        AssertParentName:=AssertParentName
End Sub

'# Assert less than or equal
Public Sub AssertLessThanOrEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ LessThanOrEqual_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLessThanOrEqual", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not <= " & ToSafeString(RightVal), _
        AssertParentName:=AssertParentName
End Sub

'# Assert not equal
Public Sub AssertNotEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Not Equal_(LeftVal, RightVal), Message:=Message, AssertName:="AssertNotEqual", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " = " & ToSafeString(RightVal), _
        AssertParentName:=AssertParentName
End Sub

'# Assert something is inside an array
Public Sub AssertInArray(Elem As Variant, Arr As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ VaseLib.InArray(Elem, Arr), Message:=Message, AssertName:="AssertInArray", _
        AssertFailMessage:="Got " & ToSafeString(Elem) & " Not In " & ToSafeString(Arr), _
        AssertParentName:=AssertParentName
End Sub

'# Assert array is of the correct size
Public Sub AssertArraySize(Size As Long, Arr As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Equal_(Size, UBound(Arr) - LBound(Arr) + 1), Message:=Message, AssertName:="AssertArraySize", _
        AssertFailMessage:="Got " & Size & " <> " & ToSafeArraySize(Arr), _
        AssertParentName:=AssertParentName
End Sub

'# Assert array is of the correct size
Public Sub AssertEmptyArray(Arr As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Equal_(-1, UBound(Arr)), Message:=Message, AssertName:="AssertEmptyArray", _
        AssertFailMessage:="Got " & ToSafeArraySize(Arr), _
        AssertParentName:=AssertParentName
End Sub

'# Assert array elements are equal
'# An composite assertion, this checks array size and values
Public Sub AssertArraysEqual(LeftArr As Variant, RightArr As Variant, Optional Message As String = "", Optional AssertParentName As String = "AssertEqualArrays")
On Error GoTo ErrHandler:
    Dim Tuple As Variant, ArrSize As Long
    If IsEmpty(LeftArr) Or IsEmpty(RightArr) Then
        AssertEqual LeftArr, RightArr, Message:=Message, AssertParentName:=AssertParentName
        Exit Sub
    End If
    
    Dim LeftSize As Long, RightSize As Long
    LeftSize = UBound(LeftArr) - LBound(LeftArr)
    RightSize = UBound(RightArr) - LBound(RightArr)
    AssertEqual LeftSize, RightSize, Message:=Message, AssertParentName:=AssertParentName
    If LeftSize <> RightSize Then Exit Sub
    
    Dim LeftIndex As Long, RightIndex As Long, Index As Long
    LeftIndex = LBound(LeftArr)
    RightIndex = LBound(RightArr)
    For Index = 0 To LeftSize
        AssertEqual LeftArr(LeftIndex + Index), RightArr(RightIndex + Index), Message:=Message, AssertParentName:=AssertParentName
    Next
ErrHandler:
    AssertErrorNotRaised Message:=Message, AssertParentName:=AssertParentName
    Err.Clear
End Sub

'# Assert error is not raised
Public Sub AssertErrorNotRaised(Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Err.Number = 0, Message:=Message, AssertName:="AssertEmptyArray", _
        AssertFailMessage:="Got Error# " & Err.Number, _
        AssertParentName:=AssertParentName
End Sub

'# Assert two variables are is thee same
Public Sub AssertIs(LeftVal As Variant, RightVal As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Is_(LeftVal, RightVal), Message:=Message, AssertName:="AssertIs", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " <> " & ToSafeString(RightVal), _
        AssertParentName:=AssertParentName
End Sub

'# Assert two variables are equal
Public Sub AssertIsNothing(Val As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Is_(Val, Nothing), Message:=Message, AssertName:="AssertIsNothing", _
        AssertFailMessage:="Got " & ToSafeString(Val) & " Is Something", _
        AssertParentName:=AssertParentName
End Sub

'# Assert two variables are equal
Public Sub AssertIsNotNothing(Val As Variant, Optional Message As String = "", Optional AssertParentName As String = "")
    Assert_ Not Is_(Val, Nothing), Message:=Message, AssertName:="AssertIsNotNothing", _
        AssertFailMessage:="Got " & ToSafeString(Val) & " Is Nothing", _
        AssertParentName:=AssertParentName
End Sub

'# Outputs a value to string to a test worthy output
'! Assumes Val is not an object
Private Function ToSafeString(Val As Variant) As String
On Error Resume Next
    If IsEmpty(Val) Then
        ToSafeString = "<<EMPTY>>"
    ElseIf IsArray(Val) Then
        ToSafeString = "<<Array(" & LBound(Val) & " To " & UBound(Val) & ")>>"
    ElseIf IsDate(Val) Then
        ToSafeString = "<<Date(" & Val & ")>>"
    Else
        ToSafeString = CStr(Val)
    End If
    
    If Err.Number <> 0 Then ToSafeString = "<<OBJECT>>"
    Err.Clear
End Function

'# Gets the size of an array without the hastle of error checking
Private Function ToSafeArraySize(Arr As Variant) As Long
On Error Resume Next
    ToSafeSize = -1
    ToSafeSize = UBound(Arr) - LBound(Arr) + 1
End Function

'=======================
'- Operator Functions  -
'=======================
Private Function Equal_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    Equal_ = False
    Equal_ = (LeftVal = RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function Like_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    Like_ = False
    Like_ = (LeftVal Like RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function LessThan_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    LessThan_ = False
    LessThan_ = (LeftVal < RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function LessThanOrEqual_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    LessThanOrEqual_ = False
    LessThanOrEqual_ = (LeftVal <= RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function GreaterThan_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    GreaterThan_ = False
    GreaterThan_ = (LeftVal > RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function GreaterThanOrEqual_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    GreaterThanOrEqual_ = False
    GreaterThanOrEqual_ = (LeftVal >= RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function Is_(LeftObj As Variant, RightObj As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    Is_ = False
    Is_ = (LeftVal Is RightVal)
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function



