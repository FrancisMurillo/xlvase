Attribute VB_Name = "VaseAssert"
'# This global variable determines if an assert method failed or passed
Private gAssertion As Boolean
Private gFirstFailed As String
Private gFirstFailedMessage As String
Private gFirstFailedAssertMessage As String

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
            Debug.Print "+ " & AssertName
        Else
            Debug.Print "- " & AssertName & " >> " & AssertFailMessage & _
                IIf(Message <> "", " >> " & Message, "")
        End If
    End If
    
    gAssertion = gAssertion And Cond ' Update assertion variable
    If Not Cond And gFirstFailed = "" Then ' Log the first fail condition for logging
        gFirstFailed = AssertName
        gFirstFailedMessage = Message
        gFirstFailedAssertMessage = AssertFailMessage
    End If
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
Public Sub AssertTrue(Cond As Boolean, Optional Message As String = "")
    Assert_ Cond, Message:=Message, AssertName:="AssertTrue", _
        AssertFailMessage:="Got False"
End Sub

'# Assert condition is false
Public Sub AssertFalse(Cond As Boolean, Optional Message As String = "")
    Assert_ Not Cond, Message:=Message, AssertName:="AssertFalse", _
        AssertFailMessage:="Got True"
End Sub

'# Assert two variables are equal
Public Sub AssertEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ Equal_(LeftVal, RightVal), Message:=Message, AssertName:="AssertEqual", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " <> " & ToSafeString(RightVal)
End Sub

'# Assert left variable is like the right variable
Public Sub AssertLike(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ Like_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLike", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not Like " & ToSafeString(RightVal)
End Sub

'# Assert greater than
Public Sub AssertGreaterThan(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ GreaterThan_(LeftVal, RightVal), Message:=Message, AssertName:="AssertGreaterThan", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not > " & ToSafeString(RightVal)
End Sub

'# Assert greater than or equal
Public Sub AssertGreaterThanOrEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ GreaterThanOrEqual_(LeftVal, RightVal), Message:=Message, AssertName:="AssertGreaterThanOrEqual", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not >= " & ToSafeString(RightVal)
End Sub

'# Assert less than
Public Sub AssertLessThan(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ LessThan_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLess", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not < " & ToSafeString(RightVal)
End Sub

'# Assert less than or equal
Public Sub AssertLessThanOrEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ LessThanOrEqual_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLessThanOrEqual", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " Not <= " & ToSafeString(RightVal)
End Sub

'# Assert not equal
Public Sub AssertNotEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ Not Equal_(LeftVal, RightVal), Message:=Message, AssertName:="AssertNotEqual", _
        AssertFailMessage:="Got " & ToSafeString(LeftVal) & " = " & ToSafeString(RightVal)
End Sub

'# Assert something is inside an array
Public Sub AssertInArray(Elem As Variant, Arr As Variant, Optional Message As String = "")
    Assert_ VaseLib.InArray(Elem, Arr), Message:=Message, AssertName:="AssertInArray", _
        AssertFailMessage:="Got " & ToSafeString(Elem) & " Not In " & ToSafeString(Arr)
End Sub

'# Assert array is of the correct size
Public Sub AssertArraySize(Size As Long, Arr As Variant, Optional Message As String = "")
    Assert_ Equal_(Size, UBound(Arr) + 1), Message:=Message, AssertName:="AssertArraySize", _
        AssertFailMessage:="Got " & Size & " <> " & ToSafeArraySize(Arr)
End Sub

'# Assert array is of the correct size
Public Sub AssertEmptyArray(Arr As Variant, Optional Message As String = "")
    Assert_ Equal_(-1, UBound(Arr)), Message:=Message, AssertName:="AssertEmptyArray", _
        AssertFailMessage:="Got " & ToSafeArraySize(Arr)
End Sub

'# Assert array elements are equal
Public Sub AssertEqualArrays(LeftArr As Variant, RightArr As Variant, Optional Message As String = "")
    Dim Tuple As Variant, ArrSize As Long
    ArrSize = UBound(LeftArr) + 1
    
    AssertArraySize UBound(LeftArr) + 1, RightArr, Message:=Message
    For Each Tuple In VaseLib.Zip(LeftArr, RightArr)
        AssertEqual Tuple(0), Tuple(1), Message:=Message
    Next
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

